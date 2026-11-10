import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { afterEach, expect, it, vi } from 'vitest';
import { ProjectImportNormalization } from '../../src/no-review/ProjectImportNormalization';
import { reviewProjectImport } from '../../src/oneoffice/projectImportReview';
import { parseProjectErpCatalog } from '../../src/oneoffice/projectErpCatalog';

afterEach(() => { vi.useRealTimers(); vi.unstubAllGlobals(); document.body.innerHTML = ''; });
const catalog = () => parseProjectErpCatalog({ sheet: 'TDR', data: [
  ['2. Danh mục gói thầu'], ['Mã', 'Tên'], ['TV.001', '<img src=x> Tư vấn'], ['TV.002', 'Tư vấn'],
  ['17. Danh mục hạng mục'], ['Mã', 'Tên'], ['P', 'Project'],
] });
async function mount() {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
  const fetch = vi.fn(); vi.stubGlobal('fetch', fetch);
  const executeScript = vi.fn(); vi.stubGlobal('chrome', { scripting: { executeScript } });
  class FakeWorker {
    onmessage: ((event: { data: unknown }) => void) | null = null;
    onerror: (() => void) | null = null;
    terminate = vi.fn(); postMessage = vi.fn();
    constructor() { instances.push(this); }
  }
  const instances: FakeWorker[] = [];
  vi.stubGlobal('Worker', FakeWorker);
  const report = reviewProjectImport({ sheet: 'S', data: [
    ['No.', 'Mã gói thầu/Mã CV', 'Nội dung công việc'], ['1.1.1', 'ITEM', 'Item'], ['1.1.1.1', null, 'Tư vấn'], ['1.1.1.1.1', null, 'Child'],
  ] }, null);
  const before = JSON.stringify(report);
  const container = document.createElement('div'); document.body.append(container);
  const root = createRoot(container);
  await act(async () => { root.render(React.createElement(ProjectImportNormalization, { report, inventory: null, disabled: false })); });
  const setInput = async (id: string, value: string) => act(async () => {
    const input = container.querySelector<HTMLInputElement>(`#${id}`)!;
    Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!.call(input, value);
    input.dispatchEvent(new Event('input', { bubbles: true }));
  });
  const setSelect = async (id: string, value: string) => act(async () => {
    const select = container.querySelector<HTMLSelectElement>(`#${id}`)!;
    select.value = value; select.dispatchEvent(new Event('change', { bubbles: true }));
  });
  await setInput('normalization-prefix', 'P'); await setSelect('normalization-mode', 'initial');
  const load = async (name = 'catalog.xlsx', size = 7 * 1024 * 1024) => {
    const input = container.querySelector<HTMLInputElement>('#erp-catalog-file')!;
    Object.defineProperty(input, 'files', { configurable: true, value: [{ name, size, arrayBuffer: async () => new ArrayBuffer(8) }] });
    await act(async () => { input.dispatchEvent(new Event('change', { bubbles: true })); });
  };
  const preview = () => act(async () => { [...container.querySelectorAll('button')].find((b) => b.textContent === 'Tạo bảng mã xem trước')!.click(); });
  return { root, container, instances, setInput, setSelect, load, preview, fetch, executeScript, report, before };
}

it('loads one sheet locally, requires explicit ERP choice, propagates it, invalidates previews and clears selections', async () => {
  const view = await mount();
  try {
    await view.preview(); expect(view.container.textContent).toContain('Cần kiểm tra');
    await view.load();
    expect(view.container.querySelector('table')).toBeNull();
    expect(view.instances[0].postMessage.mock.calls[0][0].sheet).toBe('TDR');
    await act(async () => { view.instances[0].onmessage?.({ data: { catalog: catalog() } }); });
    expect(view.container.textContent).toContain('1 mã dự án, 2 mã gói thầu');
    expect(view.container.querySelector<HTMLSelectElement>('#erp-pick-3')!.value).toBe('');
    expect(view.container.querySelector('img')).toBeNull();
    await view.setSelect('erp-pick-3', 'TV.002');
    await view.preview();
    expect(view.container.textContent).toContain('P.ITEM.TV.002.01');
    await view.setSelect('erp-pick-3', 'TV.001');
    expect(view.container.querySelector('table')).toBeNull();
    await view.preview(); expect(view.container.textContent).toContain('P.ITEM.TV.001.01');
    await act(async () => { [...view.container.querySelectorAll('button')].find((b) => b.textContent === 'Bỏ danh mục ERP và các lựa chọn')!.click(); });
    expect(view.container.querySelector('#erp-pick-3')).toBeNull();
    await view.preview(); expect(view.container.textContent).not.toContain('P.ITEM.TV.001');
    expect(JSON.stringify(view.report)).toBe(view.before);
    expect(view.fetch).not.toHaveBeenCalled(); expect(view.executeScript).not.toHaveBeenCalled();
  } finally { await act(async () => { view.root.unmount(); }); }
});

it('rejects oversized/wrong files and reports worker errors without retaining a previous catalog', async () => {
  const view = await mount();
  try {
    await view.load('catalog.xls'); await view.load('large.xlsx', 11 * 1024 * 1024);
    expect(view.instances).toHaveLength(0); expect(view.container.textContent).toContain('tối đa 10 MB');
    await view.load(); await act(async () => { view.instances[0].onmessage?.({ data: { catalog: catalog() } }); });
    await view.setSelect('erp-pick-3', 'TV.001');
    await view.load('replacement.xlsx');
    expect(view.container.querySelector('#erp-pick-3')).toBeNull();
    await act(async () => { view.instances[1].onmessage?.({ data: { error: 'Không tìm thấy sheet' } }); });
    expect(view.container.textContent).toContain('Không tìm thấy sheet');
    await view.preview(); expect(view.container.textContent).not.toContain('P.ITEM.TV.001');
  } finally { await act(async () => { view.root.unmount(); }); }
});

it('cancels reads on timeout, sheet changes and unmount, ignoring late results', async () => {
  vi.useFakeTimers();
  const view = await mount();
  try {
    await view.load(); const late = view.instances[0].onmessage;
    await act(async () => { vi.advanceTimersByTime(30000); });
    expect(view.instances[0].terminate).toHaveBeenCalled();
    await act(async () => { late?.({ data: { catalog: catalog() } }); });
    expect(view.container.textContent).toContain('30 giây'); expect(view.container.querySelector('#erp-pick-3')).toBeNull();
    await view.load(); const late2 = view.instances[1].onmessage;
    await view.setInput('erp-catalog-sheet', 'Other');
    expect(view.instances[1].terminate).toHaveBeenCalled();
    await act(async () => { late2?.({ data: { catalog: catalog() } }); });
    expect(view.container.querySelector('#erp-pick-3')).toBeNull();
    await view.load();
  } finally { await act(async () => { view.root.unmount(); }); }
  expect(view.instances[2].terminate).toHaveBeenCalled();
});

it('disables ERP remapping of old rows in update mode and resets choices when levels/mode change', async () => {
  const view = await mount();
  try {
    await view.load(); await act(async () => { view.instances[0].onmessage?.({ data: { catalog: catalog() } }); });
    await view.setSelect('erp-pick-3', 'TV.001');
    await view.setSelect('normalization-mode', 'update');
    const picker = view.container.querySelector<HTMLSelectElement>('#erp-pick-3')!;
    expect(picker.disabled).toBe(true); expect(picker.value).toBe('');
    await view.setSelect('normalization-mode', 'initial');
    expect(view.container.querySelector<HTMLSelectElement>('#erp-pick-3')!.value).toBe('');
  } finally { await act(async () => { view.root.unmount(); }); }
});

it('labels same-description codes as pending confirmation without assigning one or blocking other rows', async () => {
  const view = await mount();
  try {
    const duplicate = catalog(); duplicate.packages[0].name = duplicate.packages[1].name;
    await view.load(); await act(async () => { view.instances[0].onmessage?.({ data: { catalog: duplicate } }); });
    const picker = view.container.querySelector<HTMLSelectElement>('#erp-pick-3')!;
    expect(picker.value).toBe('');
    expect([...picker.options].filter((o) => o.textContent?.includes('2 mã cùng mô tả'))).toHaveLength(2);
    expect(view.container.textContent).toContain('chờ xác nhận với kế toán trưởng');
    await view.preview();
    expect(view.container.textContent).toContain('Có đề xuất mã mới1');
    expect(view.container.textContent).not.toContain('P.ITEM.TV.001');
    expect(view.container.textContent).not.toContain('P.ITEM.TV.002');
    expect(view.container.textContent).toContain('Cần kiểm tra2');
    await view.setSelect('erp-pick-3', 'TV.002');
    expect(view.container.textContent).toContain('Chỉ giữ lựa chọn sau khi đã xác minh');
    expect(view.fetch).not.toHaveBeenCalled(); expect(view.executeScript).not.toHaveBeenCalled();
  } finally { await act(async () => { view.root.unmount(); }); }
});
