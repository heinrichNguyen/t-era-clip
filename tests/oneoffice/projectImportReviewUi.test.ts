import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { afterEach, expect, it, vi } from 'vitest';
import { ProjectImportReview } from '../../src/no-review/ProjectImportReview';

afterEach(() => { vi.useRealTimers(); vi.unstubAllGlobals(); document.body.innerHTML = ''; });

async function mount() {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
  const fetch = vi.fn(); vi.stubGlobal('fetch', fetch);
  const instances: FakeWorker[] = [];
  class FakeWorker {
    onmessage: ((e: unknown) => void) | null = null;
    onerror: (() => void) | null = null;
    terminate = vi.fn();
    postMessage = vi.fn();
    constructor() { instances.push(this); }
  }
  vi.stubGlobal('Worker', FakeWorker);
  const container = document.createElement('div'); document.body.append(container);
  const root = createRoot(container);
  await act(async () => { root.render(React.createElement(ProjectImportReview, { inventory: null, disabled: false })); });
  const select = async (name = 'test.xlsx', size = 12) => {
    const file = { name, size, arrayBuffer: async () => new ArrayBuffer(12) };
    const input = container.querySelector('input[type=file]')!;
    Object.defineProperty(input, 'files', { configurable: true, value: [file] });
    await act(async () => { input.dispatchEvent(new Event('change', { bubbles: true })); });
  };
  return { container, root, select, instances, fetch };
}

it('renders local results as text, supports sheet selection and clears data without network requests', async () => {
  const view = await mount();
  try {
    await view.select();
    expect(view.instances).toHaveLength(1);
    await act(async () => { view.instances[0].onmessage?.({ data: { sheets: [
      { sheet: 'Tasks', data: [['No.', 'Mã gói thầu/Mã CV', 'Nội dung công việc'], ['1', 'CODE', '<img src=x onerror=alert(1)>']] },
      { sheet: 'Other', data: [['wrong']] },
    ] } }); });
    expect(view.container.textContent).toContain('1 dòng');
    expect(view.container.textContent).toContain('<img src=x onerror=alert(1)>');
    expect(view.container.querySelector('img')).toBeNull();
    expect(view.container.textContent).toContain('Chưa đối chiếu 1Office');
    const picker = view.container.querySelector('select')!;
    await act(async () => { picker.value = '1'; picker.dispatchEvent(new Event('change', { bubbles: true })); });
    expect(view.container.textContent).toContain('Thiếu cột');
    await act(async () => { view.container.querySelector('button')!.click(); });
    expect(view.container.textContent).not.toContain('test.xlsx');
    expect(view.container.querySelector('table')).toBeNull();
    expect(view.fetch).not.toHaveBeenCalled();
  } finally { await act(async () => { view.root.unmount(); }); }
});

it('rejects oversized or unsupported files before parsing', async () => {
  const view = await mount();
  try {
    await view.select('file.xls');
    expect(view.container.textContent).toContain('Chỉ nhận file .xlsx tối đa 5 MB');
    await view.select('file.xlsx', 6 * 1024 * 1024);
    expect(view.instances).toHaveLength(0);
  } finally { await act(async () => { view.root.unmount(); }); }
});

it('terminates timed-out parsing and does not display late results after clearing', async () => {
  vi.useFakeTimers();
  const view = await mount();
  try {
    await view.select();
    await act(async () => { vi.advanceTimersByTime(30000); });
    expect(view.instances[0].terminate).toHaveBeenCalled();
    expect(view.container.textContent).toContain('30 giây');
    await view.select();
    const late = view.instances[1].onmessage;
    await act(async () => { view.container.querySelector('button')!.click(); });
    await act(async () => { late?.({ data: { sheets: [{ sheet: 'Late', data: [] }] } }); });
    expect(view.container.querySelector('select')).toBeNull();
    expect(view.instances[1].terminate).toHaveBeenCalled();
  } finally { await act(async () => { view.root.unmount(); }); }
});
