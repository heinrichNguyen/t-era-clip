import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { afterEach, expect, it, vi } from 'vitest';
import { ProjectCodeNoReconciliation } from '../../src/import-review/ProjectCodeNoReconciliation';
import { reviewProjectImport } from '../../src/oneoffice/projectImportReview';
import { collectTestTaskDetails } from '../../src/oneoffice/testTaskDetails';
import type { OneOfficeProjectInventory } from '../../src/oneoffice/types';
vi.mock('../../src/oneoffice/testTaskDetails', () => ({ collectTestTaskDetails: vi.fn() }));
afterEach(() => { vi.resetAllMocks(); vi.unstubAllGlobals(); document.body.innerHTML = ''; });
const origin = 'https://testoffice.meygroup.vn:9443';
const data = (count = 1): OneOfficeProjectInventory => ({ schemaVersion: 1, source: '1office-project-api', capturedAt: 'snapshot',
  project: { type: 'project', internalId: '203', name: 'Test project', code: null, url: `${origin}/apps/work-project-project/view?ID=203` },
  collection: { complete: true, expectedTaskCount: count, loadedTaskCount: count, warnings: [] },
  tasks: Array.from({ length: count }, (_, i) => ({ internalId: String(i + 1), name: 'Existing', code: 'OLD', no: 'outline', sourceNo: '1', parentInternalId: null, parentNo: null, depth: 1, sourceOrder: i, url: `${origin}/apps/work-task-task/view?ID=${i + 1}` })),
});
const report = reviewProjectImport({ sheet: 'Approved', data: [['No.', 'Mã gói thầu/Mã CV', 'Nội dung công việc', 'ID'], ['2', 'NEW', '<img src=x onerror=alert(1)>', '1']] }, null);
async function mount(inventory = data()) {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true); const fetch = vi.fn(); vi.stubGlobal('fetch', fetch);
  const container = document.createElement('div'); document.body.append(container); const root = createRoot(container); const busy = vi.fn();
  await act(async () => root.render(React.createElement(ProjectCodeNoReconciliation, { report, inventory, disabled: false, sourceTabId: 42, onBusyChange: busy })));
  const button = (text: string) => [...container.querySelectorAll<HTMLButtonElement>('button')].find((item) => item.textContent?.startsWith(text))!;
  const change = async (selector: string, value: string) => { await act(async () => {
    const input = container.querySelector<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>(selector)!;
    const proto = input.tagName === 'SELECT' ? HTMLSelectElement.prototype : input.tagName === 'TEXTAREA' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
    Object.getOwnPropertyDescriptor(proto, 'value')!.set!.call(input, value);
    input.dispatchEvent(new Event(input.tagName === 'SELECT' ? 'change' : 'input', { bubbles: true }));
  }); };
  return { container, root, busy, fetch, button, change };
}
it('requires confirmations, maps by ID, shows code and No. together, and invalidates the preview on edits', async () => {
  const view = await mount();
  try {
    expect(collectTestTaskDetails).not.toHaveBeenCalled();
    await view.change('select[aria-label="Cột ID 1Office (nếu có)"]', 'ID');
    await view.change('select[aria-label="Ý nghĩa trường No."]', 'custom-no');
    await act(async () => { view.container.querySelectorAll<HTMLInputElement>('input[type=checkbox]').forEach((input) => input.click()); });
    await act(async () => view.button('Xem bảng').click());
    expect(view.container.textContent).toContain('OLD → NEW'); expect(view.container.textContent).toContain('1 → 2');
    expect(view.container.textContent).toContain('Phương án gồm 2 bước');
    expect(view.container.querySelector('img')).toBeNull();
    expect(view.container.textContent).toContain('<img src=x onerror=alert(1)>');
    expect(view.container.textContent).toContain('Chưa');
    await view.change('textarea', '2 1');
    expect(view.container.querySelector('table')).toBeNull();
    expect(view.fetch).not.toHaveBeenCalled(); expect(collectTestTaskDetails).not.toHaveBeenCalled();
  } finally { await act(async () => view.root.unmount()); }
});
it('reads at most 20 missing task details only on request, without discussion collection or writes', async () => {
  const inventory = data(25); inventory.tasks.forEach((task) => { task.code = null; task.sourceNo = undefined; });
  vi.mocked(collectTestTaskDetails).mockImplementation(async (_tab, id) => ({ taskId: id, sourceUrl: `${origin}/apps/work-task-task/view?ID=${id}`, capturedAt: 'now', identityVerified: true,
    evidenceSource: 'widget-response', code: { state: 'value', displayValue: `CODE.${id}` }, no: { state: 'empty', displayValue: '' }, error: null }));
  const view = await mount(inventory);
  try {
    await act(async () => { view.button('Đọc hiện trạng').click(); view.button('Đọc hiện trạng').click(); });
    expect(collectTestTaskDetails).toHaveBeenCalledTimes(20);
    expect(view.busy.mock.calls).toEqual([[true], [false]]); expect(view.fetch).not.toHaveBeenCalled();
    expect(inventory.tasks[0].code).toBeNull();
  } finally { await act(async () => view.root.unmount()); }
});
it('does not offer supplemental reads for production or unsupported project context', async () => {
  const inventory = data(); inventory.project.url = 'https://office.meygroup.vn/apps/work-project-project/view?ID=203';
  const view = await mount(inventory);
  try { expect(view.button('Đọc hiện trạng')).toBeUndefined(); expect(collectTestTaskDetails).not.toHaveBeenCalled(); }
  finally { await act(async () => view.root.unmount()); }
});
it('discards late read results after closing the workflow', async () => {
  const inventory = data(2); inventory.tasks.forEach((task) => { task.code = null; task.sourceNo = undefined; });
  let resolve!: (value: any) => void; vi.mocked(collectTestTaskDetails).mockReturnValue(new Promise((done) => { resolve = done; }));
  const view = await mount(inventory);
  await act(async () => view.button('Đọc hiện trạng').click());
  await act(async () => view.root.unmount());
  await act(async () => resolve({ identityVerified: false }));
  expect(collectTestTaskDetails).toHaveBeenCalledTimes(1); expect(view.container.textContent).toBe('');
});
