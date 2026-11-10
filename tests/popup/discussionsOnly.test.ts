import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { afterEach, expect, it, vi } from 'vitest';
import { DiscussionPopup } from '../../src/discussions/Popup';
import { collectOneOfficeDiscussion, collectOneOfficeProjectInventory, getActiveTab } from '../../src/oneoffice/collectionClient';
import { openOneOfficeReview } from '../../src/oneoffice/reviewSession';
import { openOneOfficeProjectBatchReview } from '../../src/oneoffice/projectBatchSession';
import { collectOneOfficeReportInventory } from '../../src/oneoffice/reportTaskInventory';
vi.mock('../../src/oneoffice/collectionClient', () => ({ collectOneOfficeDiscussion: vi.fn(), collectOneOfficeProjectInventory: vi.fn(), getActiveTab: vi.fn(),
  isOneOfficeReportPageUrl: (url: string) => url.includes('/apps/report-report/view'),
  isSupportedOneOfficeUrl: (url: string) => url.startsWith('https://office.meygroup.vn/'), isOneOfficeProjectPageUrl: (url: string) => url.includes('/work-project-project/view') }));
vi.mock('../../src/oneoffice/reviewSession', () => ({ openOneOfficeReview: vi.fn() }));
vi.mock('../../src/oneoffice/projectBatchSession', () => ({ openOneOfficeProjectBatchReview: vi.fn() }));
vi.mock('../../src/oneoffice/reportTaskInventory', () => ({ collectOneOfficeReportInventory: vi.fn() }));
afterEach(() => { vi.resetAllMocks(); vi.unstubAllGlobals(); document.body.innerHTML = ''; });
const taskTab = { id: 7, title: 'Task', url: 'https://office.meygroup.vn/apps/work-task-task/view?ID=1' };
const projectTab = { ...taskTab, url: 'https://office.meygroup.vn/apps/work-project-project/view?ID=1' };
async function mount(tab = taskTab) {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true); vi.mocked(getActiveTab).mockResolvedValue(tab);
  const div = document.createElement('div'); document.body.append(div); const root = createRoot(div);
  await act(async () => root.render(React.createElement(DiscussionPopup)));
  const button = (prefix: string) => [...div.querySelectorAll('button')].find((item) => item.textContent?.startsWith(prefix))!;
  return { root, div, button };
}
it('has only discussion actions, waits for a click and disables project collection on task pages', async () => {
  const view = await mount();
  try {
    expect(view.div.textContent).not.toMatch(/Capture|Chuẩn hóa|lịch sử truy cập|nhập dữ liệu|đổi mã/i);
    expect(view.button('Thu thập thảo luận toàn dự án').disabled).toBe(true);
    expect(collectOneOfficeDiscussion).not.toHaveBeenCalled(); expect(collectOneOfficeProjectInventory).not.toHaveBeenCalled();
  } finally { await act(async () => view.root.unmount()); }
});
it('collects a task once and opens its existing review only on explicit request', async () => {
  const result = { collection: { loadedRootCommentCount: 1, loadedReplyCount: 2, complete: true, warnings: [] } } as any;
  vi.mocked(collectOneOfficeDiscussion).mockResolvedValue(result);
  const view = await mount();
  try {
    await act(async () => { view.button('Thu thập thảo luận trang').click(); view.button('Thu thập thảo luận trang').click(); });
    expect(collectOneOfficeDiscussion).toHaveBeenCalledTimes(1); expect(openOneOfficeReview).not.toHaveBeenCalled();
    await act(async () => view.button('Xem trước').click());
    expect(openOneOfficeReview).toHaveBeenCalledExactlyOnceWith(result);
  } finally { await act(async () => view.root.unmount()); }
});
it('checks the project inventory before opening batch collection and refuses partial inventories', async () => {
  const inventory = { tasks: [{}], collection: { complete: false, warnings: [] } } as any;
  vi.mocked(collectOneOfficeProjectInventory).mockResolvedValue(inventory);
  const view = await mount(projectTab);
  try {
    await act(async () => view.button('Thu thập thảo luận toàn').click()); expect(openOneOfficeProjectBatchReview).not.toHaveBeenCalled();
    inventory.collection.complete = true;
    await act(async () => view.button('Thu thập thảo luận toàn').click());
    expect(openOneOfficeProjectBatchReview).toHaveBeenCalledExactlyOnceWith(projectTab, inventory);
  } finally { await act(async () => view.root.unmount()); }
});
it('rejects unsupported origins and suppresses a late result after unmount', async () => {
  const view = await mount({ ...taskTab, url: 'https://example.com/' });
  expect(view.button('Thu thập thảo luận trang').disabled).toBe(true);
  await act(async () => view.root.unmount());
  let done!: (value: any) => void; vi.mocked(collectOneOfficeDiscussion).mockReturnValue(new Promise((resolve) => { done = resolve; }));
  const next = await mount();
  await act(async () => next.button('Thu thập thảo luận trang').click());
  await act(async () => next.root.unmount());
  await act(async () => done({ collection: { complete: true } }));
  expect(openOneOfficeReview).not.toHaveBeenCalled();
});

it('collects the report through the same batch review without collecting the entire project', async () => {
  const tab = { id: 7, title: 'Báo cáo', url: 'https://office.meygroup.vn/apps/report-report/view?ID=74&department_ids[]=424' };
  const inventory = { tasks: [{}], collection: { complete: true, warnings: [] } } as any;
  vi.mocked(collectOneOfficeReportInventory).mockResolvedValue(inventory);
  const view = await mount(tab);
  try {
    expect(view.button('Thu thập thảo luận trang').disabled).toBe(true);
    await act(async () => view.button('Thu thập thảo luận danh sách trong báo cáo').click());
    expect(collectOneOfficeReportInventory).toHaveBeenCalledExactlyOnceWith(7, 'https://office.meygroup.vn', tab.url);
    expect(openOneOfficeProjectBatchReview).toHaveBeenCalledExactlyOnceWith(tab, inventory);
    expect(collectOneOfficeProjectInventory).not.toHaveBeenCalled();
  } finally { await act(async () => view.root.unmount()); }
});
