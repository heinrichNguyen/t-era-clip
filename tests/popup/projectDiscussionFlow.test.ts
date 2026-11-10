import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import { PopupApp } from '../../src/popup/App';
import type { OneOfficeProjectInventory } from '../../src/oneoffice/types';
import { projectInventoryWarningLabel } from '../../src/shared/projectInventoryWarning';

const api = vi.hoisted(() => ({ getActiveTab: vi.fn(), collectOneOfficeProjectInventory: vi.fn(), collectOneOfficeFilteredTaskInventory: vi.fn(), collectOneOfficeDiscussion: vi.fn(), open: vi.fn() }));
vi.mock('../../src/oneoffice/collectionClient', () => ({
  ...api,
  isSupportedOneOfficeUrl: () => true,
  isOneOfficeProjectPageUrl: (url: string) => url.includes('work-project-project/view'),
  isOneOfficeFilteredTaskListUrl: (url: string) => url.includes('filtered-list'),
  isOneOfficeReportPageUrl: (url: string) => url.includes('/apps/report-report/view'),
}));
vi.mock('../../src/oneoffice/projectBatchSession', () => ({ openOneOfficeProjectBatchReview: api.open }));

const tab = { id: 3, title: 'Project test', url: 'https://office.meygroup.vn/apps/work-project-project/view?ID=870' };
const inventory = (complete = true, count = 2): OneOfficeProjectInventory => ({
  schemaVersion: 1, source: '1office-project-api', capturedAt: '2026-10-21T00:00:00Z',
  project: { type: 'project', internalId: '870', name: 'Project test', code: 'TEST', url: tab.url },
  collection: { complete, loadedTaskCount: count, expectedTaskCount: 2, warnings: [complete ? '24 duplicate task No. value(s) were found.' : 'Only one task loaded.'] },
  tasks: Array.from({ length: count }, (_, i) => ({ sourceOrder: i + 1, internalId: String(i + 1), no: '1', code: `T${i}`, name: `Task ${i}`, url: 'https://office.meygroup.vn/task', depth: 1, parentNo: null, parentInternalId: null })),
});
let container: HTMLDivElement;
let root: ReturnType<typeof createRoot>;
let listeners: Array<(value: Record<string, unknown>, area: string) => void>;
beforeEach(async () => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
  for (const mock of Object.values(api)) mock.mockReset();
  api.getActiveTab.mockResolvedValue(tab); api.collectOneOfficeProjectInventory.mockResolvedValue(inventory()); api.open.mockResolvedValue(undefined);
  listeners = [];
  vi.stubGlobal('chrome', { runtime: {}, storage: {
    local: { get: (_: unknown, done: (data: unknown) => void) => done({ teraClipItems: [] }) },
    onChanged: { addListener: (listener: typeof listeners[number]) => listeners.push(listener), removeListener: (listener: typeof listeners[number]) => { listeners = listeners.filter((r) => r !== listener); } },
  } });
  container = document.createElement('div'); document.body.append(container); root = createRoot(container);
  await act(async () => { root.render(React.createElement(PopupApp)); });
});
afterEach(async () => { await act(async () => { root.unmount(); }); container.remove(); vi.unstubAllGlobals(); });
const button = (text: string) => [...container.querySelectorAll('button')].find((b) => b.textContent?.trim() === text)!;
const click = async (text: string) => { expect(button(text)).toBeDefined(); await act(async () => { button(text).click(); }); };

it('keeps inventory inspection before collection but completes the normal project launch with one click', async () => {
  expect(button('Thu thập thảo luận toàn dự án')).toBeDefined();
  expect(container.textContent).not.toContain('Danh sách công việc dự án');
  await click('Thu thập thảo luận toàn dự án');
  expect(api.collectOneOfficeProjectInventory).toHaveBeenCalledExactlyOnceWith(3, 'https://office.meygroup.vn');
  expect(api.open).toHaveBeenCalledExactlyOnceWith(tab, inventory());
  expect(api.collectOneOfficeProjectInventory.mock.invocationCallOrder[0]).toBeLessThan(api.open.mock.invocationCallOrder[0]);
  expect(api.collectOneOfficeDiscussion).not.toHaveBeenCalled();
  expect(api.collectOneOfficeFilteredTaskInventory).not.toHaveBeenCalled();
  expect(container.textContent).toContain('Có 24 giá trị No. bị trùng.');
  expect(container.querySelectorAll('section')).toHaveLength(1);
});

it('stops on partial inventories until the user explicitly continues without recollecting', async () => {
  const partial = inventory(false, 1); api.collectOneOfficeProjectInventory.mockResolvedValue(partial);
  await click('Thu thập thảo luận toàn dự án');
  expect(api.open).not.toHaveBeenCalled();
  expect(container.textContent).toContain('Danh sách chưa đầy đủ');
  expect(container.textContent).toContain('Only one task loaded.');
  await click('Tiếp tục với 1 công việc đã lấy');
  expect(api.open).toHaveBeenCalledExactlyOnceWith(tab, partial);
  expect(api.collectOneOfficeProjectInventory).toHaveBeenCalledTimes(1);
});

it('retains check-only for the normalization workspace without starting discussions', async () => {
  await click('Chỉ kiểm tra danh sách công việc');
  expect(api.open).not.toHaveBeenCalled();
  expect(container.textContent).toContain('2 công việc');
  await click('Thu thập thảo luận toàn dự án');
  expect(api.open).toHaveBeenCalledTimes(1); expect(api.collectOneOfficeProjectInventory).toHaveBeenCalledTimes(1);
});

it('does not proceed after inspection failure or an empty inventory and permits retry', async () => {
  api.collectOneOfficeProjectInventory.mockRejectedValueOnce(new Error('Source unavailable'));
  await click('Thu thập thảo luận toàn dự án');
  expect(api.open).not.toHaveBeenCalled(); expect(container.textContent).toContain('Source unavailable');
  api.collectOneOfficeProjectInventory.mockResolvedValueOnce(inventory(true, 0));
  await click('Thu thập thảo luận toàn dự án');
  expect(api.open).not.toHaveBeenCalled(); expect(container.textContent).toContain('Chưa có công việc');
  await click('Thu thập thảo luận toàn dự án');
  expect(api.open).toHaveBeenCalledTimes(1);
});

it('retries opening with the same checked inventory rather than sending a second inventory request', async () => {
  api.open.mockRejectedValueOnce(new Error('Cannot open collection page'));
  await click('Thu thập thảo luận toàn dự án');
  expect(container.textContent).toContain('Cannot open collection page');
  await click('Thu thập thảo luận toàn dự án');
  expect(api.open).toHaveBeenCalledTimes(2); expect(api.collectOneOfficeProjectInventory).toHaveBeenCalledTimes(1);
});

it('prevents duplicate and competing actions while inventory inspection is in progress', async () => {
  let resolve!: (value: OneOfficeProjectInventory) => void;
  api.collectOneOfficeProjectInventory.mockImplementation(() => new Promise((done) => { resolve = done; }));
  await click('Thu thập thảo luận toàn dự án');
  expect(button('Đang kiểm tra danh sách...').disabled).toBe(true);
  expect(button('Chỉ kiểm tra danh sách công việc').disabled).toBe(true);
  expect(button('Chỉ lấy thảo luận trang hiện tại').disabled).toBe(true);
  await click('Đang kiểm tra danh sách...');
  expect(api.collectOneOfficeProjectInventory).toHaveBeenCalledTimes(1);
  expect(api.open).not.toHaveBeenCalled();
  await act(async () => { resolve(inventory()); });
  expect(api.open).toHaveBeenCalledTimes(1);
});

it('does not open collection using results returned after temporary data was cleared', async () => {
  let resolve!: (value: OneOfficeProjectInventory) => void;
  api.collectOneOfficeProjectInventory.mockImplementation(() => new Promise((done) => { resolve = done; }));
  await click('Thu thập thảo luận toàn dự án');
  await act(async () => { listeners.forEach((listener) => listener({ oneOfficeTemporaryClearRevision: {} }, 'session')); });
  await act(async () => { resolve(inventory()); });
  expect(api.open).not.toHaveBeenCalled(); expect(container.textContent).not.toContain('2 công việc');
});

it('keeps the existing two-action filtered-list flow and its collector unchanged', async () => {
  await act(async () => { root.unmount(); }); root = createRoot(container);
  api.getActiveTab.mockResolvedValue({ ...tab, url: 'https://office.meygroup.vn/filtered-list' });
  api.collectOneOfficeFilteredTaskInventory.mockResolvedValue(inventory());
  await act(async () => { root.render(React.createElement(PopupApp)); });
  expect(button('Thu thập thảo luận toàn dự án')).toBeUndefined();
  await click('Kiểm tra danh sách công việc đang lọc');
  expect(api.open).not.toHaveBeenCalled(); expect(api.collectOneOfficeProjectInventory).not.toHaveBeenCalled();
  await click('Thu thập thảo luận danh sách này'); expect(api.open).toHaveBeenCalledTimes(1);
});

it('translates only the known duplicate-No warning without suppressing other warnings', () => {
  expect(projectInventoryWarningLabel('24 duplicate task No. value(s) were found.')).toBe('Có 24 giá trị No. bị trùng.');
  expect(projectInventoryWarningLabel('Permission denied')).toBe('Permission denied');
});
