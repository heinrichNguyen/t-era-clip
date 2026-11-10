import { afterEach, describe, expect, it, vi } from 'vitest';
import { buildOneOfficeProjectWorkbook } from '../../src/oneoffice/export/projectXlsx';
import type { OneOfficeProjectInventory, OneOfficeProjectTaskInventoryItem } from '../../src/oneoffice/types';

const api = vi.hoisted(() => ({ comments: vi.fn(), fast: vi.fn(), resolve: vi.fn() }));
vi.mock('../../src/oneoffice/reportTaskLookup', () => ({ resolveReportTask: api.resolve }));
vi.mock('../../src/oneoffice/projectCommentApi', () => ({ collectOneOfficeTaskFromCommentApi: api.comments }));
vi.mock('../../src/oneoffice/projectFastPathProbe', () => ({ collectOneOfficeTaskFastPath: api.fast }));
const origin = 'https://office.meygroup.vn';
const unresolved = (index: number): OneOfficeProjectTaskInventoryItem => ({
  sourceOrder: index, internalId: '', collectionKey: `report-row:${index}`, identityError: 'Chưa lấy được thảo luận: chưa xác minh ID.',
  no: '1', code: null, name: `Công việc ${index}`, url: '', depth: 0, parentNo: null, parentInternalId: null,
});
const inventory = (tasks: OneOfficeProjectTaskInventoryItem[]): OneOfficeProjectInventory => ({
  schemaVersion: 1, source: '1office-report-dom', capturedAt: '2026-11-05T00:00:00Z',
  project: { type: 'report', internalId: '74', code: null, name: 'Báo cáo đã lọc', url: `${origin}/apps/report-report/view?ID=74&department_ids[]=424` },
  collection: { complete: true, expectedTaskCount: tasks.length, loadedTaskCount: tasks.length,
    warnings: ['Có công việc chưa xác minh ID; giữ toàn bộ danh sách.'] }, tasks,
});
afterEach(() => { vi.unstubAllGlobals(); vi.clearAllMocks(); });

describe('unresolved report rows', () => {
  it('refills each free worker without waiting for slower peers, with at most three requests and stable output order', async () => {
    const tasks = Array.from({length: 6}, (_, index) => ({...unresolved(index), identityError: undefined,
      internalId: String(89876 + index), url: `${origin}/apps/work-normal-normal/view?ID=${89876 + index}`}));
    const source = inventory(tasks);
    vi.stubGlobal('chrome', {tabs: {get: async () => ({url: source.project.url})}});
    const releases = new Map<string, () => void>();
    let active = 0; let peak = 0;
    api.comments.mockImplementation((_tab, task) => new Promise(resolve => {
      active++; peak = Math.max(peak, active);
      releases.set(task.internalId, () => {
        active--;
        resolve({schemaVersion: 1, source: '1office-comment-api', entity: {type: 'task', internalId: task.internalId},
          comments: [], collection: {complete: true, loadedRootCommentCount: 0, loadedReplyCount: 0, warnings: []}});
      });
    }));
    const {collectOneOfficeProjectTaskDiscussions} = await import('../../src/oneoffice/projectBatchCollector');
    const running = collectOneOfficeProjectTaskDiscussions(5, source, () => {}, () => false);
    await vi.waitFor(() => expect(releases.size).toBe(3));
    releases.get(tasks[1].internalId)!();
    await vi.waitFor(() => expect(releases.size).toBe(4));
    expect(active).toBe(3);
    releases.get(tasks[3].internalId)!();
    await vi.waitFor(() => expect(releases.size).toBe(5));
    releases.get(tasks[4].internalId)!();
    await vi.waitFor(() => expect(releases.size).toBe(6));
    for (const index of [5, 2, 0]) releases.get(tasks[index].internalId)!();
    const result = await running;
    expect(peak).toBe(3);
    expect(result.tasks.map(row => row.task.internalId)).toEqual(tasks.map(task => task.internalId));
    expect(result.tasks.every(row => row.state === 'complete')).toBe(true);
    expect(api.resolve).not.toHaveBeenCalled();
    expect(api.fast).not.toHaveBeenCalled();
  });
  it('stops scheduling new report rows on cancellation and preserves all rows', async () => {
    const tasks = Array.from({length: 5}, (_, index) => ({...unresolved(index), identityError: undefined,
      internalId: String(89876 + index), url: `${origin}/apps/work-normal-normal/view?ID=${89876 + index}`}));
    const source = inventory(tasks);
    vi.stubGlobal('chrome', {tabs: {get: async () => ({url: source.project.url})}});
    let cancelled = false;
    api.comments.mockImplementation(async () => {cancelled = true; return null;});
    const {collectOneOfficeProjectTaskDiscussions} = await import('../../src/oneoffice/projectBatchCollector');
    const result = await collectOneOfficeProjectTaskDiscussions(5, source, () => {}, () => cancelled);
    expect(api.comments).toHaveBeenCalledTimes(1);
    expect(result.tasks).toHaveLength(5);
    expect(result.tasks.every(row => row.state === 'cancelled')).toBe(true);
  });
  it('resolves IDs in the review worker and never uses navigation fallback for report discussions', async () => {
    const pending = {...unresolved(0), code: 'CODE', identityError: undefined};
    const source = inventory([pending]);
    const update = vi.fn();
    vi.stubGlobal('chrome', { tabs: {get: async () => ({url: source.project.url}), update} });
    api.resolve.mockResolvedValue({...pending, internalId: '89876', url: `${origin}/apps/work-normal-normal/view?ID=89876`,
      identityWarning: 'Giữ mã báo cáo, đối chiếu dạng mã chi tiết đã xác minh.'});
    api.comments.mockResolvedValue(null); api.fast.mockResolvedValue(null);
    const { collectOneOfficeProjectTaskDiscussions } = await import('../../src/oneoffice/projectBatchCollector');
    const result = await collectOneOfficeProjectTaskDiscussions(5, source, () => {}, () => false);
    expect(api.resolve).toHaveBeenCalledExactlyOnceWith(5, source.project.url, pending);
    expect(result.tasks[0].state).toBe('failed');
    expect(result.tasks[0].error).toContain('không mở trang chi tiết');
    expect(result.inventory.tasks[0].internalId).toBe('89876');
    expect(result.warnings).toContain('CODE: Giữ mã báo cáo, đối chiếu dạng mã chi tiết đã xác minh.');
    expect(JSON.stringify(buildOneOfficeProjectWorkbook(result))).toContain('Giữ mã báo cáo, đối chiếu dạng mã chi tiết đã xác minh.');
    expect(source.tasks[0].internalId).toBe('');
    expect(update).not.toHaveBeenCalled();
  });
  it('retains every blank-ID row independently and exports them without making any discussion requests', async () => {
    const source = inventory([unresolved(0), unresolved(1), unresolved(2)]);
    const executeScript = vi.fn(); const update = vi.fn();
    vi.stubGlobal('chrome', { tabs: {get: async () => ({url: source.project.url}), update}, scripting: {executeScript} });
    const { collectOneOfficeProjectTaskDiscussions } = await import('../../src/oneoffice/projectBatchCollector');
    const progress = vi.fn();
    const result = await collectOneOfficeProjectTaskDiscussions(5, source, progress, () => false);
    expect(result.tasks).toHaveLength(3);
    expect(result.tasks.map(row => row.task.name)).toEqual(source.tasks.map(row => row.name));
    expect(result.tasks.every(row => row.state === 'unresolved' && row.discussion === null)).toBe(true);
    expect(progress.mock.calls.at(-1)?.[0]).toMatchObject({completedCount: 3, totalCount: 3});
    expect(api.comments).not.toHaveBeenCalled(); expect(api.fast).not.toHaveBeenCalled();
    expect(update).not.toHaveBeenCalled(); expect(executeScript).not.toHaveBeenCalled();
    const workbook = buildOneOfficeProjectWorkbook(result);
    expect(workbook[1].data).toHaveLength(4);
    expect(workbook[2].data).toHaveLength(1);
    expect(workbook[1].data.slice(1).every(row => (row[10] as any).value === '' && (row[24] as any).value === 'Chưa lấy được thảo luận')).toBe(true);
    expect(JSON.stringify(workbook)).not.toContain('report-row:');
  });
  it('collects only confirmed IDs while preserving unresolved rows and their original order', async () => {
    const task = { ...unresolved(1), internalId: '89876', identityError: undefined, code: 'CODE', url: `${origin}/apps/work-normal-normal/view?ID=89876` };
    const source = inventory([unresolved(0), task, unresolved(2)]);
    vi.stubGlobal('chrome', { tabs: {get: async () => ({url: source.project.url})} });
    api.comments.mockResolvedValue({schemaVersion: 1, source: '1office-comment-api', entity: {type: 'task', internalId: task.internalId},
      comments: [], collection: {complete: true, loadedRootCommentCount: 0, loadedReplyCount: 0, warnings: []} });
    const { collectOneOfficeProjectTaskDiscussions } = await import('../../src/oneoffice/projectBatchCollector');
    const result = await collectOneOfficeProjectTaskDiscussions(5, source, () => {}, () => false);
    expect(result.tasks.map(row => row.state)).toEqual(['unresolved', 'complete', 'unresolved']);
    expect(api.comments).toHaveBeenCalledExactlyOnceWith(5, task);
    expect(api.fast).not.toHaveBeenCalled();
  });
});
