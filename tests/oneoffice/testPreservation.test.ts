import { afterEach, expect, it, vi } from 'vitest';
import { findTestPreservationSample, type TestPreservationProbe } from '../../src/oneoffice/testPreservation';
import { collectOneOfficeProjectInventory } from '../../src/oneoffice/collectionClient';
import { collectOneOfficeTaskFromCommentApi } from '../../src/oneoffice/projectCommentApi';
import { collectTestTaskDetails } from '../../src/oneoffice/testTaskDetails';
import { TEST_BASELINE_ORIGIN as origin, type TestBaseline } from '../../src/oneoffice/testBaseline';
import type { OneOfficeDiscussionExport, OneOfficeProjectInventory } from '../../src/oneoffice/types';

vi.mock('../../src/oneoffice/collectionClient', () => ({ collectOneOfficeProjectInventory: vi.fn() }));
vi.mock('../../src/oneoffice/projectCommentApi', () => ({ collectOneOfficeTaskFromCommentApi: vi.fn() }));
vi.mock('../../src/oneoffice/testTaskDetails', () => ({ collectTestTaskDetails: vi.fn() }));
afterEach(() => { vi.resetAllMocks(); });
const inventory = (count = 25): OneOfficeProjectInventory => ({
  schemaVersion: 1, source: '1office-project-api', capturedAt: '2026-10-22',
  project: { type: 'project', internalId: '203', name: 'Test', code: null, url: `${origin}/apps/work-project-project/view?ID=203` },
  collection: { complete: true, expectedTaskCount: count, loadedTaskCount: count, warnings: [] },
  tasks: Array.from({ length: count }, (_, i) => ({ internalId: String(i + 1), name: `Task ${i + 1}`, code: null,
    no: String(i + 1), depth: 1, parentNo: null, parentInternalId: null, sourceOrder: i,
    url: `${origin}/apps/work-task-task/view?ID=${i + 1}` })),
});
const baseline = (): TestBaseline => ({ schemaVersion: 1, mode: 'read-only-test-baseline', origin, startedAt: 'start', completedAt: 'end',
  inventory: inventory(), samples: [], limitations: [] });
const discussion = (id: string, count = 0, complete = true) => ({
  entity: { internalId: id }, collection: { complete }, comments: Array.from({ length: count }, (_, i) => ({ id: String(i + 100) })),
}) as OneOfficeDiscussionExport;
const run = (b = baseline(), previous: TestPreservationProbe | null = null, id: string | null = null, cancelled = () => false) =>
  findTestPreservationSample(42, b, previous, id, vi.fn(), cancelled);
const mockEmpty = () => {
  vi.mocked(collectOneOfficeProjectInventory).mockResolvedValue(inventory());
  vi.mocked(collectOneOfficeTaskFromCommentApi).mockImplementation(async (_tab, task) => discussion(task.internalId));
};
it('reads a fresh inventory and at most 20 discussions, with explicit continuation', async () => {
  mockEmpty(); const b = baseline(); const untouched = JSON.stringify(b);
  const first = await run(b);
  expect(first.attempts).toHaveLength(20); expect(first.remaining).toBe(5); expect(first.sample).toBeNull();
  expect(first.reportResults).toBe('not-collected');
  expect(collectTestTaskDetails).not.toHaveBeenCalled();
  const second = await run(b, first);
  expect(second.attempts).toHaveLength(25); expect(second.remaining).toBe(0);
  expect(collectOneOfficeProjectInventory).toHaveBeenCalledTimes(2);
  expect(collectOneOfficeTaskFromCommentApi).toHaveBeenCalledTimes(25);
  expect(new Set(vi.mocked(collectOneOfficeTaskFromCommentApi).mock.calls.map((call) => call[1].internalId)).size).toBe(25);
  expect(JSON.stringify(b)).toBe(untouched);
});
it('stops at the first complete nonempty discussion and reads only that detail', async () => {
  mockEmpty(); vi.mocked(collectOneOfficeTaskFromCommentApi).mockImplementation(async (_tab, task) => discussion(task.internalId, task.internalId === '3' ? 2 : 0));
  vi.mocked(collectTestTaskDetails).mockResolvedValue({ taskId: '3', identityVerified: true, progress: { state: 'value', percent: 0 } } as any);
  const result = await run();
  expect(result.attempts).toHaveLength(3); expect(result.remaining).toBe(22);
  expect(result.sample?.discussion?.comments).toHaveLength(2);
  expect(result.sample?.details?.progress?.percent).toBe(0);
  expect(collectTestTaskDetails).toHaveBeenCalledExactlyOnceWith(42, '3');
});
it('records failed and incomplete reads as unverified, not as zero comments, without retries', async () => {
  mockEmpty(); vi.mocked(collectOneOfficeTaskFromCommentApi).mockRejectedValueOnce(new Error('403')).mockResolvedValueOnce(discussion('2', 3, false)).mockResolvedValueOnce(discussion('999', 1));
  const result = await run();
  expect(result.attempts.slice(0, 3).map((item) => item.state)).toEqual(['unverified', 'unverified', 'unverified']);
  expect(result.sample).toBeNull(); expect(collectOneOfficeTaskFromCommentApi).toHaveBeenCalledTimes(20);
});
it('retains discussion evidence if the detail read fails, without approving the sample', async () => {
  mockEmpty(); vi.mocked(collectOneOfficeTaskFromCommentApi).mockResolvedValue(discussion('1', 1));
  vi.mocked(collectTestTaskDetails).mockRejectedValue(new Error('private payload'));
  const result = await run();
  expect(result.sample?.discussion?.comments).toHaveLength(1);
  expect(result.sample?.detailError).toContain('Chưa đọc được');
  expect(JSON.stringify(result)).not.toContain('private payload');
  expect(result).not.toHaveProperty('approved');
});
it('looks up a manual ID only in the fresh test inventory', async () => {
  mockEmpty(); await run(baseline(), null, '23');
  expect(collectOneOfficeTaskFromCommentApi).toHaveBeenCalledExactlyOnceWith(42, expect.objectContaining({ internalId: '23' }));
});
it.each(['999', '../1', '1&ID=2', ''])('rejects an unknown or malformed ID before any request: %s', async (id) => {
  await expect(run(baseline(), null, id)).rejects.toThrow('ID');
  expect(collectOneOfficeProjectInventory).not.toHaveBeenCalled();
  expect(collectOneOfficeTaskFromCommentApi).not.toHaveBeenCalled();
});
it.each(['project', 'origin', 'partial', 'duplicate', 'tree-change', 'ids-change'])('refuses unsafe or changed fresh inventory: %s', async (mode) => {
  const fresh = inventory();
  if (mode === 'project') fresh.project.internalId = '870';
  if (mode === 'origin') fresh.project.url = 'https://office.meygroup.vn/apps/work-project-project/view?ID=203';
  if (mode === 'partial') fresh.collection.complete = false;
  if (mode === 'duplicate') fresh.tasks[1] = fresh.tasks[0];
  if (mode === 'tree-change') fresh.tasks[1].parentInternalId = '1';
  if (mode === 'ids-change') fresh.tasks[1] = { ...fresh.tasks[1], internalId: '99', url: `${origin}/apps/work-task-task/view?ID=99` };
  vi.mocked(collectOneOfficeProjectInventory).mockResolvedValue(fresh);
  await expect(run()).rejects.toThrow();
  expect(collectOneOfficeTaskFromCommentApi).not.toHaveBeenCalled();
});
it('refuses a cursor from a different baseline', async () => {
  mockEmpty(); const previous = await run(); previous.baselineStartedAt = 'different'; vi.clearAllMocks();
  await expect(run(baseline(), previous)).rejects.toThrow('hiện trạng');
  expect(collectOneOfficeProjectInventory).not.toHaveBeenCalled();
});
it('stops between requests when cancelled, without starting detail reads', async () => {
  mockEmpty(); let stopped = false;
  vi.mocked(collectOneOfficeTaskFromCommentApi).mockImplementation(async () => { stopped = true; return discussion('1', 1); });
  await expect(run(baseline(), null, null, () => stopped)).rejects.toThrow('Đã dừng');
  expect(collectOneOfficeTaskFromCommentApi).toHaveBeenCalledTimes(1);
  expect(collectTestTaskDetails).not.toHaveBeenCalled();
});
