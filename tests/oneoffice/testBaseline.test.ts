import { afterEach, expect, it, vi } from 'vitest';
import { collectTestBaseline, selectTestBaselineSamples, TEST_BASELINE_ORIGIN as origin } from '../../src/oneoffice/testBaseline';
import { collectOneOfficeProjectInventory } from '../../src/oneoffice/collectionClient';
import { collectOneOfficeTaskFromCommentApi } from '../../src/oneoffice/projectCommentApi';
import { collectTestTaskDetails } from '../../src/oneoffice/testTaskDetails';
import type { OneOfficeProjectInventory, OneOfficeProjectTaskInventoryItem } from '../../src/oneoffice/types';

vi.mock('../../src/oneoffice/collectionClient', () => ({ collectOneOfficeProjectInventory: vi.fn() }));
vi.mock('../../src/oneoffice/projectCommentApi', () => ({ collectOneOfficeTaskFromCommentApi: vi.fn() }));
vi.mock('../../src/oneoffice/testTaskDetails', () => ({ collectTestTaskDetails: vi.fn() }));
afterEach(() => { vi.resetAllMocks(); });
const task = (id: string, parent: string | null = null, depth = 1): OneOfficeProjectTaskInventoryItem => ({
  internalId: id, code: `TEST.${id}`, name: `Task ${id}`, no: 'same-visible-no', sourceNo: id === '4' ? '' : '1',
  parentInternalId: parent, parentNo: null, depth, sourceOrder: Number(id), url: `${origin}/apps/work-task-task/view?ID=${id}`,
});
const inventory = (): OneOfficeProjectInventory => ({
  schemaVersion: 1, source: '1office-project-api', capturedAt: '2026-10-21',
  project: { type: 'project', internalId: '203', name: 'TEST', code: null, url: `${origin}/apps/work-project-project/view?ID=203` },
  collection: { complete: true, loadedTaskCount: 6, expectedTaskCount: 6, warnings: [] },
  tasks: [task('1'), task('2', '1', 2), task('3', '2', 3), task('4'), { ...task('5'), assigneeNames: ['Test user'] }, task('6')],
});
it('selects five distinct IDs including a real parent-child pair, deep leaf and raw blank No.', () => {
  const samples = selectTestBaselineSamples(inventory());
  expect(samples.map((sample) => sample.task.internalId)).toEqual(['1', '2', '3', '4', '5']);
  expect(samples[3].task.sourceNo).toBe('');
  expect(samples[1].task.parentInternalId).toBe(samples[0].task.internalId);
});
it('does not treat an unverified No. as a verified blank', () => {
  const data = inventory(); data.tasks[3].sourceNo = undefined;
  const samples = selectTestBaselineSamples(data);
  expect(samples.some((sample) => sample.reason.startsWith('No. gốc đang trống'))).toBe(false);
});
it.each(['production', 'project', 'count', 'duplicate', 'missing-parent', 'cycle', 'url-id', 'mixed-origin'])('rejects unsafe source: %s', (mode) => {
  const data = inventory();
  if (mode === 'production') data.project.url = 'https://office.meygroup.vn/apps/work-project-project/view?ID=203';
  if (mode === 'project') data.project.internalId = '870';
  if (mode === 'count') data.collection.expectedTaskCount = 7;
  if (mode === 'duplicate') data.tasks[5] = { ...data.tasks[0] };
  if (mode === 'missing-parent') data.tasks[1].parentInternalId = '999';
  if (mode === 'cycle') data.tasks[0].parentInternalId = '3';
  if (mode === 'url-id') data.tasks[1].url = `${origin}/apps/work-task-task/view?ID=999`;
  if (mode === 'mixed-origin') data.tasks[1].url = 'https://office.meygroup.vn/apps/work-task-task/view?ID=2';
  expect(() => selectTestBaselineSamples(data)).toThrow();
});
it('reads a fresh inventory and only five discussions, recording gaps without a write/navigation fallback', async () => {
  vi.mocked(collectOneOfficeProjectInventory).mockResolvedValue(inventory());
  vi.mocked(collectOneOfficeTaskFromCommentApi).mockResolvedValue(null).mockRejectedValueOnce(new Error('HTTP 403'));
  vi.mocked(collectTestTaskDetails).mockRejectedValue(new Error('Chi tiết chưa xác minh'));
  const chrome = { tabs: { update: vi.fn() }, storage: { local: { set: vi.fn() } } };
  vi.stubGlobal('chrome', chrome);
  try {
    const result = await collectTestBaseline(42, vi.fn(), () => false);
    expect(collectOneOfficeProjectInventory).toHaveBeenCalledWith(42, origin);
    expect(collectOneOfficeTaskFromCommentApi).toHaveBeenCalledTimes(5);
    expect(collectTestTaskDetails).toHaveBeenCalledTimes(5);
    expect(result.samples[0].detailError).toBe('Chi tiết chưa xác minh');
    expect(result.mode).toBe('read-only-test-baseline');
    expect(result.samples[0].error).toBe('HTTP 403');
    expect(result.samples.every((sample) => sample.error)).toBe(true);
    expect(result.limitations.join(' ')).toContain('phần trăm hoàn thành');
    expect(chrome.tabs.update).not.toHaveBeenCalled(); expect(chrome.storage.local.set).not.toHaveBeenCalled();
  } finally { vi.unstubAllGlobals(); }
});
it('stops before a read when cancelled, and refuses incomplete inventory before any discussion request', async () => {
  await expect(collectTestBaseline(42, vi.fn(), () => true)).rejects.toThrow('Đã hủy');
  expect(collectOneOfficeProjectInventory).not.toHaveBeenCalled();
  const data = inventory(); data.collection.complete = false;
  vi.mocked(collectOneOfficeProjectInventory).mockResolvedValue(data);
  await expect(collectTestBaseline(42, vi.fn(), () => false)).rejects.toThrow('đầy đủ');
  expect(collectOneOfficeTaskFromCommentApi).not.toHaveBeenCalled();
});
