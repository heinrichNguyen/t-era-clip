import { collectOneOfficeProjectInventory } from './collectionClient';
import { collectOneOfficeTaskFromCommentApi } from './projectCommentApi';
import { collectTestTaskDetails } from './testTaskDetails';
import { selectTestBaselineSamples, TEST_BASELINE_ORIGIN, type TestBaseline, type TestBaselineSample } from './testBaseline';

export interface TestPreservationProbe {
  mode: 'read-only-test-preservation';
  baselineStartedAt: string;
  origin: string;
  capturedAt: string;
  attempts: { taskId: string; state: 'empty' | 'found' | 'unverified' }[];
  sample: TestBaselineSample | null;
  remaining: number;
  reportResults: 'not-collected';
}

export const TEST_PRESERVATION_BATCH_SIZE = 20;

// A fresh inventory is mandatory. A search never approves a sample or starts a save.
export async function findTestPreservationSample(
  tabId: number, baseline: TestBaseline, previous: TestPreservationProbe | null, taskId: string | null,
  progress: (message: string) => void, cancelled: () => boolean,
): Promise<TestPreservationProbe> {
  const check = () => { if (cancelled()) throw new Error('Đã dừng tìm mẫu. Chưa ghi dữ liệu.'); };
  check();
  selectTestBaselineSamples(baseline.inventory);
  if (baseline.origin !== TEST_BASELINE_ORIGIN || baseline.mode !== 'read-only-test-baseline') throw new Error('Sai bản hiện trạng test.');
  if (taskId !== null && (!/^\d+$/.test(taskId) || !baseline.inventory.tasks.some((task) => task.internalId === taskId))) {
    throw new Error('ID phải thuộc danh sách hiện trạng dự án test 203.');
  }
  if (previous && (previous.mode !== 'read-only-test-preservation' || previous.baselineStartedAt !== baseline.startedAt || previous.origin !== baseline.origin ||
    previous.attempts.some((item) => !baseline.inventory.tasks.some((task) => task.internalId === item.taskId)))) throw new Error('Lượt tìm không thuộc bản hiện trạng này.');
  progress('Đang kiểm tra lại danh sách dự án test, không ghi...');
  const fresh = await collectOneOfficeProjectInventory(tabId, TEST_BASELINE_ORIGIN);
  check(); selectTestBaselineSamples(fresh);
  const hierarchy = (tasks: typeof fresh.tasks) => JSON.stringify(tasks.map((task) => [task.internalId, task.parentInternalId]).sort((a, b) => a[0]!.localeCompare(b[0]!)));
  if (hierarchy(fresh.tasks) !== hierarchy(baseline.inventory.tasks)) throw new Error('Danh sách hoặc quan hệ cha/con đã thay đổi. Hãy lập lại bản hiện trạng trước khi tìm tiếp.');
  const attempts = [...(previous?.attempts ?? [])];
  const attempted = new Set(attempts.map((item) => item.taskId));
  const tasks = taskId === null ? fresh.tasks.filter((task) => !attempted.has(task.internalId)).slice(0, TEST_PRESERVATION_BATCH_SIZE)
    : fresh.tasks.filter((task) => task.internalId === taskId);
  let sample: TestBaselineSample | null = null;
  for (const [index, task] of tasks.entries()) {
    check(); progress(`Đang tìm thảo luận ${index + 1}/${tasks.length}, ID ${task.internalId}...`);
    let state: TestPreservationProbe['attempts'][number]['state'] = 'unverified';
    try {
      const discussion = await collectOneOfficeTaskFromCommentApi(tabId, task);
      check();
      if (discussion?.collection.complete && discussion.entity.internalId === task.internalId) {
        state = discussion.comments.length ? 'found' : 'empty';
        if (state === 'found') sample = { task, reason: 'Có thảo luận đọc được; chưa phê duyệt ghi', discussion, error: null };
      }
    } catch { check(); }
    const entry = { taskId: task.internalId, state };
    const existing = attempts.findIndex((item) => item.taskId === task.internalId);
    if (existing >= 0) attempts[existing] = entry; else attempts.push(entry);
    if (sample) {
      check(); progress(`Đã tìm thấy thảo luận, đang đọc mã/No. và tiến độ ID ${task.internalId}...`);
      try { sample.details = await collectTestTaskDetails(tabId, task.internalId); sample.detailError = sample.details.error; }
      catch { sample.detailError = 'Chưa đọc được khối thông tin chi tiết của mẫu.'; }
      check(); break;
    }
  }
  return { mode: 'read-only-test-preservation', baselineStartedAt: baseline.startedAt, origin: baseline.origin,
    capturedAt: new Date().toISOString(), attempts, sample: sample ?? previous?.sample ?? null, remaining: fresh.tasks.length - new Set(attempts.map((item) => item.taskId)).size,
    reportResults: 'not-collected' };
}
