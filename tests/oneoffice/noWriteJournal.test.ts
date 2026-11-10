import { afterEach, expect, it, vi } from 'vitest';
import { clearResolvedNoWriteJournal, finishNoWriteJournal, prepareNoWriteJournal, readNoWriteJournal, reconcileNoWriteJournal, withNoWriteLock } from '../../src/oneoffice/noWriteJournal';
import { inspectNoWriteJournal, runNoUpdatePlan, type NoExperimentResult, type NoWriteRequest } from '../../src/oneoffice/noWriteExperiment';
import type { OneOfficeProjectInventory } from '../../src/oneoffice/types';
import { NO_TEST_TASKS } from '../../src/oneoffice/noTestScope';

afterEach(() => vi.unstubAllGlobals());
const request = (): NoWriteRequest => ({
  changes: [
    { internalId: '100119', code: 'TEST01.3', before: '1.3', after: '1.4' },
    { internalId: '100122', code: 'TEST01.INSERT', before: '1.4', after: '1.3' },
  ], sourceTasks: [{ internalId: '100119', no: '1.3' }, { internalId: '100122', no: '1.4' }],
});
const readResult = (first: string, second: string): NoExperimentResult => ({
  readOnly: true, complete: true, error: null, phase: 'Đã đọc', saveAttempts: 0,
  tasks: [{ code: 'TEST01.3', before: first, temporary: '99001.1', applied: false, restored: false, error: null },
    { code: 'TEST01.INSERT', before: second, temporary: '99001.2', applied: false, restored: false, error: null }],
});
function setup() {
  const data: Record<string, any> = {};
  let failSet = false;
  const order: string[] = [];
  const api = { runtime: {} as { lastError?: unknown }, storage: { local: {
    get: (key: string, done: (value: unknown) => void) => done(structuredClone({ [key]: data[key] })),
    set: (value: Record<string, unknown>, done: () => void) => {
      order.push('persist');
      if (failSet) api.runtime.lastError = { message: 'quota' }; else Object.assign(data, structuredClone(value));
      done(); delete api.runtime.lastError;
    },
    remove: (key: string, done: () => void) => { delete data[key]; done(); },
  } }, tabs: { get: vi.fn().mockResolvedValue({ url: 'https://office.meygroup.vn/apps/work-project-project/view?ID=870' }), update: vi.fn().mockResolvedValue({ active: true }) },
  scripting: { executeScript: vi.fn(async () => { order.push('execute'); return [{ result: readResult('1.3', '1.4') }]; }) } };
  vi.stubGlobal('chrome', api);
  vi.stubGlobal('navigator', { locks: { request: async (_: string, __: unknown, action: (lock: unknown) => unknown) => action({}) } });
  return { api, data, order, fail: () => { failSet = true; } };
}

it('persists intent before execution and refuses to replace an unresolved journal', async () => {
  const app = setup(); const journal = await prepareNoWriteJournal(request());
  expect(await readNoWriteJournal()).toEqual(journal);
  expect(journal.status).toBe('pending');
  expect(journal.rows.every((row) => row.state === 'unknown')).toBe(true);
  await expect(prepareNoWriteJournal(request())).rejects.toThrow('chưa đối soát');
  expect(app.order).toEqual(['persist']);
});
it('persists and reconciles five tasks with distinct temporary numbers and an original blank', async () => {
  setup();
  const changes = NO_TEST_TASKS.map((task, index) => ({ internalId: task.internalId, code: task.code, before: index === 4 ? '' : `1.${index + 1}`, after: `2.${index + 1}` }));
  const journal = await prepareNoWriteJournal({ changes, sourceTasks: [] });
  expect(journal.rows).toHaveLength(5);
  expect(new Set(journal.rows.map((row) => row.temporary)).size).toBe(5);
  expect((await readNoWriteJournal())!.rows[4].before).toBe('');
  const result = readResult('1.3', '1.4');
  result.tasks = changes.map((row) => ({ code: row.code, before: row.after, temporary: '', applied: false, restored: false, error: null }));
  expect((await reconcileNoWriteJournal(journal, result)).status).toBe('verified');
});

it.each([
  ['1.4', '1.3', 'verified'], ['1.3', '1.4', 'unchanged'], ['99001.1', '1.3', 'pending'], ['9.9', '1.3', 'pending'],
])('reconciles %s / %s without replaying writes', async (first, second, status) => {
  setup(); const journal = await prepareNoWriteJournal(request());
  const reconciled = await reconcileNoWriteJournal(journal, readResult(first, second));
  expect(reconciled.status).toBe(status);
  expect(reconciled.rows.map((row) => row.observed)).toEqual([first, second]);
  expect(await readNoWriteJournal()).toEqual(reconciled);
});

it('does not resolve an incomplete read or permit deletion of pending intent', async () => {
  setup(); const journal = await prepareNoWriteJournal(request());
  await expect(reconcileNoWriteJournal(journal, { ...readResult('1.4', '1.3'), complete: false })).rejects.toThrow('Chưa đọc');
  await expect(clearResolvedNoWriteJournal()).rejects.toThrow('Không xóa');
  expect((await readNoWriteJournal())?.status).toBe('pending');
});

it('stores only minimal fields and permits deleting a fully reconciled record', async () => {
  const app = setup(); const input = request();
  Object.assign(input.changes[0], { csrfToken: 'never-store', title: 'private title' });
  const journal = await prepareNoWriteJournal(input);
  expect(JSON.stringify(app.data)).not.toMatch(/never-store|private title|csrfToken/);
  await reconcileNoWriteJournal(journal, readResult('1.4', '1.3'));
  await clearResolvedNoWriteJournal();
  expect(await readNoWriteJournal()).toBeNull();
});

it('keeps unconfirmed rows pending when only part of the result was verified', async () => {
  setup(); const journal = await prepareNoWriteJournal(request());
  const result = readResult('1.3', '1.4'); result.readOnly = false; result.complete = false; result.saveAttempts = 2;
  result.tasks[1] = { ...result.tasks[1], before: '1.4', after: '1.3', finalized: true };
  const saved = await finishNoWriteJournal(journal, result);
  expect(saved.status).toBe('pending'); expect(saved.rows[0].state).toBe('unknown'); expect(saved.rows[1].state).toBe('after');
});

const inventory = () => ({ project: { internalId: '870', url: 'https://office.meygroup.vn/apps/work-project-project/view?ID=870' }, collection: { complete: true },
  tasks: request().changes.map((row) => ({ ...row, name: row.code, no: row.before, sourceNo: row.before })),
}) as OneOfficeProjectInventory;

it('never injects a write when local persistence fails', async () => {
  const app = setup(); app.fail();
  await expect(runNoUpdatePlan(1, inventory(), 'TEST01.3\t1.4\nTEST01.INSERT\t1.3')).rejects.toThrow('nhật ký');
  expect(app.api.scripting.executeScript).not.toHaveBeenCalled(); expect(app.api.tabs.update).not.toHaveBeenCalled();
});

it('leaves durable pending intent if the source execution is interrupted', async () => {
  const app = setup(); app.api.scripting.executeScript.mockRejectedValueOnce(new Error('tab closed'));
  await expect(runNoUpdatePlan(1, inventory(), 'TEST01.3\t1.4\nTEST01.INSERT\t1.3')).rejects.toThrow('tab closed');
  expect(app.order).toEqual(['persist']); // The rejected mock bypasses its recording implementation.
  expect((await readNoWriteJournal())?.status).toBe('pending');
  await expect(runNoUpdatePlan(1, inventory(), 'TEST01.3\t1.4\nTEST01.INSERT\t1.3')).rejects.toThrow('chưa đối soát');
  expect(app.api.scripting.executeScript).toHaveBeenCalledTimes(1);
});

it('recovery injects only the read-only runner and resolves a previously interrupted record', async () => {
  const app = setup(); await prepareNoWriteJournal(request());
  expect((await inspectNoWriteJournal(1)).status).toBe('unchanged');
  expect(app.api.scripting.executeScript).toHaveBeenCalledWith(expect.objectContaining({ args: [true, null, true, false, ['100119', '100122']] }));
});

it('refuses a second extension writer when a lock is already held', async () => {
  setup(); vi.stubGlobal('navigator', { locks: { request: async (_: string, __: unknown, action: (lock: unknown) => unknown) => action(null) } });
  const action = vi.fn(); await expect(withNoWriteLock(action)).rejects.toThrow('Một trang'); expect(action).not.toHaveBeenCalled();
});

it('fails closed for corrupt persisted records instead of treating them as absent', async () => {
  const app = setup(); app.data.oneOfficeNoWriteJournalV1 = { version: 9 };
  await expect(readNoWriteJournal()).rejects.toThrow('không hợp lệ');
  await expect(prepareNoWriteJournal(request())).rejects.toThrow('không hợp lệ');
});
