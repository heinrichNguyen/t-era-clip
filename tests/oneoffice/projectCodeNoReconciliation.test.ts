import { expect, it } from 'vitest';
import { reconcileProjectCodeNo, type CodeNoOptions } from '../../src/oneoffice/projectCodeNoReconciliation';
import { reviewProjectImport } from '../../src/oneoffice/projectImportReview';
import type { OneOfficeProjectInventory } from '../../src/oneoffice/types';
import type { TestTaskDetails } from '../../src/oneoffice/testTaskDetails';

const origin = 'https://testoffice.meygroup.vn:9443';
const inventory = (): OneOfficeProjectInventory => ({ schemaVersion: 1, source: '1office-project-api', capturedAt: 'snapshot',
  project: { type: 'project', internalId: '203', code: null, name: 'Test', url: `${origin}/apps/work-project-project/view?ID=203` },
  collection: { complete: true, expectedTaskCount: 3, loadedTaskCount: 3, warnings: [] },
  tasks: ['A', 'B', 'C'].map((code, i) => ({ internalId: String(i + 1), code, no: `outline-${i}`, sourceNo: `1.${i + 1}`, name: code, parentInternalId: i ? '1' : null,
    parentNo: null, depth: i ? 2 : 1, sourceOrder: i, url: `${origin}/apps/work-task-task/view?ID=${i + 1}` })),
});
const options: CodeNoOptions = { idColumn: 'ID', oldCodeColumn: '', codeColumn: 'Mã gói thầu/Mã CV', noColumn: 'No.', mappings: '', approvedFile: true, confirmedProject: true, noMeaning: 'custom-no' };
const report = (rows: (string | number | null)[][]) => reviewProjectImport({ sheet: 'Approved', data: [['No.', 'Mã gói thầu/Mã CV', 'Nội dung công việc', 'ID', 'Mã cũ'], ...rows] }, null);
const fixture = () => report([['1.2', 'B', 'Task 1', '1', 'A'], ['1.1', 'A', 'Task 2', '2', 'B']]);
it('changes both code and No. by stable ID, preserving tasks and parent relationships', () => {
  const data = inventory(); const before = JSON.stringify(data);
  const result = reconcileProjectCodeNo(fixture(), data, options);
  expect(result.errors).toEqual([]);
  expect(result.rows.map((row) => [row.internalId, row.beforeCode, row.afterCode, row.beforeNo, row.afterNo, row.state])).toEqual([
    ['1', 'A', 'B', '1.1', '1.2', 'changed'], ['2', 'B', 'A', '1.2', '1.1', 'changed'],
  ]);
  expect(result.operations).toHaveLength(6); expect(result.operations.filter((step) => step.temporary)).toHaveLength(2);
  expect(result.unchangedOutsideFile).toBe(1); expect(result.writeEnabled).toBe(false);
  expect(JSON.stringify(data)).toBe(before);
});
it('simulates every planned step without collisions and reaches approved values', () => {
  const data = inventory(); const result = reconcileProjectCodeNo(fixture(), data, options);
  const state = new Map(data.tasks.map((task) => [task.internalId, { code: task.code!, sourceNo: task.sourceNo! }]));
  for (const step of result.operations) {
    expect(state.get(step.internalId)![step.field]).toBe(step.before);
    expect([...state].some(([id, values]) => id !== step.internalId && values[step.field] === step.after)).toBe(false);
    state.get(step.internalId)![step.field] = step.after;
  }
  expect(state.get('1')).toEqual({ code: 'B', sourceNo: '1.2' });
  expect(state.get('2')).toEqual({ code: 'A', sourceNo: '1.1' });
  expect(state.get('3')).toEqual({ code: 'C', sourceNo: '1.3' });
});
it('assigns a code to an already-created task with a verified blank code by ID', () => {
  const data = inventory(); data.tasks[0].code = ''; data.tasks[0].sourceNo = '';
  const result = reconcileProjectCodeNo(report([['1.4', 'PROJECT.NEW', 'Imported task', '1', '']]), data, options);
  expect(result.rows[0].beforeCode).toBe(''); expect(result.rows[0].state).toBe('changed');
  expect(result.operations).toHaveLength(2);
});
it('does not infer identity from a task name or No., or classify an unmatched row as new', () => {
  const result = reconcileProjectCodeNo(report([['1.1', 'NEW.A', 'A', '', '']]), inventory(), options);
  expect(result.rows[0].internalId).toBeNull(); expect(result.operations).toEqual([]);
  expect(result.rows[0].issues.join(' ')).toContain('Không tự coi đây là việc mới');
});
it('uses a unique old code when target code changes, or an explicit row-to-ID mapping', () => {
  const source = report([['1.1', 'NEW.A', 'Task', '', 'A']]);
  const byCode = reconcileProjectCodeNo(source, inventory(), { ...options, oldCodeColumn: 'Mã cũ' });
  expect(byCode.rows[0].internalId).toBe('1'); expect(byCode.rows[0].match).toBe('old-code');
  const byId = reconcileProjectCodeNo(source, inventory(), { ...options, mappings: '2 1' });
  expect(byId.rows[0].internalId).toBe('1'); expect(byId.rows[0].afterCode).toBe('NEW.A');
});
it.each(['duplicate-id', 'duplicate-code', 'duplicate-no', 'unknown-id', 'conflicting-id', 'old-code-mismatch', 'blank-code', 'blank-no', 'letter-no'])('blocks unsafe mappings or targets: %s', (mode) => {
  const source = fixture(); const config = { ...options };
  if (mode === 'duplicate-id') source.rows[1].values.ID = '1';
  if (mode === 'duplicate-code') source.rows[1].values[config.codeColumn] = 'B';
  if (mode === 'duplicate-no') source.rows[1].values[config.noColumn] = '1.2';
  if (mode === 'unknown-id') source.rows[0].values.ID = '999';
  if (mode === 'conflicting-id') config.mappings = '2 2';
  if (mode === 'old-code-mismatch') { config.oldCodeColumn = 'Mã cũ'; source.rows[0].values['Mã cũ'] = 'C'; }
  if (mode === 'blank-code') source.rows[0].values[config.codeColumn] = '';
  if (mode === 'blank-no') source.rows[0].values[config.noColumn] = '';
  if (mode === 'letter-no') source.rows[0].values[config.noColumn] = '1.2a';
  const result = reconcileProjectCodeNo(source, inventory(), config);
  expect(result.rows.some((row) => row.state === 'review')).toBe(true); expect(result.operations).toEqual([]);
});
it.each(['file', 'project', 'no-meaning', 'outline', 'missing-column', 'partial', 'project-url', 'foreign-task'])('blocks missing confirmations or unsafe scope: %s', (mode) => {
  const data = inventory(); const config = { ...options };
  if (mode === 'file') config.approvedFile = false;
  if (mode === 'project') config.confirmedProject = false;
  if (mode === 'no-meaning') config.noMeaning = '';
  if (mode === 'outline') config.noMeaning = 'outline';
  if (mode === 'missing-column') config.idColumn = 'Absent';
  if (mode === 'partial') data.collection.complete = false;
  if (mode === 'project-url') data.project.internalId = '999';
  if (mode === 'foreign-task') data.tasks[0].url = 'https://example.com/apps/work-task-task/view?ID=1';
  const result = reconcileProjectCodeNo(fixture(), data, config);
  expect(result.errors.length).toBeGreaterThan(0); expect(result.operations).toEqual([]);
});
it('blocks target collisions with untouched tasks and unknown fields outside the file', () => {
  const data = inventory(); data.tasks[2].sourceNo = undefined;
  const result = reconcileProjectCodeNo(report([['1.1', 'C', 'Task', '1', '']]), data, options);
  expect(result.rows[0].issues.join(' ')).toContain('ngoài file');
  expect(result.errors.join(' ')).toContain('Chưa đủ'); expect(result.operations).toEqual([]);
});
it('never substitutes outline numbers for raw No. and overlays verified detail without mutating the inventory', () => {
  const data = inventory(); data.tasks[0].code = null; data.tasks[0].sourceNo = undefined;
  const source = report([['1.4', 'NEW.A', 'Task', '1', '']]);
  expect(reconcileProjectCodeNo(source, data, options).rows[0].beforeNo).toBeNull();
  const details: TestTaskDetails = { taskId: '1', sourceUrl: `${origin}/apps/work-task-task/view?ID=1`, capturedAt: 'now', identityVerified: true,
    evidenceSource: 'widget-response', code: { state: 'value', displayValue: 'A' }, no: { state: 'empty', displayValue: '' }, error: null };
  const result = reconcileProjectCodeNo(source, data, options, { '1': details });
  expect(result.rows[0].beforeCode).toBe('A'); expect(result.rows[0].beforeNo).toBe(''); expect(result.operations).toHaveLength(2);
  expect(data.tasks[0].sourceNo).toBeUndefined();
  expect(reconcileProjectCodeNo(source, data, options, { '1': { ...details, taskId: '999' } }).errors.join(' ')).toContain('nguồn chi tiết');
});
it('plans a 1000-task cycle using only one temporary code and one temporary No.', () => {
  const data = inventory(); data.tasks = Array.from({ length: 1000 }, (_, i) => ({ ...data.tasks[0], internalId: String(i + 1), code: `P.${i + 1}`, sourceNo: `1.${i + 1}`, url: `${origin}/apps/work-task-task/view?ID=${i + 1}` }));
  data.collection.loadedTaskCount = data.collection.expectedTaskCount = data.tasks.length;
  const source = report(data.tasks.map((task, i) => [`1.${(i + 1) % 1000 + 1}`, `P.${(i + 1) % 1000 + 1}`, task.name, task.internalId, '']));
  const result = reconcileProjectCodeNo(source, data, options);
  expect(result.errors).toEqual([]); expect(result.operations).toHaveLength(2002);
  expect(result.operations.filter((step) => step.temporary)).toHaveLength(2);
});
it.each(['numeric-code', 'numeric-outline', 'unsafe-numeric-id'])('rejects lossy or partial Excel identifiers: %s', (mode) => {
  const source = fixture();
  if (mode === 'numeric-code') source.rows[0].values[options.codeColumn] = '1';
  if (mode === 'numeric-outline') source.rows[0].values[options.noColumn] = 1.1;
  if (mode === 'unsafe-numeric-id') source.rows[0].values.ID = 9007199254740992;
  const result = reconcileProjectCodeNo(source, inventory(), options);
  expect(result.rows[0].state).toBe('review'); expect(result.operations).toEqual([]);
});
