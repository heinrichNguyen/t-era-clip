import { describe, expect, it } from 'vitest';
import { buildImportDiff, importCalendarDate, type ImportDiffOptions } from '../../src/oneoffice/projectImportDiff';
import { reviewProjectImport } from '../../src/oneoffice/projectImportReview';
import { proposeImportNormalization } from '../../src/oneoffice/projectImportNormalization';
import type { OneOfficeProjectInventory, OneOfficeProjectTaskInventoryItem } from '../../src/oneoffice/types';

const oldTask = (patch = {}): OneOfficeProjectTaskInventoryItem => ({ internalId: '10', code: 'P.A', name: 'Old name', sourceNo: '2', no: '2', parentInternalId: null,
  sourceOrder: 0, url: '', depth: 1, parentNo: null, plannedStart: '01/10/2026', plannedEnd: '2026-10-30 00:00:00', baselineStart: '2026-09-01', baselineEnd: '2026-09-30', ...patch });
const snapshot = (tasks = [oldTask()]): OneOfficeProjectInventory => ({ schemaVersion: 1, source: '1office-project-api', capturedAt: '2026-10-20', project: { type: 'project', internalId: '1', code: 'P', name: 'Project', url: '' }, collection: { complete: true, expectedTaskCount: tasks.length, loadedTaskCount: tasks.length, warnings: [] }, tasks });
const raw = (extra: (string | number | Date | null)[] = ['2026-10-01', '2026-10-31', '2026-01-01']) => ({ sheet: 'S', data: [['No.', 'Mã gói thầu/Mã CV', 'Nội dung công việc', 'Bắt đầu', 'Kết thúc', 'Bắt đầu thực tế'], ['1', 'P.A', 'New name', ...extra]] });
const setup = (inventory = snapshot(), data = raw()) => {
  const report = reviewProjectImport(data, inventory);
  const proposal = proposeImportNormalization(report, inventory, { mode: 'update', prefix: 'P', firstTaskLevel: 1, packageLevel: 2, mappings: '', newRows: '' });
  return { report, proposal, inventory };
};
const opts = (columns: ImportDiffOptions['columns'] = { name: 'Nội dung công việc' }): ImportDiffOptions => ({ columns, confirmProject: true, confirmPlanDates: true });
const diff = (s: ReturnType<typeof setup>, options = opts()) => buildImportDiff(s.report, s.proposal, s.inventory, options);

describe('before/after comparison only, no write plan', () => {
  it('compares allowlisted fields by stable ID/code and normalizes equivalent calendar dates', () => {
    const result = diff(setup(), opts({ name: 'Nội dung công việc', sourceNo: 'No.', plannedStart: 'Bắt đầu', plannedEnd: 'Kết thúc' }));
    expect(result.errors).toEqual([]);
    expect(result.rows[0].state).toBe('changed');
    expect(result.rows[0].fields.map((f) => f.state)).toEqual(['changed','changed','same','changed']);
    expect(result.rows[0].fields[2].before).toBe('2026-10-01');
    expect(result.snapshotAt).toBe('2026-10-20');
  });
  it('keeps blank and NA sources without treating them as requests to clear existing values', () => {
    const s = setup(snapshot(), raw([null, 'NA', '2026-01-01']));
    const result = diff(s, opts({ plannedStart: 'Bắt đầu', plannedEnd: 'Kết thúc' }));
    expect(result.rows[0].fields.every((f) => f.state === 'keep')).toBe(true);
    expect(result.rows[0].state).toBe('same');
  });
  it('distinguishes an unverified No. from an explicitly known blank No.', () => {
    const unknown = diff(setup(snapshot([oldTask({ sourceNo: undefined, no: '99' })])), opts({ sourceNo: 'No.' }));
    expect(unknown.rows[0].fields[0].state).toBe('unknown');
    expect(unknown.rows[0].state).toBe('review');
    const blank = diff(setup(snapshot([oldTask({ sourceNo: '' })])), opts({ sourceNo: 'No.' }));
    expect(blank.rows[0].fields[0].state).toBe('changed'); expect(blank.rows[0].fields[0].beforeKnown).toBe(true);
  });
  it('does not assume missing date values are verified blank dates', () => {
    const result = diff(setup(snapshot([oldTask({ plannedStart: '' })])), opts({ plannedStart: 'Bắt đầu' }));
    expect(result.rows[0].state).toBe('review');
    expect(result.rows[0].fields[0].state).toBe('unknown');
  });
  it('requires project and date semantic confirmations', () => {
    expect(diff(setup(), { ...opts(), confirmProject: false }).errors.join(' ')).toContain('đúng dự án');
    const result = diff(setup(), { ...opts({ plannedStart: 'Bắt đầu' }), confirmPlanDates: false });
    expect(result.errors.join(' ')).toContain('ngày kế hoạch');
    expect(result.rows[0].fields).toEqual([]);
  });
  it('rejects actual-date mappings, protected fields, duplicated mappings and empty selections', () => {
    const s = setup();
    expect(diff(s, opts({ plannedStart: 'Bắt đầu thực tế' })).errors.join(' ')).toContain('Không dùng cột');
    expect(diff(s, opts({ progress: 'Tiến độ' } as ImportDiffOptions['columns'])).errors.join(' ')).toContain('không được phép');
    expect(diff(s, opts({ plannedStart: 'Bắt đầu', baselineStart: 'Bắt đầu' })).errors.join(' ')).toContain('nhiều trường');
    expect(diff(s, opts({})).errors.join(' ')).toContain('ít nhất một trường');
    expect(diff(s, opts({ name: '' })).errors.join(' ')).toContain('từng trường đã bật');
  });
  it('does not match a changed ID/code or duplicated inventory ID', () => {
    const s = setup();
    for (const tasks of [[oldTask({ code: 'OTHER' })], [oldTask(), oldTask({ code: 'OTHER' })]]) {
      const result = buildImportDiff(s.report, s.proposal, snapshot(tasks), opts());
      expect(result.rows[0].state).toBe('blocked'); expect(result.rows[0].fields).toEqual([]);
    }
  });
  it('flags absent/changed parent relationships without proposing a reparent operation', () => {
    const result = diff(setup(snapshot([oldTask({ parentInternalId: '999' })])));
    expect(result.rows[0].state).toBe('review');
    expect(result.rows[0].issues.join(' ')).toContain('chưa xác nhận quan hệ');
    expect(result.rows[0].fields.every((f) => f.field === 'name')).toBe(true);
  });
  it('flags a No. target currently occupied by another task', () => {
    const result = diff(setup(snapshot([oldTask(), oldTask({ internalId: '20', code: 'P.B', sourceNo: '1' })])), opts({ sourceNo: 'No.' }));
    expect(result.rows[0].state).toBe('review'); expect(result.rows[0].issues.join(' ')).toContain('No. đích');
  });
  it('validates resulting date pairs against unchanged dates but ignores unrelated fields', () => {
    const result = diff(setup(snapshot(), raw(['2026-11-01', null, null])), opts({ plannedStart: 'Bắt đầu' }));
    expect(result.rows[0].state).toBe('review'); expect(result.rows[0].issues.join(' ')).toContain('sau ngày kết thúc');
    expect(diff(setup(snapshot([oldTask({ plannedEnd: '2026-01-01' })]))).rows[0].state).toBe('changed');
    const missingEnd = diff(setup(snapshot([oldTask({ plannedEnd: undefined })]), raw(['2026-10-02', null, null])), opts({ plannedStart: 'Bắt đầu' }));
    expect(missingEnd.rows[0].state).toBe('review');
    expect(missingEnd.rows[0].issues.join(' ')).toContain('Chưa đủ hai mốc ngày');
  });
  it('labels first-time tasks as drafts and never synthesizes existing IDs', () => {
    const report = reviewProjectImport(raw(), null);
    const proposal = proposeImportNormalization(report, null, { mode: 'initial', prefix: 'P', firstTaskLevel: 1, packageLevel: 2, mappings: '', newRows: '' });
    const result = buildImportDiff(report, proposal, null, opts());
    expect(result.rows[0].state).toBe('draft'); expect(result.rows[0].internalId).toBeNull();
    expect(result.rows[0].fields[0].state).toBe('draft');
  });
  it('preserves source data and does not include unselected or protected fields', () => {
    const s = setup(); const before = JSON.stringify(s);
    const result = diff(s);
    expect(result.rows[0].fields.map((f) => f.field)).toEqual(['name']);
    expect(JSON.stringify(s)).toBe(before);
    expect(diff(s, opts({ name: 'Missing' })).rows[0].fields[0].state).toBe('unknown');
  });
});

it('parses only unambiguous valid calendar dates without discarding nonzero times', () => {
  for (const value of ['31/02/2026', '2026-02-29', '10/20/2026', '2026-01-01T01:00:00Z', 46000, new Date('2026-01-01T01:00:00Z')]) expect(importCalendarDate(value)).toBeNull();
  expect(importCalendarDate('29/02/2024')).toBe('2024-02-29');
  expect(importCalendarDate(new Date('2026-10-01T00:00:00Z'))).toBe('2026-10-01');
});
