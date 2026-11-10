import { describe, expect, it } from 'vitest';
import { buildNoUpdatePlan } from '../../src/oneoffice/noUpdatePlan';
import type { OneOfficeProjectInventory } from '../../src/oneoffice/types';

function inventory(): OneOfficeProjectInventory {
  return {
    schemaVersion: 1, source: '1office-project-api', capturedAt: '2026-10-14T00:00:00Z',
    project: { type: 'project', internalId: '870', name: 'test', url: 'https://office.meygroup.vn/apps/work-project-project/view?ID=870' },
    collection: { complete: true, expectedTaskCount: 3, loadedTaskCount: 3, warnings: [] },
    tasks: ['A', 'B', 'C'].map((code, i) => ({
      code, internalId: String(i + 1), name: code, no: `1.${i + 1}`, sourceOrder: i,
      url: '', depth: 2, parentNo: '1', parentInternalId: '100',
    })),
  };
}

describe('No. bulk preview (no writes)', () => {
  it('uses verified raw blanks rather than a generated outline when planning writes', () => {
    const source = inventory(); source.tasks[2].sourceNo = '';
    const plan = buildNoUpdatePlan(source, 'C\t1.3');
    expect(plan.changes[0]).toMatchObject({ before: '', after: '1.3' });
    expect(plan.errors).toEqual([]);
  });
  it('matches by code, preserves ID and leaves the source untouched', () => {
    const source = inventory();
    const before = JSON.stringify(source);
    const plan = buildNoUpdatePlan(source, 'C\t1.4');
    expect(plan.errors).toEqual([]);
    expect(plan.changes).toEqual([{ code: 'C', internalId: '3', name: 'C', before: '1.3', after: '1.4' }]);
    expect(plan.untouched).toBe(2);
    expect(JSON.stringify(source)).toBe(before);
  });
  it('accepts a simultaneous swap, not sequential overwrites by No.', () => {
    expect(buildNoUpdatePlan(inventory(), 'B\t1.3\nC\t1.2').errors).toEqual([]);
  });
  it('detects collisions with untouched tasks', () => {
    expect(buildNoUpdatePlan(inventory(), 'B\t1.3').errors.join()).toContain('No. 1.3 bị trùng');
  });
  it('resolves an existing duplicate when the supplied mapping is unambiguous', () => {
    const source = inventory();
    source.tasks[2].no = '1.2';
    expect(buildNoUpdatePlan(source, 'C\t1.3').errors).toEqual([]);
  });
  it('does not treat unknown codes as existing tasks or create them', () => {
    const plan = buildNoUpdatePlan(inventory(), 'NEW\t1.4');
    expect(plan.changes).toEqual([]);
    expect(plan.additions).toEqual([{ code: 'NEW', no: '1.4' }]);
    expect(plan.errors.join()).toContain('Chưa hỗ trợ tạo');
  });
  it('rejects duplicate input codes even if the numbers agree', () => {
    expect(buildNoUpdatePlan(inventory(), 'A\t1.1\nA\t1.1').errors.join()).toContain('bị lặp');
  });
  it('rejects duplicate source codes and invalid or duplicate IDs', () => {
    const source = inventory();
    source.tasks[1].code = 'A';
    source.tasks[1].internalId = '1';
    const plan = buildNoUpdatePlan(source, 'A\t1.4');
    expect(plan.errors.join()).toContain('Mã A trùng');
    expect(plan.errors.join()).toContain('ID công việc');
    expect(plan.rows).toEqual([]);
  });
  it('keeps No. as text and allows deep numbering without changing hierarchy', () => {
    const source = inventory();
    const plan = buildNoUpdatePlan(source, 'A\t1.2.3.4.5.6.7.8.9.10');
    expect(plan.errors).toEqual([]);
    expect(source.tasks[0].parentInternalId).toBe('100');
    expect(source.tasks[0].depth).toBe(2);
  });
  it('rejects malformed input, headers and ambiguous numeric formats', () => {
    for (const value of ['', 'Mã công việc\tNo.', 'A\t01.2', 'A\t1..2', 'A\t1.0', 'A\t1.2\textra', 'A,1.2']) {
      expect(buildNoUpdatePlan(inventory(), value).errors.length).toBeGreaterThan(0);
    }
  });
  it('does not assume incomplete inventory is safe to write', () => {
    const source = inventory();
    source.collection.complete = false;
    expect(buildNoUpdatePlan(source, 'A\t1.4').errors.join()).toContain('chưa được xác nhận đầy đủ');
  });
  it('counts unchanged rows and supports Windows line endings', () => {
    const plan = buildNoUpdatePlan(inventory(), 'A\t1.1\r\nC\t1.4\r\n');
    expect(plan.errors).toEqual([]);
    expect(plan.unchanged).toBe(1);
    expect(plan.changes).toHaveLength(1);
  });
});
