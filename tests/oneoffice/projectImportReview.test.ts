import { describe, expect, it } from 'vitest';
import { reviewProjectImport, type ImportSheet } from '../../src/oneoffice/projectImportReview';
import type { OneOfficeProjectInventory } from '../../src/oneoffice/types';

const header = ['No.', 'Text5 (WBS gốc)', 'Mã gói thầu/Mã CV', 'Nội dung công việc', 'MQH', 'Bắt đầu thực tế'];
const sheet = (rows: ImportSheet['data']): ImportSheet => ({ sheet: 'Sheet1', data: [header, ...rows] });
const inventory = (codes: string[], complete = true) => ({
  project: { internalId: '870' }, collection: { complete },
  tasks: codes.map((code, i) => ({ code, internalId: String(i + 1) })),
} as OneOfficeProjectInventory);

describe('local MS Project pre-import review', () => {
  it('matches only unique exact codes in a complete source and never uses name or No.', () => {
    const result = reviewProjectImport(sheet([
      ['1', '1', 'PACKAGE', 'Parent'], ['1.1', '', 'CODE-A', 'Child'], ['1.2', '', 'code-a', 'Child'],
    ]), inventory(['PACKAGE', 'CODE-A']));
    expect(result.rows.map((r) => r.match)).toEqual(['matched', 'matched', 'unmatched']);
    expect(result.rows[1].internalId).toBe('2');
    expect(result.rows[2].internalId).toBeNull();
    expect(result.levels).toEqual({ 1: 1, 2: 2 });
    expect(result.warnings.join(' ')).toContain('MQH');
    expect(result.warnings.join(' ')).toContain('không xác minh công thức');
  });
  it('blocks missing, local numeric and duplicate codes including duplicates on 1Office', () => {
    const result = reviewProjectImport(sheet([
      ['1', '', '', 'Missing'], ['2', '', 1, 'Numeric'], ['3', '', 'DUP', 'A'], ['4', '', 'DUP', 'B'],
      ['5', '', 'SERVER', 'C'],
    ]), inventory(['SERVER', 'SERVER']));
    expect(result.rows.every((r) => r.match === 'blocked' && r.internalId === null)).toBe(true);
    expect(result.rows[4].issues).toContain('Mã bị trùng trong danh sách 1Office');
  });
  it('accepts legacy alphanumeric outlines for inspection but requests numbering review', () => {
    const result = reviewProjectImport(sheet([
      ['1', '', 'ROOT', 'Root'], ['1.2a', '', 'PARENT', 'P'], ['1.2a.1', '', 'CHILD', 'C'],
      ['13.2a', '', 'BAD', 'Bad'], ['1.3.1', '', 'EARLY', 'Early'], ['1.3', '', 'LATE', 'Late'],
    ]), null);
    expect(result.rows[2].level).toBe(3);
    expect(result.rows[2].parentCandidate).toBe('1.2a');
    expect(result.rows[3].issues).toContain('Chưa xác định duy nhất dòng cha theo No.');
    expect(result.rows[4].issues).toContain('Dòng cha theo No. nằm sau công việc, cần kiểm tra cấp');
  });
  it('does not classify codes against an incomplete or absent inventory', () => {
    for (const source of [null, inventory(['ROOT'], false)]) {
      expect(reviewProjectImport(sheet([['1', '', 'ROOT', 'R']]), source).rows[0].match).toBe('unavailable');
    }
  });
  it('rejects missing/duplicate headers and oversized sheets', () => {
    expect(() => reviewProjectImport({ sheet: 'X', data: [['No.']] }, null)).toThrow('Thiếu cột');
    expect(() => reviewProjectImport({ sheet: 'X', data: [[...header, 'No.']] }, null)).toThrow('trùng');
    expect(() => reviewProjectImport(sheet(Array(10001).fill(['1', '', 'CODE', 'N'])), null)).toThrow('10.000');
  });
  it('retains dates, blank cells, dependencies and text without evaluating them or mutating inputs', () => {
    const date = new Date('2026-10-01T00:00:00Z');
    const input = sheet([['1', null, 'ROOT', '<img src=x onerror=alert(1)>', '5SS+10 d', date]]);
    const before = JSON.stringify(input);
    const report = reviewProjectImport(input, null);
    expect(report.rows[0].values.MQH).toBe('5SS+10 d');
    expect(report.rows[0].values['Bắt đầu thực tế']).toBe(date);
    expect(JSON.stringify(input)).toBe(before);
  });
});

it.skipIf(!process.env.TERA_IMPORT_TEST_FILE)('reads the supplied real workbook with the packaged Excel reader', async () => {
  const { default: readExcel } = await import('read-excel-file/node');
  const sheets = await readExcel(process.env.TERA_IMPORT_TEST_FILE!);
  const report = reviewProjectImport(sheets[0], null);
  expect(report.rows).toHaveLength(773);
  expect(report.rows.filter((r) => !r.code)).toHaveLength(143);
  expect(report.rows.filter((r) => !r.originalWbs)).toHaveLength(229);
  expect(report.rows.some((r) => r.no === '1.1.2a')).toBe(true);
  expect(Math.max(...report.rows.map((r) => r.level || 0))).toBe(9);
});
