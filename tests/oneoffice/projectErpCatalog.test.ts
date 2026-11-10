import { describe, expect, it } from 'vitest';
import { erpPackageNameKey, indexErpPackageNames, indexErpPackages, parseProjectErpCatalog, suggestErpPackages } from '../../src/oneoffice/projectErpCatalog';
import { proposeImportNormalization } from '../../src/oneoffice/projectImportNormalization';
import { reviewProjectImport } from '../../src/oneoffice/projectImportReview';
import type { OneOfficeProjectInventory } from '../../src/oneoffice/types';

const catalog = () => parseProjectErpCatalog({ sheet: 'TDR', data: [
  ['1. Danh mục đối tượng', null], ['Mã', 'Tên'], ['PRIVATE', 'Private person'],
  ['2. Danh mục gói thầu', null], ['Mã ', 'Tên'],
  ['TV.001', 'Tư vấn quy hoạch'], ['TV.002', 'Tư vấn quy hoạch'], ['KS.001', 'Khảo sát địa hình'],
  ['17. Danh mục hạng mục', null], ['Mã', 'Tên'], ['P', 'Dự án P'],
  ['18. Danh mục công đoạn sản xuất', null], ['Mã', 'Tên'], ['OTHER', 'Not a segment code'],
] });
const report = (packageCode: string | null = null) => reviewProjectImport({ sheet: 'Tasks', data: [
  ['No.', 'Mã gói thầu/Mã CV', 'Nội dung công việc'],
  ['1.1.1', 'ITEM', 'Hạng mục'], ['1.1.1.1', packageCode, 'Tư vấn quy hoạch'], ['1.1.1.1.1', null, 'Child'],
] }, null);
const options = { mode: 'initial' as const, prefix: 'P', firstTaskLevel: 3, packageLevel: 4, mappings: '', newRows: '' };

describe('ERP catalog parsing and suggestions', () => {
  it('retains only supported sections and their provenance, not personnel or other catalogs', () => {
    const result = catalog();
    expect(result.packages).toHaveLength(3); expect(result.projects).toHaveLength(1);
    expect(result.packages[0]).toEqual({ code: 'TV.001', name: 'Tư vấn quy hoạch', row: 6 });
    expect(JSON.stringify(result)).not.toContain('PRIVATE'); expect(JSON.stringify(result)).not.toContain('OTHER');
  });
  it('does not treat generic two-column lists or personnel as a package catalog', () => {
    for (const data of [[['Mã', 'Tên'], ['A', 'Work']], [['1. Danh mục nhân viên'], ['Mã', 'Tên'], ['A', 'Person']]]) {
      expect(() => parseProjectErpCatalog({ sheet: 'S', data })).toThrow('Không tìm thấy');
    }
  });
  it('deduplicates exact records but preserves code conflicts and duplicate names for manual review', () => {
    const result = parseProjectErpCatalog({ sheet: 'S', data: [
      ['2.Danh mục gói thầu'], ['Mã', 'Tên'], ['A', 'Same'], ['A', 'Same'], ['B', 'Same'], ['A', 'Different'], ['bad code', 'Skip'],
    ] });
    expect(result.packages).toHaveLength(3);
    expect(indexErpPackages(result).get('A')).toHaveLength(2);
    expect(result.warnings.join(' ')).toContain('mã gói thầu ứng với nhiều tên');
    expect(result.warnings.join(' ')).toContain('bị bỏ qua');
    expect(suggestErpPackages(result, 'Same').entries).toHaveLength(2);
  });
  it('ranks accent-insensitive name/code suggestions, caps results and never assigns a selection', () => {
    expect(suggestErpPackages(catalog(), 'TU VAN QUY HOACH').entries.map((r) => r.code)).toEqual(['TV.001', 'TV.002']);
    expect(suggestErpPackages(catalog(), 'KS.001').entries[0].code).toBe('KS.001');
    expect(suggestErpPackages(catalog(), '').entries).toHaveLength(0);
    const many = { ...catalog(), packages: Array.from({ length: 30 }, (_, row) => ({ code: `GT.${row}`, name: 'Công việc', row })) };
    expect(suggestErpPackages(many, 'Công việc').entries).toHaveLength(20);
    expect(suggestErpPackages(many, 'Công việc').total).toBe(30);
  });
  it('enforces sheet limits', () => {
    expect(() => parseProjectErpCatalog({ sheet: 'S', data: Array.from({ length: 100001 }, () => []) })).toThrow('giới hạn');
    expect(() => parseProjectErpCatalog({ sheet: 'S', data: [Array(21).fill(null)] })).toThrow('giới hạn');
  });
  it('prioritizes planning scope and scale rather than generic consultancy/design words', () => {
    const result = { ...catalog(), packages: [
      { code: 'GENERAL', name: 'Tư vấn thiết kế ý tưởng QH tổng MB', row: 1 },
      { code: 'CONSTRUCTION', name: 'Tư vấn thiết kế triển khai BVTC', row: 2 },
      { code: 'WRONG_SCALE', name: 'Tư vấn chỉnh sửa quy hoạch 1/500', row: 3 },
      ...['031', '032', '033'].map((suffix, i) => ({ code: `TV.TC.02.${suffix}`, name: 'Tư vấn chỉnh sửa quy hoạch 1/2000 để cập nhật QHPK địa phương', row: i + 4 })),
    ] };
    const suggested = suggestErpPackages(result, 'Tư vấn thiết kế QH 1/2000 điều chỉnh');
    expect(suggested.entries.map((r) => r.code)).toEqual(['TV.TC.02.031', 'TV.TC.02.032', 'TV.TC.02.033']);
    expect(suggested.total).toBe(3);
    const names = indexErpPackageNames(result);
    expect(names.get(erpPackageNameKey(suggested.entries[0]))?.size).toBe(3);
    expect(suggestErpPackages(result, 'TV.TC.02.032').entries[0].code).toBe('TV.TC.02.032');
    expect(suggestErpPackages(result, 'TV.TC.02').entries).toHaveLength(3);
  });
  it('does not treat 1/20000 as 1/2000 or equate names that differ by accents', () => {
    const result = { ...catalog(), packages: [
      { code: 'A', name: 'Tư vấn quy hoạch 1/20000', row: 1 },
      { code: 'B', name: 'Tư vấn quy hoạch 1/2000', row: 2 },
      { code: 'C', name: 'Mở đường', row: 3 }, { code: 'D', name: 'Mơ đường', row: 4 },
    ] };
    expect(suggestErpPackages(result, 'quy hoạch 1/2000').entries.map((r) => r.code)).toEqual(['B']);
    const names = indexErpPackageNames(result);
    expect(names.get(erpPackageNameKey(result.packages[2]))?.size).toBe(1);
    expect(names.get(erpPackageNameKey(result.packages[3]))?.size).toBe(1);
  });
  it('groups distinct ERP codes with the same description without merging their identities', () => {
    const result = { ...catalog(), packages: [
      { code: 'A', name: 'A: Tư vấn quy hoạch', row: 1 },
      { code: 'B', name: 'Tư vấn quy hoạch', row: 2 },
      { code: 'C', name: 'Tư vấn quy hoạch - Hợp đồng 123', row: 3 },
    ] };
    expect(indexErpPackageNames(result).get(erpPackageNameKey(result.packages[0]))).toEqual(new Set(['A', 'B']));
    expect(result.packages).toHaveLength(3);
  });
});

describe('ERP-aware code proposals', () => {
  it('accepts exact case-sensitive package membership and keeps ERP code intact under the item', () => {
    const result = proposeImportNormalization(report('KS.001'), null, options, catalog());
    expect(result.rows.every((r) => r.kind === 'proposed')).toBe(true);
    expect(result.rows[1].erpCode).toBe('KS.001');
    expect(result.rows[1].code).toBe('P.ITEM.KS.001');
    expect(result.rows[2].code).toBe('P.ITEM.KS.001.01');
    const lower = proposeImportNormalization(report('ks.001'), null, options, catalog());
    expect(lower.rows[1].kind).toBe('blocked');
  });
  it('never uses a name suggestion automatically for missing or invalid codes', () => {
    for (const code of [null, 'UNKNOWN', '1']) {
      const result = proposeImportNormalization(report(code), null, options, catalog());
      expect(result.rows[1].kind).toBe('blocked'); expect(result.rows[2].kind).toBe('blocked');
    }
  });
  it('uses an explicitly selected ERP component without altering the source file and propagates to children', () => {
    const input = report(); const before = JSON.stringify(input);
    const result = proposeImportNormalization(input, null, { ...options, erpPackages: { 3: 'TV.002' } }, catalog());
    expect(result.rows[1].code).toBe('P.ITEM.TV.002');
    expect(result.rows[2].code).toBe('P.ITEM.TV.002.01');
    expect(result.rows.every((r) => r.kind === 'proposed')).toBe(true);
    expect(JSON.stringify(input)).toBe(before);
  });
  it('validates an explicit full code using known parent boundaries, not dot depth', () => {
    const result = proposeImportNormalization(report(), null, { ...options, mappings: '3 P.ITEM.KS.001' }, catalog());
    expect(result.rows[1].kind).toBe('proposed'); expect(result.rows[1].erpCode).toBe('KS.001');
    const conflict = proposeImportNormalization(report(), null, { ...options, mappings: '3 P.ITEM.KS.001', erpPackages: { 3: 'TV.001' } }, catalog());
    expect(conflict.rows[1].kind).toBe('blocked');
    expect(conflict.rows[1].issues.join(' ')).toContain('khác mã ERP');
  });
  it('blocks mismatched projects, stale selections, non-package selections and conflicting catalog codes', () => {
    for (const patch of [{ prefix: 'UNKNOWN' }, { erpPackages: { 999: 'TV.001' } }, { erpPackages: { 2: 'TV.001' } }, { erpPackages: { 3: 'OLD' } }]) {
      expect(proposeImportNormalization(report(), null, { ...options, ...patch }, catalog()).errors.length).toBeGreaterThan(0);
    }
    const duplicate = catalog(); duplicate.packages.push({ code: 'TV.001', name: 'Different', row: 100 });
    expect(proposeImportNormalization(report('TV.001'), null, options, duplicate).rows[1].kind).toBe('blocked');
    expect(proposeImportNormalization(report(), null, { ...options, erpPackages: { 3: 'TV.001' } }, duplicate).errors.length).toBeGreaterThan(0);
  });
  it('cannot use ERP choices to rename an existing 1Office record', () => {
    const input = report('P.ITEM.KS.001'); input.rows[0].code = 'P.ITEM';
    const old = { collection: { complete: true }, tasks: [{ code: 'P.ITEM', internalId: '1' }, { code: 'P.ITEM.KS.001', internalId: '2' }] } as OneOfficeProjectInventory;
    const result = proposeImportNormalization(input, old, { ...options, mode: 'update', erpPackages: { 3: 'TV.001' } }, catalog());
    expect(result.rows[1].kind).toBe('blocked');
    expect(result.rows[1].code).toBe('P.ITEM.KS.001'); expect(result.rows[1].internalId).toBeNull();
    const unchanged = proposeImportNormalization(input, old, { ...options, mode: 'update' }, catalog());
    expect(unchanged.rows[1].kind).toBe('existing'); expect(unchanged.rows[1].internalId).toBe('2');
  });
});

it.skipIf(!process.env.TERA_ERP_TEST_FILE)('reads the provided ERP workbook selected TDR sheet without retaining unrelated catalogs', async () => {
  const { readSheet } = await import('read-excel-file/node');
  const data = await readSheet(process.env.TERA_ERP_TEST_FILE!, 'TDR');
  const result = parseProjectErpCatalog({ sheet: 'TDR', data });
  expect(result.projects).toHaveLength(35); expect(result.packages).toHaveLength(2243);
  expect(result.projects.find((r) => r.code === 'CNHY')?.row).toBe(45573);
  expect(indexErpPackages(result).get('TV.TC.01.001')).toHaveLength(1);
  const suggestions = suggestErpPackages(result, 'Tư vấn thiết kế QH 1/2000 điều chỉnh');
  expect(suggestions.entries.slice(0, 3).map((r) => r.code)).toEqual(['TV.TC.02.031', 'TV.TC.02.032', 'TV.TC.02.033']);
  expect(suggestions.total).toBeLessThan(485);
  expect(indexErpPackageNames(result).get(erpPackageNameKey(suggestions.entries[0]))?.size).toBe(3);
  expect(result.warnings).toEqual([]);
});
