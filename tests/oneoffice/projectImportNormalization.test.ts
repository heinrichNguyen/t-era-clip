import { describe, expect, it } from 'vitest';
import { proposeImportNormalization, type NormalizationOptions } from '../../src/oneoffice/projectImportNormalization';
import { reviewProjectImport } from '../../src/oneoffice/projectImportReview';
import type { OneOfficeProjectInventory } from '../../src/oneoffice/types';

const defaults: NormalizationOptions = { mode: 'initial', prefix: 'P', firstTaskLevel: 3, packageLevel: 4, mappings: '', newRows: '' };
const report = (rows: (string | number | null)[][]) => reviewProjectImport({ sheet: 'Sheet1', data: [['No.', 'Mã gói thầu/Mã CV', 'Nội dung công việc'], ...rows] }, null);
const inventory = (codes: string[], complete = true) => ({ project: { internalId: '1', name: 'Project' }, collection: { complete }, tasks: codes.map((code, i) => ({ code, internalId: String(i + 1) })) } as OneOfficeProjectInventory);
const source = () => report([['1', 'P', 'Project'], ['1.1', 'Phase', 'Phase'], ['1.1.1', 'ITEM', 'Item'], ['1.1.1.1', 'PKG.001', 'Package'], ['1.1.1.1.1', 1, 'Child'], ['1.1.1.1.1.1', null, 'Grandchild']]);

describe('local code normalization proposals', () => {
  it('builds the configured code hierarchy without modifying No. or source values', () => {
    const input = source(); const before = JSON.stringify(input);
    const result = proposeImportNormalization(input, null, defaults);
    expect(result.rows.map((r) => r.code)).toEqual(['', '', 'P.ITEM', 'P.ITEM.PKG.001', 'P.ITEM.PKG.001.01', 'P.ITEM.PKG.001.01.01']);
    expect(result.rows[5].parentCode).toBe('P.ITEM.PKG.001.01');
    expect(result.rows.slice(0,2).every((r) => r.kind === 'excluded')).toBe(true);
    expect(result.rows.slice(2).every((r) => r.kind === 'proposed')).toBe(true);
    expect(JSON.stringify(input)).toBe(before);
  });
  it('keeps existing IDs and codes in update mode even when file order/No. changes', () => {
    const input = report([['1.1.1', 'P.ITEM', 'I'], ['1.1.1.1', 'P.PKG', 'P'], ['1.1.1.1.2', 'P.PKG.02', 'B'], ['1.1.1.1.1', 'P.PKG.01', 'A']]);
    const result = proposeImportNormalization(input, inventory(['P.ITEM','P.PKG','P.PKG.01','P.PKG.02']), { ...defaults, mode: 'update' });
    expect(result.rows.every((r) => r.kind === 'existing')).toBe(true);
    expect(result.rows[2].code).toBe('P.PKG.02');
    expect(result.rows[2].internalId).toBe('4');
  });
  it('preserves dotted project, item and ERP package components through deep child levels', () => {
    const input = report([
      ['1.1.1', 'HM.02', 'Item'], ['1.1.1.1', 'TV.TC.01.001', 'Package'],
      ['1.1.1.1.1', 1, 'Parent'], ['1.1.1.1.1.1', 1, 'Child'],
      ['1.1.1.1.1.1.1', 1, 'Grandchild'], ['1.1.1.1.1.1.1.1', 1, 'Great grandchild'],
    ]);
    const result = proposeImportNormalization(input, null, { ...defaults, prefix: 'HP.HH1' });
    expect(result.rows.map((r) => r.code)).toEqual([
      'HP.HH1.HM.02', 'HP.HH1.HM.02.TV.TC.01.001',
      'HP.HH1.HM.02.TV.TC.01.001.01', 'HP.HH1.HM.02.TV.TC.01.001.01.01',
      'HP.HH1.HM.02.TV.TC.01.001.01.01.01', 'HP.HH1.HM.02.TV.TC.01.001.01.01.01.01',
    ]);
    expect(result.rows.every((r) => r.kind === 'proposed')).toBe(true);
    expect(result.rows[2].group).toBe('Segment');
  });
  it('distinguishes the same package component under different items', () => {
    const result = proposeImportNormalization(report([
      ['1.1.1', 'A', 'A'], ['1.1.1.1', 'GT.001', 'Package A'],
      ['1.1.2', 'B', 'B'], ['1.1.2.1', 'GT.001', 'Package B'],
    ]), null, defaults);
    expect(result.rows.map((r) => r.code)).toEqual(['P.A', 'P.A.GT.001', 'P.B', 'P.B.GT.001']);
    expect(result.rows.every((r) => r.kind === 'proposed')).toBe(true);
  });
  it('does not double-prefix complete codes with the expected ancestry', () => {
    const result = proposeImportNormalization(report([
      ['1.1.1', 'P.ITEM', 'Item'], ['1.1.1.1', 'P.ITEM.GT.001', 'Package'],
      ['1.1.1.1.1', 'P.ITEM.GT.001.07', 'Child'],
    ]), null, defaults);
    expect(result.rows.map((r) => r.code)).toEqual(['P.ITEM', 'P.ITEM.GT.001', 'P.ITEM.GT.001.07']);
    expect(result.rows.every((r) => r.kind === 'proposed')).toBe(true);
  });
  it('blocks inconsistent explicit full codes and their children rather than silently adding an item', () => {
    const input = report([
      ['1.1.1', 'ITEM', 'Item'], ['1.1.1.1', 'PKG', 'Package'], ['1.1.1.1.1', 1, 'Child'],
    ]);
    for (const mappings of ['3 P.PKG', '3 P.ITEMX.PKG', '2 OTHER.ITEM']) {
      const result = proposeImportNormalization(input, null, { ...defaults, mappings });
      expect(result.rows[1].kind).toBe('blocked');
      expect(result.rows[2].kind).toBe('blocked');
      expect(result.rows.flatMap((r) => r.issues).join(' ')).toContain('Mã đầy đủ phải nối tiếp');
    }
    const legacyFull = proposeImportNormalization(report([
      ['1.1.1', 'ITEM', 'I'], ['1.1.1.1', 'P.PKG', 'P'],
    ]), null, defaults);
    expect(legacyFull.rows[1].kind).toBe('blocked');
    expect(legacyFull.rows[1].code).toBe('P.PKG');
  });
  it('propagates a manually supplied full package code to its descendants', () => {
    const result = proposeImportNormalization(report([
      ['1.1.1', 'ITEM', 'I'], ['1.1.1.1', null, 'P'], ['1.1.1.1.1', null, 'Child'],
    ]), null, { ...defaults, mappings: '3 P.ITEM.GT.001' });
    expect(result.rows[2].code).toBe('P.ITEM.GT.001.01');
    expect(result.rows.every((r) => r.kind === 'proposed')).toBe(true);
  });
  it('extends intermediate parent levels without using code dot counts', () => {
    const result = proposeImportNormalization(report([
      ['1.1.1', 'ITEM', 'I'], ['1.1.1.1', 'MID.01', 'Intermediate'], ['1.1.1.1.1', 'GT.001', 'P'],
    ]), null, { ...defaults, packageLevel: 5 });
    expect(result.rows[2].code).toBe('P.ITEM.MID.01.GT.001');
    expect(result.rows.every((r) => r.kind === 'proposed')).toBe(true);
  });
  it('keeps legacy codes but requires new packages to extend the existing item code in update mode', () => {
    const input = report([['1.1.1', 'P.ITEM', 'I'], ['1.1.1.1', 'GT.001', 'New package']]);
    const old = inventory(['P.ITEM']);
    const result = proposeImportNormalization(input, old, { ...defaults, mode: 'update', newRows: '3' });
    expect(result.rows[0].kind).toBe('existing');
    expect(result.rows[0].internalId).toBe('1');
    expect(result.rows[1].code).toBe('P.ITEM.GT.001');
    expect(result.rows[1].kind).toBe('proposed');
  });
  it('does not create codes for unidentified old tasks unless explicitly marked new', () => {
    const input = report([['1.1.1','P.ITEM','I'],['1.1.1.1','P.PKG','P'],['1.1.1.1.1',1,'New']]);
    const old = inventory(['P.ITEM','P.PKG','P.PKG.01','P.PKG.02']);
    const blocked = proposeImportNormalization(input, old, { ...defaults, mode: 'update' });
    expect(blocked.rows[2].kind).toBe('blocked'); expect(blocked.rows[2].code).toBe('');
    const proposed = proposeImportNormalization(input, old, { ...defaults, mode: 'update', newRows: '4' });
    expect(proposed.rows[2].code).toBe('P.PKG.03'); expect(proposed.rows[2].kind).toBe('proposed');
    const mapped = proposeImportNormalization(input, old, { ...defaults, mode: 'update', mappings: '4\tP.PKG.02' });
    expect(mapped.rows[2].kind).toBe('existing'); expect(mapped.rows[2].internalId).toBe('4');
  });
  it('never changes a matched source code through an override or marks it as new', () => {
    const input = report([['1.1.1','P.ITEM','I']]);
    const old = inventory(['P.ITEM', 'P.OTHER']);
    for (const patch of [{ mappings: '2 P.OTHER' }, { newRows: '2' }]) {
      const result = proposeImportNormalization(input, old, { ...defaults, mode: 'update', ...patch });
      expect(result.rows[0].kind).toBe('blocked'); expect(result.rows[0].internalId).toBeNull();
    }
  });
  it('reserves explicitly supplied later codes before allocating suffixes', () => {
    const input = report([['1.1.1','ITEM','I'],['1.1.1.1','PKG','P'],['1.1.1.1.1',1,'New'],['1.1.1.1.2','P.ITEM.PKG.01','Specified']]);
    const result = proposeImportNormalization(input, null, defaults);
    expect(result.rows[2].code).toBe('P.ITEM.PKG.02');
    expect(result.rows[3].code).toBe('P.ITEM.PKG.01');
  });
  it('blocks duplicate codes and their descendants instead of choosing a parent arbitrarily', () => {
    const input = report([['1.1.1','A','A'],['1.1.2','A','B'],['1.1.2.1','B','C'],['1.1.2.1.1',1,'D']]);
    const result = proposeImportNormalization(input, null, defaults);
    expect(result.rows.every((r) => r.kind === 'blocked')).toBe(true);
  });
  it('requires a complete system inventory for update mode and blocks duplicate system codes', () => {
    for (const old of [null, inventory(['P.ITEM'], false), inventory(['P.ITEM','P.ITEM'])]) {
      const result = proposeImportNormalization(report([['1.1.1','P.ITEM','I']]), old, { ...defaults, mode: 'update' });
      expect(result.rows[0].kind).toBe('blocked');
    }
  });
  it('does not re-prefix an already existing source code in initial mode', () => {
    const result = proposeImportNormalization(report([['1.1.1','ITEM','I']]), inventory(['ITEM']), defaults);
    expect(result.rows[0].kind).toBe('blocked');
    expect(result.rows[0].issues.join(' ')).toContain('cần dùng chế độ cập nhật');
  });
  it('requires explicit missing package codes and valid preceding parents', () => {
    const input = report([['1.1.1',null,'Missing'],['1.1.1.1','PKG','Child'],['1.2.1.1','ORPHAN','O']]);
    const result = proposeImportNormalization(input, null, defaults);
    expect(result.rows.every((r) => r.kind === 'blocked')).toBe(true);
    const fixed = proposeImportNormalization(input, null, { ...defaults, mappings: '2 P.ITEM' });
    expect(fixed.rows[0].kind).toBe('proposed'); expect(fixed.rows[1].parentCode).toBe('P.ITEM');
  });
  it('reports invalid settings, duplicate mappings and nonexistent new rows', () => {
    for (const patch of [{ prefix: 'bad prefix' }, { firstTaskLevel: 5, packageLevel: 4 }, { mappings: '4 A\n4 B' }, { mappings: '999 CODE' }, { newRows: '999' }]) {
      const result = proposeImportNormalization(source(), null, { ...defaults, ...patch });
      expect(result.errors.length).toBeGreaterThan(0);
      expect(result.rows.every((r) => r.kind === 'blocked' || r.kind === 'excluded')).toBe(true);
    }
  });
  it('allows legacy letter outlines without stripping them or claiming hierarchy on 1Office', () => {
    const result = proposeImportNormalization(report([['1.1.2a','A','A'],['1.1.2a.1','B','B']]), null, defaults);
    expect(result.rows[0].no).toBe('1.1.2a'); expect(result.rows[1].parentCode).toBe('P.A');
    expect(result.warnings.join(' ')).toContain('không thay đổi quan hệ');
  });
});

it.skipIf(!process.env.TERA_IMPORT_TEST_FILE)('proposes locally against the real raw MS Project export without mutating or assigning system IDs', async () => {
  const { default: readExcel } = await import('read-excel-file/node');
  const sheets = await readExcel(process.env.TERA_IMPORT_TEST_FILE!);
  const review = reviewProjectImport(sheets[0], null);
  const before = JSON.stringify(review);
  const proposal = proposeImportNormalization(review, null, { ...defaults, prefix: 'CNHY' });
  expect(proposal.rows).toHaveLength(773);
  expect(proposal.rows.every((r) => r.internalId === null)).toBe(true);
  expect(proposal.rows.some((r) => r.kind === 'blocked')).toBe(true);
  expect(new Set(proposal.rows.filter((r) => r.kind === 'proposed').map((r) => r.code)).size).toBe(proposal.rows.filter((r) => r.kind === 'proposed').length);
  expect(JSON.stringify(review)).toBe(before);
});
