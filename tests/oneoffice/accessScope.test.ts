// @vitest-environment jsdom
import { describe, expect, it, vi } from 'vitest';
import { activeAccessSnapshot, collectScopedAccessReport, selectAccessAccounts } from '../../src/oneoffice/accessScope';
import { reportFromAccessSnapshot, type AccessRecord } from '../../src/oneoffice/accessHistory';

const record = (id: number, account: string): AccessRecord => ({ id: String(id), account, name: account,
  timeRaw: '12:00:00 29/09/2026', timeIso: '2026-09-29T05:00:00.000Z' });

describe('access scope', () => {
  it('expands TDK/TDU groups, isolates members and independent companies, and never includes locked accounts', () => {
    const codes = ['TDK', 'TDH', 'TDP', 'TDC', 'TDW', 'TDU', 'TDU.CKKG', 'TDU.GNKG', 'TDU.VLXD', 'TDU.XMHT', 'TDG', 'TDT', 'TDR', 'TDU.OTHER', 'TDH.EXTRA'];
    const users = codes.map((companyCode) => ({ account: companyCode, companyCode, position: '', department: '' }));
    const personnel = { collectedAt: '', warnings: [], activeAccounts: [...codes, 'parentTDK', 'parentTDU'], users: [
      ...users, { account: 'parentTDK', companyCode: '', companyGroup: 'TDK', position: '', department: '' },
      { account: 'parentTDU', companyCode: '', companyGroup: 'TDU', position: '', department: '' },
      { account: 'LOCKED', companyCode: 'TDH', position: '', department: '' },
    ] };
    const pick = (company: string) => selectAccessAccounts(personnel, { company, position: '', department: '' });
    expect(pick('TDK')).toEqual(['TDK', 'TDH', 'TDP', 'TDC', 'TDW', 'parentTDK']);
    expect(pick('TDU')).toEqual(['TDU', 'TDU.CKKG', 'TDU.GNKG', 'TDU.VLXD', 'TDU.XMHT', 'parentTDU']);
    for (const company of ['TDH', 'TDP', 'TDC', 'TDW', 'TDU.CKKG', 'TDU.GNKG', 'TDU.VLXD', 'TDU.XMHT', 'TDG', 'TDT', 'TDR']) {
      expect(pick(company)).toEqual([company]);
    }
  });
  it('uses company hierarchy rather than guessing from R department names; combines conditions', () => {
    const personnel = { collectedAt: '', warnings: [], activeAccounts: ['A', 'B', 'C', 'D'], users: [
      { account: 'A', companyCode: 'TDR', company: 'Meyland', position: 'Giám đốc', department: 'R - Hạ tầng', departmentCode: 'TDR.HT' },
      { account: 'B', companyCode: 'OTHER', company: 'Other', position: 'TDR - Giám đốc', department: 'R - Hạ tầng' },
      { account: 'C', companyCode: 'TDR', position: 'Chuyên viên', department: 'R - Hạ tầng' },
      { account: 'D', position: '', department: '' },
    ] };
    expect(selectAccessAccounts(personnel, { company: 'tdr', position: 'giám đốc', department: 'TDR.HT' })).toEqual(['A']);
    expect(selectAccessAccounts(personnel, { company: 'Meyland', position: '', department: '' })).toEqual(['A']);
    expect(() => selectAccessAccounts({ ...personnel, users: [personnel.users[3]] }, { company: 'TDR', position: '', department: '' })).toThrow('Chưa xác minh');
  });
  it('removes locked and unknown accounts from old records, recalculates counts, and fails closed without status', () => {
    const snapshot = { collectedAt: '', recordCount: 3, oldestTime: 'old', records: [record(3, 'A'), record(2, 'LOCKED'), record(1, 'UNKNOWN')] };
    const personnel = { collectedAt: '', activeAccounts: ['A'], users: [], warnings: [] };
    const active = activeAccessSnapshot(snapshot, personnel);
    expect(active.records.map((row) => row.account)).toEqual(['A']);
    expect(active.recordCount).toBe(1);
    expect(reportFromAccessSnapshot(active, '2026-09-30').users.map((row) => row.account)).toEqual(['A']);
    expect(() => activeAccessSnapshot(snapshot, { collectedAt: '', users: [], warnings: [] })).toThrow('Chưa xác minh');
    expect(selectAccessAccounts({ ...personnel, users: [{ account: 'LOCKED', position: '', department: '', companyCode: 'TDR' }] },
      { company: 'TDR', position: '', department: '' })).toEqual([]);
  });
  it('keeps exact accounts only, preserves scope when changing date, and never scans empty selections', async () => {
    const read = vi.fn(async () => ({ total: 2, start: 1, end: 2, records: [record(2, 'AA'), record(1, 'A')] }));
    const result = await collectScopedAccessReport('2026-09-30', ['A'], read, new AbortController().signal, vi.fn(), 'TDR');
    expect(result.users.map((user) => user.account)).toEqual(['A']);
    expect(result.recordCount).toBe(1);
    expect(reportFromAccessSnapshot(result.snapshot!, '2026-09-28').scope).toBe('TDR');
    expect(reportFromAccessSnapshot(result.snapshot!, '2026-09-28').users).toHaveLength(0);
    read.mockClear();
    await expect(collectScopedAccessReport('2026-09-30', [], read, new AbortController().signal, vi.fn(), 'TDR')).rejects.toThrow('Không có tài khoản');
    expect(read).not.toHaveBeenCalled();
  });
  it('shares six network slots across accounts and pages, including rechecks', async () => {
    let active = 0; let maximum = 0;
    const accounts = Array.from({ length: 10 }, (_, i) => `A${i}`);
    const read = async (page: number, account: string) => {
      active++; maximum = Math.max(maximum, active);
      await new Promise((resolve) => setTimeout(resolve, 1)); active--;
      const base = Number(account.slice(1)) * 1000;
      const rows = Array.from({ length: 101 }, (_, i) => record(base + 101 - i, account));
      return { total: 101, start: (page - 1) * 50 + 1, end: Math.min(page * 50, 101), records: rows.slice((page - 1) * 50, page * 50) };
    };
    const result = await collectScopedAccessReport('2026-09-30', accounts, read, new AbortController().signal, vi.fn(), 'TDR');
    expect(maximum).toBe(6); expect(active).toBe(0);
    expect(result.recordCount).toBe(1010); expect(result.users).toHaveLength(10);
  });
  it('drains in-flight requests on failure or cancellation without returning partial reports', async () => {
    let active = 0;
    const controller = new AbortController();
    const read = async (_page: number, account: string) => {
      active++;
      await new Promise((resolve) => setTimeout(resolve, 1)); active--;
      if (account === 'A') throw new Error('source failure');
      return { total: 0, start: 0, end: 0, records: [] };
    };
    await expect(collectScopedAccessReport('2026-09-30', ['A', 'B', 'C'], read, controller.signal, vi.fn(), 'TDR')).rejects.toThrow('source failure');
    expect(active).toBe(0);
    controller.abort();
    const spy = vi.fn(read);
    await expect(collectScopedAccessReport('2026-09-30', ['A'], spy, controller.signal, vi.fn(), 'TDR')).rejects.toThrow('hủy');
    expect(spy).not.toHaveBeenCalled();
  });
});
