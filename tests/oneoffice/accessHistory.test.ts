// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { collectAccessReport, cutoffEnd, isAccessHistoryUrl, readAccessHistoryPage, loadAccessPageFromTab, reportFromAccessSnapshot, type AccessRecord, type AccessPage } from '../../src/oneoffice/accessHistory';
import { buildAccessWorkbook } from '../../src/oneoffice/export/accessXlsx';

const record = (id: number, timeIso = '2026-09-30T16:59:59Z', account = 'A'): AccessRecord => ({
  id: String(id), account, name: '=sample', timeRaw: '23:59:59 30/09/2026', timeIso,
});
const data = (records: AccessRecord[], total = records.length, start = 1): AccessPage => ({ records, total, start, end: records.length ? start + records.length - 1 : 0 });
afterEach(() => { document.querySelectorAll('iframe').forEach((frame) => frame.remove()); vi.unstubAllGlobals(); vi.restoreAllMocks(); vi.useRealTimers(); });

describe('historical access report', () => {
  it('replenishes the rolling window before the slow last page in the old batch finishes, with at most six reads', async () => {
    let active = 0; let maximum = 0; let release: (() => void) | undefined;
    const calls: number[] = [];
    const read = async (page: number) => {
      calls.push(page); active++; maximum = Math.max(maximum, active);
      await Promise.resolve();
      if (page === 7) await new Promise<void>((resolve) => { release = resolve; });
      const rows = Array.from({length:50}, (_, i) => record(400 - (page - 1) * 50 - i));
      active--; return data(rows, 400, (page - 1) * 50 + 1);
    };
    const promise = collectAccessReport('2026-09-30', read, new AbortController().signal, vi.fn());
    await vi.waitFor(() => expect(calls).toContain(8));
    expect(maximum).toBe(6);
    release!();
    const report = await promise;
    expect(report.recordCount).toBe(400);
    expect(report.users[0].eventCount).toBe(400);
  });
  it('drains pending reads on cancellation without starting more pages', async () => {
    const controller = new AbortController();
    const calls: number[] = [];
    const read = async (page: number) => {
      calls.push(page);
      if (page === 2) controller.abort();
      return data(Array.from({length:50}, (_, i) => record(500 - (page - 1) * 50 - i)), 500, (page - 1) * 50 + 1);
    };
    await expect(collectAccessReport('2026-09-30', read, controller.signal, vi.fn())).rejects.toThrow('hủy');
    expect(calls).toEqual([1,2,3,4,5,6,7]);
  });
  it('handles newer IDs shifting page boundaries without counting overlapping rows twice', async () => {
    let firstCalls = 0;
    const read = vi.fn(async (page: number) => {
      if (page === 1 && ++firstCalls === 1) return data(Array.from({length:50}, (_, i) => record(51 - i)), 51);
      if (page === 1) return data(Array.from({length:50}, (_, i) => record(52 - i)), 52);
      return data([record(2), record(1)], 52, 51);
    });
    const report = await collectAccessReport('2026-09-30', read, new AbortController().signal, vi.fn());
    expect(report.recordCount).toBe(51);
    expect(report.users[0].eventCount).toBe(51);
    expect(report.snapshot?.records.some((item) => item.id === '52')).toBe(false);
  });
  it('recalculates dates from a complete snapshot without another network read', async () => {
    const read = vi.fn(async () => data([record(2, '2026-10-01T00:00:00Z'), record(1)]));
    const report = await collectAccessReport('2026-10-05', read, new AbortController().signal, vi.fn());
    const older = reportFromAccessSnapshot(report.snapshot!, '2026-09-30');
    expect(older.users[0]).toMatchObject({id:'1', eventCount:1});
    expect(older.collectedAt).toBe(report.collectedAt);
    expect(read).toHaveBeenCalledTimes(2);
  });
  it('validates dates and restricts the source to the known origin/path', () => {
    expect(cutoffEnd('2026-09-30')).toBe(Date.parse('2026-09-30T17:00:00Z'));
    expect(() => cutoffEnd('2026-02-30')).toThrow();
    expect(() => cutoffEnd('')).toThrow();
    expect(isAccessHistoryUrl('https://office.meygroup.vn/apps/admin-reconciliation-histories/list?p=2')).toBe(true);
    expect(isAccessHistoryUrl('https://evil.test/apps/admin-reconciliation-histories')).toBe(false);
    expect(isAccessHistoryUrl('https://office.meygroup.vn/apps/work')).toBe(false);
  });
  it('groups by account and excludes later events without conflating distinct IDs', async () => {
    const page = data([record(1), record(2), record(3, '2026-09-30T17:00:00Z'), record(4, '2026-09-01T00:00:00Z', 'B')]);
    const report = await collectAccessReport('2026-09-30', async () => page, new AbortController().signal, vi.fn());
    expect(report.users).toHaveLength(2);
    expect(report.users[0]).toMatchObject({ id: '2', eventCount: 2 });
    expect(report.recordCount).toBe(4);
    const workbook = buildAccessWorkbook(report);
    expect(workbook.map((sheet) => sheet.sheet)).toEqual(['Tổng hợp truy cập', 'Thông tin báo cáo']);
    expect(workbook[0].data[1][2]).toMatchObject({ value: '=sample', type: String });
  });
  it('reads all pages including the short final page', async () => {
    const read = vi.fn(async (page: number) => page === 1 ? data(Array.from({length:50}, (_, i) => record(i + 1)), 113)
      : data(Array.from({length:page === 2 ? 50 : 13}, (_, i) => record((page - 1) * 50 + i + 1)), 113, (page - 1) * 50 + 1));
    const report = await collectAccessReport('2026-09-30', read, new AbortController().signal, vi.fn());
    expect(report.recordCount).toBe(113);
    expect(read.mock.calls.map(([page]) => page)).toEqual([1, 2, 3, 1]);
  });
  it('blocks duplicates, drifting totals, wrong pages and changed final data', async () => {
    const first = data(Array.from({length:50}, (_, i) => record(i + 1)), 51);
    for (const second of [data([record(1)], 51, 51), data([record(51)], 52, 51), data([record(51)], 51, 1)]) {
      await expect(collectAccessReport('2026-09-30', async (page) => page === 1 ? first : second, new AbortController().signal, vi.fn())).rejects.toThrow();
    }
    let calls = 0;
    await expect(collectAccessReport('2026-09-30', async () => data([record(++calls)]), new AbortController().signal, vi.fn())).rejects.toThrow();
  });
  it('supports cancellation and empty retained history', async () => {
    const controller = new AbortController(); controller.abort();
    await expect(collectAccessReport('2026-09-30', async () => data([record(1)]), controller.signal, vi.fn())).rejects.toThrow('hủy');
    expect((await collectAccessReport('2026-09-30', async () => data([], 0, 0), new AbortController().signal, vi.fn())).users).toEqual([]);
  });
});

const html = (time = '23:59:59 30/09/2026') => `<table><tbody><tr><td></td><td>Tài khoản</td><td>Họ và tên</td><td>Thời gian</td><td>Ip</td><td>Trình duyệt</td></tr><tr><td><input name="ID[]" value="1"></td><td>A</td><td>User A</td><td>${time}</td><td>PRIVATE_IP</td><td>PRIVATE_UA</td></tr></tbody></table><div>Hiển thị 1 - 1 / 1 bản ghi</div>`;
describe('access history page reader', () => {
  const structured = (changes: Record<string, unknown> = {}) => {
    const props = { object: 'admin-reconciliation-histories',
      headers: ['username', 'fullname', 'date'].map((name) => ({ name })),
      rows: [{ ID: 1, vals: { username: 'A', fullname: "User O'Brien", date: '23:59:59 30/09/2026', ip: 'PRIVATE_IP', user_agent: 'PRIVATE_UA' } }],
      page: { total: 1, current: 1, range: 50, rq: 'p', params: { v: 'list' } },
      events: { created: ['throw new Error("DO_NOT_EXECUTE")'] }, ...changes };
    return `<elem name='table' data-props='${JSON.stringify(props).replace(/&/g, '&amp;').replace(/'/g, '&#39;')}'></elem>`;
  };
  it('sends the verified keyword field and refuses a source that does not acknowledge it', async () => {
    vi.stubGlobal('location', { origin: 'https://office.meygroup.vn', pathname: '/apps/admin-reconciliation-histories' });
    const fetchMock = vi.fn(async (_url: string) => ({ ok: true, text: async () => structured({ vars: { s: 'A' } }) }));
    vi.stubGlobal('fetch', fetchMock);
    expect(await readAccessHistoryPage(1, true, 'A')).toMatchObject({ ok: true });
    expect(new URL(fetchMock.mock.calls[0][0]).searchParams.get('s')).toBe('A');
    vi.stubGlobal('fetch', async () => ({ ok: true, text: async () => structured() }));
    expect(await readAccessHistoryPage(1, true, 'A')).toMatchObject({ ok: false, error: expect.stringContaining('Không đọc toàn bộ') });
    vi.stubGlobal('fetch', async () => ({ ok: true, text: async () => html() }));
    expect(await readAccessHistoryPage(1, true, 'A')).toMatchObject({ ok: false });
  });
  it('reads serialized table rows without loading frames or executing events', async () => {
    vi.stubGlobal('location', { origin: 'https://office.meygroup.vn', pathname: '/apps/admin-reconciliation-histories' });
    const fetchMock = vi.fn(async () => ({ ok: true, text: async () => structured() }));
    vi.stubGlobal('fetch', fetchMock);
    const append = vi.spyOn(document.body, 'append');
    const result = await readAccessHistoryPage(1, true);
    expect(result).toMatchObject({ ok: true, data: {total: 1, records: [{account:'A', name: "User O'Brien", timeIso:'2026-09-30T16:59:59.000Z'}]} });
    expect(JSON.stringify(result)).not.toContain('PRIVATE_');
    expect(append).not.toHaveBeenCalled();
    expect(new URL(fetchMock.mock.calls[0][0]).searchParams.get('_json')).toBe('1');
    expect(document.querySelector('iframe')).toBeNull();
  });
  it('rejects malformed serialized tables, missing rows and wrong pagination', async () => {
    vi.stubGlobal('location', { origin: 'https://office.meygroup.vn', pathname: '/apps/admin-reconciliation-histories' });
    for (const body of [structured({object:'other'}), structured({headers:[]}), structured({rows:[]}),
      structured({page:{total:1,current:2,range:50,rq:'p'}}), structured({page:{total:1,current:1,range:100,rq:'p'}}),
      structured({page:{total:-1,current:1,range:50,rq:'p'}}), structured({rows:[{ID:1,vals:{}}]}),
      "<elem name='table' data-props='invalid'></elem>", structured() + structured()]) {
      vi.stubGlobal('fetch', async () => ({ok:true,text:async()=>body}));
      expect(await readAccessHistoryPage(1,true)).toMatchObject({ok:false});
    }
    vi.stubGlobal('fetch', async () => ({ok:true,text:async()=>structured({rows:[],page:{total:0,current:1,range:50,rq:'p'}})}));
    expect(await readAccessHistoryPage(1)).toEqual({total:0,start:0,end:0,records:[]});
  });
  it('fails immediately for an unrecognized app shell rather than loading another page', async () => {
    vi.stubGlobal('location', { origin: 'https://office.meygroup.vn', pathname: '/apps/admin-reconciliation-histories' });
    vi.stubGlobal('fetch', async () => ({ ok: true, text: async () => '<div id="app"></div>' }));
    expect(await readAccessHistoryPage(2, true)).toMatchObject({ ok: false, error: expect.stringContaining('Không mở trang ngầm') });
    expect(document.querySelector('iframe')).toBeNull();
  });
  it('returns a serializable error instead of losing an injected exception', async () => {
    vi.stubGlobal('location', { origin: 'https://office.meygroup.vn', pathname: '/apps/admin-reconciliation-histories' });
    vi.stubGlobal('fetch', async () => ({ ok: false, status: 403 }));
    expect(await readAccessHistoryPage(1, true)).toEqual({ ok: false, error: 'Không tải được trang 1 (HTTP 403).' });
    const executeScript = vi.fn(async ({func, args}) => [{result: await func(...args)}]);
    vi.stubGlobal('chrome', { scripting: { executeScript } });
    await expect(loadAccessPageFromTab(42, 1)).rejects.toThrow('HTTP 403');
    expect(executeScript).toHaveBeenCalledWith(expect.objectContaining({target:{tabId:42}, args:[1,true,null,'https://office.meygroup.vn']}));
  });
  it('transfers successful pages and distinguishes JSON and transport failures without exposing payloads', async () => {
    vi.stubGlobal('location', { origin: 'https://office.meygroup.vn', pathname: '/apps/admin-reconciliation-histories' });
    vi.stubGlobal('chrome', { scripting: { executeScript: async ({func, args}) => [{result: await func(...args)}] } });
    vi.stubGlobal('fetch', async () => ({ ok: true, text: async () => html() }));
    expect(await loadAccessPageFromTab(42, 1)).toMatchObject({total:1});
    vi.stubGlobal('fetch', async () => ({ ok: true, text: async () => '{"private":"PRIVATE_DATA"}' }));
    await expect(loadAccessPageFromTab(42, 1)).rejects.toThrow('dạng JSON');
    expect(JSON.stringify(await readAccessHistoryPage(1,true))).not.toContain('PRIVATE_DATA');
    vi.stubGlobal('fetch', async () => { throw new TypeError('Failed to fetch'); });
    await expect(loadAccessPageFromTab(42, 1)).rejects.toThrow('lỗi kết nối hoặc chuyển hướng');
  });
  it('parses the observed td-header layout and omits IP/browser data', async () => {
    vi.stubGlobal('location', { origin: 'https://office.meygroup.vn', pathname: '/apps/admin-reconciliation-histories' });
    const fetchMock = vi.fn(async () => ({ ok: true, text: async () => html() })); vi.stubGlobal('fetch', fetchMock);
    expect(await readAccessHistoryPage(1)).toMatchObject({ total: 1, records: [{ account: 'A', timeIso: '2026-09-30T16:59:59.000Z' }] });
    expect(JSON.stringify(await readAccessHistoryPage(1))).not.toContain('PRIVATE_');
    expect(fetchMock.mock.calls[0][1]).toMatchObject({ credentials: 'include', redirect: 'error' });
  });
  it('rejects login responses, invalid times and repeated page ranges', async () => {
    vi.stubGlobal('location', { origin: 'https://office.meygroup.vn', pathname: '/apps/admin-reconciliation-histories' });
    for (const body of ['<form><input type="password"></form>', html('25:00:00 30/09/2026'), html('12:00:00 31/02/2026')]) {
      vi.stubGlobal('fetch', async () => ({ ok: true, text: async () => body }));
      await expect(readAccessHistoryPage(1)).rejects.toThrow();
    }
    vi.stubGlobal('fetch', async () => ({ ok: true, text: async () => html() }));
    await expect(readAccessHistoryPage(2)).rejects.toThrow('phạm vi');
  });
  it('rejects foreign source tabs and HTTP errors', async () => {
    vi.stubGlobal('location', { origin: 'https://evil.test', pathname: '/apps/admin-reconciliation-histories' });
    await expect(readAccessHistoryPage(1)).rejects.toThrow('Tab nguồn');
    vi.stubGlobal('location', { origin: 'https://office.meygroup.vn', pathname: '/apps/admin-reconciliation-histories' });
    vi.stubGlobal('fetch', async () => ({ ok: false, status: 403 }));
    await expect(readAccessHistoryPage(1)).rejects.toThrow('403');
  });
});
