// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { loadAccessPersonnel, readAccessPersonnelSource } from '../../src/oneoffice/accessPersonnel';
import { reportFromAccessSnapshot } from '../../src/oneoffice/accessHistory';
import { buildAccessWorkbook } from '../../src/oneoffice/export/accessXlsx';

afterEach(() => { vi.unstubAllGlobals(); vi.restoreAllMocks(); });
const directory = (rows: unknown[], total = rows.length, current = 1) => {
  const props = { object: 'admin-user', rows, page: { total, current, range: 50, rq: 'p' },
    events: { created: ['throw new Error("DO_NOT_EXECUTE")'] } };
  return `<elem name="table" data-props='${JSON.stringify(props).replace(/&/g, '&amp;').replace(/'/g, '&#39;')}'></elem>`;
};
const setup = (fetcher: (url: string) => Promise<unknown>) => {
  vi.stubGlobal('location', { origin: 'https://office.meygroup.vn', pathname: '/apps/admin-reconciliation-histories' });
  vi.stubGlobal('fetch', async (value: string) => {
    const response = await fetcher(value) as any;
    if (!response.ok || !new URL(value).pathname.includes('admin-user')) return response;
    const body = await response.text();
    const doc = new DOMParser().parseFromString(body, 'text/html');
    const table = doc.querySelector('elem[name="table"][data-props]');
    if (!table) return { ...response, text: async () => body };
    const props = JSON.parse(table.getAttribute('data-props')!);
    props.page.params = { tab: new URL(value).searchParams.get('tab') };
    table.setAttribute('data-props', JSON.stringify(props));
    return { ...response, text: async () => doc.body.innerHTML };
  });
  vi.stubGlobal('chrome', { scripting: { executeScript: async ({ func, args }: any) => [{ result: await func(...args) }] } });
};
const response = (body: string) => ({ ok: true, text: async () => body });

describe('access personnel enrichment', () => {
  it('inherits TDK/TDU grouping through uncoded wrappers without changing actual company codes', async () => {
    const row = (ID: number, code: string, rows: unknown[] = []) => ({ ID, vals: { title: 'Same source name', type: 'BRANCH' }, fields: { title: code ? `[${code}] Same source name` : 'Same source name' }, rows });
    const rows = [row(1, 'TAPDOAN', [row(2, '', [row(3, 'TDK'), row(4, 'TDH')]), row(5, '', [row(6, 'TDU'), row(7, 'TDU.CKKG')]), row(8, 'TDG')])];
    const props = { object: 'admin-department', rows, page: { current: 1, total: 8, range: 8 } };
    setup(async () => response(`<elem name="table" data-props='${JSON.stringify(props)}'></elem>`));
    const result = await readAccessPersonnelSource({ kind: 'organization' });
    expect(result).toMatchObject({ ok: true, data: [
      { id: '1', companyCode: 'TAPDOAN' }, { id: '2', companyCode: '', companyGroup: 'TDK' },
      { id: '3', companyCode: 'TDK', companyGroup: 'TDK' }, { id: '4', companyCode: 'TDH', companyGroup: 'TDK' },
      { id: '5', companyCode: '', companyGroup: 'TDU' }, { id: '6', companyCode: 'TDU', companyGroup: 'TDU' },
      { id: '7', companyCode: 'TDU.CKKG', companyGroup: 'TDU' }, { id: '8', companyCode: 'TDG' },
    ] });
    if (result.ok) expect((result.data as any[])[7].companyGroup).toBeUndefined();
  });
  it('refuses a response that ignores the requested active/blocked tab', async () => {
    setup(async () => response(''));
    vi.stubGlobal('fetch', async () => response(directory([{ ID: 1, vals: { username: 'LOCKED' } }])));
    expect(await readAccessPersonnelSource({ kind: 'directory', tab: 'active', page: 1 })).toMatchObject({ ok: false });
    const props = { object: 'admin-user', rows: [{ ID: 1, vals: { username: 'LOCKED' } }], page: { total: 1, current: 1, range: 50, rq: 'p', params: { tab: 'block' } } };
    vi.stubGlobal('fetch', async () => response(`<elem name="table" data-props='${JSON.stringify(props)}'></elem>`));
    expect(await readAccessPersonnelSource({ kind: 'directory', tab: 'active', page: 1 })).toMatchObject({ ok: false });
    expect(await readAccessPersonnelSource({ kind: 'directory', tab: 'block', page: 1 })).toMatchObject({ ok: true });
  });
  it('removes real and entity-encoded icon HTML before splitting department/position and exporting', async () => {
    setup(async () => response(JSON.stringify({ items: [{ ID: 1,
      placeholder: '<span class="icon-workflow"></span> R - Hạ tầng || &lt;span class="icon-conference-call"&gt;&lt;/span&gt; TDR - Giám đốc<script>PRIVATE_SCRIPT</script>' }] })));
    const result = await readAccessPersonnelSource({ kind: 'lookup', ids: ['1'] });
    expect(result).toEqual({ ok: true, data: [{ id: '1', department: 'R - Hạ tầng', position: 'TDR - Giám đốc' }] });
    const report = reportFromAccessSnapshot({ collectedAt: '', recordCount: 1, oldestTime: '',
      records: [{ id: '1', account: 'A', name: 'A', timeIso: '2026-09-29T00:00:00Z', timeRaw: '' }],
      personnel: { collectedAt: '', warnings: [], users: [{ account: 'A', position: '<span></span> Giám đốc', department: '&lt;b&gt;Hạ tầng&lt;/b&gt;', company: '<b>Meyland</b>' }] } }, '2026-09-30');
    expect(buildAccessWorkbook(report)[0].data[1].slice(6).map((cell: any) => cell.value)).toEqual(['Giám đốc', 'Hạ tầng', 'Meyland']);
  });
  it('loads the complete organization tree and joins companies by numeric department IDs', async () => {
    const rows = [{ ID: 394, vals: { title: 'Tập đoàn', type: 'COMPANY' }, fields: { title: '[TAPDOAN] Tập đoàn' }, rows: [
      { ID: 397, vals: { title: 'Meyland', type: 'BRANCH' }, fields: { title: '<a>2. [TDR] Meyland</a>' }, rows: [
        { ID: 400, vals: { title: 'R - Hội đồng quản trị', type: 'DEPART' }, fields: { title: '[TDR.HDQT] R - Hội đồng quản trị' } },
      ] },
    ] }];
    const organization = `<elem name="table" data-props='${JSON.stringify({ object: 'admin-department', rows, page: { current: 1, total: 3, range: 3 } })}'></elem>`;
    setup(async (value) => {
      const url = new URL(value);
      if (url.pathname.includes('admin-department')) return response(organization);
      if (url.pathname.includes('admin-user')) return response(directory(url.searchParams.get('tab') === 'active' ?
        [{ ID: 1, vals: { username: 'A', department_id: 400 } }] : []));
      return response(JSON.stringify({ items: [{ ID: 1, placeholder: '<span></span> R - Hội đồng quản trị || TDR - Giám đốc' }] }));
    });
    const result = await loadAccessPersonnel(42, null, new AbortController().signal, vi.fn());
    expect(result.users).toEqual([{ account: 'A', position: 'TDR - Giám đốc', department: 'R - Hội đồng quản trị', company: 'Meyland', companyCode: 'TDR', departmentCode: 'TDR.HDQT' }]);
    expect(result.warnings).toEqual([]);
    setup(async () => response(organization.replace('"total":3', '"total":4')));
    expect(await readAccessPersonnelSource({ kind: 'organization' })).toMatchObject({ ok: false });
  });
  it('reads only stable account IDs, without requiring visible position/department columns', async () => {
    setup(async () => response(directory([{ ID: 1, vals: { username: 'A', email: 'PRIVATE_EMAIL', department_id: 'ignore' } }])));
    const result = await readAccessPersonnelSource({ kind: 'directory', tab: 'active', page: 1 });
    expect(result).toEqual({ ok: true, data: { total: 1, users: [{ id: '1', account: 'A', department: '' }] } });
    expect(JSON.stringify(result)).not.toContain('PRIVATE_');
  });
  it('uses the verified department/position separator and does not confuse missing position with department', async () => {
    setup(async () => response(JSON.stringify({ items: [
      { ID: 1, placeholder: ' Dept A || Position A', gender: 'PRIVATE' },
      { ID: 2, placeholder: 'Dept B' }, { ID: 3, placeholder: '' },
    ] })));
    expect(await readAccessPersonnelSource({ kind: 'lookup', ids: ['1', '2', '3'] })).toEqual({ ok: true, data: [
      { id: '1', department: 'Dept A', position: 'Position A' },
      { id: '2', department: 'Dept B', position: '' }, { id: '3', department: '', position: '' },
    ] });
  });
  it('rejects unknown, duplicate IDs and changed placeholder formats', async () => {
    for (const items of [ [{ ID: 2, placeholder: 'Dept' }], [{ ID: 1, placeholder: 'Dept || Pos || Other' }],
      [{ ID: 1, placeholder: 'Dept' }, { ID: 1, placeholder: 'Dept' }], [{ ID: 1, placeholder: 123 }]]) {
      setup(async () => response(JSON.stringify({ items })));
      expect(await readAccessPersonnelSource({ kind: 'lookup', ids: ['1'] })).toMatchObject({ ok: false });
    }
  });
  it('accepts missing/null placeholders without losing other users in the same batch', async () => {
    setup(async () => response(JSON.stringify({ items: [
      { ID: 1, name: 'No personnel' }, { ID: 2, placeholder: null }, { ID: 3, placeholder: 'Dept || Position' },
    ] })));
    expect(await readAccessPersonnelSource({ kind: 'lookup', ids: ['1', '2', '3'] })).toEqual({ ok: true, data: [
      { id: '1', department: '', position: '' }, { id: '2', department: '', position: '' },
      { id: '3', department: 'Dept', position: 'Position' },
    ] });
  });
  it('excludes locked accounts even if old histories or lookup metadata exist', async () => {
    setup(async (value) => {
      const url = new URL(value);
      if (url.pathname.endsWith('/list')) return response(directory(url.searchParams.get('tab') === 'active' ?
        [{ ID: 1, vals: { username: 'A' } }] :
        [{ ID: 2, vals: { username: 'LOCKED', department_id: 999 }, fields: { department_id: '<a>Locked department</a>' } }]));
      return response(JSON.stringify({ items: [{ ID: 1, placeholder: 'Dept || Position' }] }));
    });
    const result = await loadAccessPersonnel(42, ['A', 'LOCKED'], new AbortController().signal, vi.fn());
    expect(result.users).toEqual([{ account: 'A', position: 'Position', department: 'Dept' }]);
    expect(result.activeAccounts).toEqual(['A']);
    expect(result.warnings.join(' ')).toContain('đã khóa');
  });
  it('keeps successful batches when another batch fails and records the failure', async () => {
    setup(async (value) => {
      const url = new URL(value);
      if (url.pathname.endsWith('/list')) {
        if (url.searchParams.get('tab') === 'block') return response(directory([]));
        const page = Number(url.searchParams.get('p'));
        return response(directory(Array.from({ length: page === 3 ? 1 : 50 }, (_, i) => {
          const id = (page - 1) * 50 + i + 1;
          return { ID: id, vals: { username: `A${id}`, department_id: 999 }, fields: { department_id: 'Directory dept' } };
        }), 101, page));
      }
      const ids = url.searchParams.getAll('ids[]');
      if (ids.includes('101')) return { ok: false, status: 503 };
      return response(JSON.stringify({ items: ids.map((ID) => ({ ID, placeholder: 'Dept || Position' })) }));
    });
    const result = await loadAccessPersonnel(42, Array.from({ length: 101 }, (_, i) => `A${i + 1}`), new AbortController().signal, vi.fn());
    expect(result.users.filter((user) => user.position === 'Position')).toHaveLength(100);
    expect(result.users[100]).toEqual({ account: 'A101', position: '', department: 'Directory dept' });
    expect(result.warnings.join(' ')).toContain('HTTP 503');
  });
  it('rejects foreign origins, malformed directories and expired sessions without exposing response bodies', async () => {
    setup(async () => response('PRIVATE_MALFORMED'));
    expect(await readAccessPersonnelSource({ kind: 'lookup', ids: ['1'] })).toMatchObject({ ok: false });
    vi.stubGlobal('location', { origin: 'https://evil.test', pathname: '/apps/admin-reconciliation-histories' });
    expect(await readAccessPersonnelSource({ kind: 'directory', tab: 'active', page: 1 })).toMatchObject({ ok: false });
    setup(async () => ({ ok: false, status: 403 }));
    expect(await readAccessPersonnelSource({ kind: 'directory', tab: 'active', page: 1 })).toMatchObject({ ok: false, error: expect.stringContaining('403') });
    setup(async () => response(directory([{ ID: 1, vals: { username: 'A' } }], 51)));
    expect(await readAccessPersonnelSource({ kind: 'directory', tab: 'active', page: 1 })).toMatchObject({ ok: false });
  });
  it('includes locked users, refuses ambiguous accounts and never joins by name', async () => {
    setup(async (value) => {
      const url = new URL(value);
      if (url.pathname.endsWith('/list')) return response(directory(url.searchParams.get('tab') === 'active' ?
        [{ ID: 1, vals: { username: 'A' } }, { ID: 2, vals: { username: 'DUP' } }] :
        [{ ID: 3, vals: { username: 'LOCKED' } }, { ID: 4, vals: { username: 'DUP' } }]));
      expect(url.searchParams.getAll('ids[]')).toEqual(['1']);
      return response(JSON.stringify({ items: [{ ID: 1, placeholder: 'Dept || Pos' }] }));
    });
    const result = await loadAccessPersonnel(42, ['A', 'LOCKED', 'DUP', 'UNKNOWN'], new AbortController().signal, vi.fn());
    expect(result.users).toEqual([{ account: 'A', position: 'Pos', department: 'Dept' }]);
    expect(result.activeAccounts).toEqual(['A']);
    expect(result.warnings.join(' ')).toContain('3 tài khoản');
  });
  it('reads all directory pages and groups lookups into at most 100 IDs', async () => {
    const batches: number[] = [];
    setup(async (value) => {
      const url = new URL(value);
      if (url.pathname.endsWith('/list')) {
        if (url.searchParams.get('tab') === 'block') return response(directory([]));
        const page = Number(url.searchParams.get('p'));
        return response(directory(Array.from({ length: page === 3 ? 13 : 50 }, (_, i) => {
          const id = (page - 1) * 50 + i + 1; return { ID: id, vals: { username: `A${id}` } };
        }), 113, page));
      }
      const ids = url.searchParams.getAll('ids[]'); batches.push(ids.length);
      return response(JSON.stringify({ items: ids.map((ID) => ({ ID, placeholder: 'Dept || Pos' })) }));
    });
    const result = await loadAccessPersonnel(42, Array.from({ length: 113 }, (_, i) => `A${i + 1}`), new AbortController().signal, vi.fn());
    expect(result.users).toHaveLength(113); expect(batches).toEqual([100, 13]);
  });
  it('blocks directory drift, duplicate IDs and cancellation', async () => {
    let calls = 0;
    setup(async () => response(directory([{ ID: ++calls, vals: { username: 'A' } }])));
    await expect(loadAccessPersonnel(42, ['A'], new AbortController().signal, vi.fn())).rejects.toThrow('thay đổi');
    setup(async () => response(directory([{ ID: 1, vals: { username: 'A' } }])));
    await expect(loadAccessPersonnel(42, ['A'], new AbortController().signal, vi.fn())).rejects.toThrow('ID trùng');
    const controller = new AbortController(); controller.abort();
    await expect(loadAccessPersonnel(42, ['A'], controller.signal, vi.fn())).rejects.toThrow('hủy');
  });
  it('preserves enrichment on date changes, keeps original columns and writes new fields as literal text', () => {
    const personnel = { collectedAt: '2026-10-06T00:00:00Z', users: [{ account: 'A', position: '=not a formula', department: 'Dept' }], warnings: ['Missing account'] };
    const report = reportFromAccessSnapshot({ collectedAt: personnel.collectedAt, recordCount: 1, oldestTime: '', personnel,
      records: [{ id: '1', account: 'A', name: 'Historical name', timeIso: '2026-09-30T00:00:00Z', timeRaw: '07:00:00 30/09/2026' }] }, '2026-09-30');
    expect(report.personnel).toBe(personnel);
    const workbook = buildAccessWorkbook(report);
    expect(workbook[0].data[0].map((cell: any) => cell.value)).toEqual(['STT', 'Tài khoản', 'Họ và tên tại bản ghi', 'Truy cập cuối trước ngày chốt', 'Ngày chốt', 'Số bản ghi đến ngày chốt', 'Vị trí hiện tại', 'Phòng ban hiện tại', 'Công ty hiện tại']);
    expect(workbook[0].data[1][6]).toMatchObject({ value: '=not a formula', type: String, format: '@' });
    expect(workbook[1].data.flat().some((cell: any) => cell.value === 'Missing account')).toBe(true);
  });
});
