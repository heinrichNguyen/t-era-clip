import type { AccessPersonnel } from './accessHistory';
import { oneOfficeOrigin } from './origins';

type Organization = Array<{ id: string; company: string; companyCode: string; companyGroup?: string; departmentCode: string }>;
type DirectoryPage = { total: number; users: Array<{ id: string; account: string; department: string; departmentId?: string }> };
type LookupUsers = Array<{ id: string; position: string; department: string }>;
type Request = { kind: 'directory'; tab: 'active' | 'block'; page: number } | { kind: 'lookup'; ids: string[] } | { kind: 'organization' };
type Result = { ok: true; data: DirectoryPage | LookupUsers | Organization } | { ok: false; error: string };

// Injected into the authenticated source tab. Parse data only, never site scripts.
export async function readAccessPersonnelSource(request: Request, expectedOrigin = 'https://office.meygroup.vn'): Promise<Result> {
  // Must stay inside the injected function, which cannot reference module helpers.
  const plain = (value: string) => {
    let result = value;
    for (let pass = 0; pass < 4; pass++) {
      const doc = new DOMParser().parseFromString(result, 'text/html');
      doc.querySelectorAll('script, style, iframe, object').forEach((node) => node.remove());
      const next = (doc.body.textContent ?? '').replace(/\s+/g, ' ').trim();
      if (next === result) break;
      result = next;
    }
    return result === '--' ? '' : result;
  };
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 30000);
  try {
    if (!['https://office.meygroup.vn', 'https://testoffice.meygroup.vn', 'https://testoffice.meygroup.vn:9443'].includes(expectedOrigin) || location.origin !== expectedOrigin ||
      !/^\/apps\/admin-reconciliation-histories(?:\/list)?\/?$/.test(location.pathname)) {
      throw new Error('Tab nguồn không còn ở Lịch sử truy cập 1Office.');
    }
    const url = new URL(request.kind === 'organization' ? '/apps/admin-department/list' :
      request.kind === 'directory' ? '/apps/admin-user/list' : '/app/autofill/users/items', location.origin);
    if (request.kind === 'organization') {
      url.search = new URLSearchParams({ v: 'list', tab: 'department', _json: '1', inlineLogin: '1' }).toString();
    } else if (request.kind === 'directory') {
      if (!['active', 'block'].includes(request.tab) || !Number.isInteger(request.page) || request.page < 1 || request.page > 100) {
        throw new Error('Trang danh sách tài khoản không hợp lệ.');
      }
      url.search = new URLSearchParams({ v: 'list', tab: request.tab, p: String(request.page), _json: '1', inlineLogin: '1' }).toString();
    } else {
      if (!Array.isArray(request.ids) || !request.ids.length || request.ids.length > 100 ||
        request.ids.some((id) => !/^\d+$/.test(id) || !Number.isSafeInteger(Number(id)) || Number(id) <= 0)) {
        throw new Error('Lô tài khoản không hợp lệ.');
      }
      request.ids.forEach((id) => url.searchParams.append('ids[]', id));
    }
    const response = await fetch(url.href, { credentials: 'include', redirect: 'error', signal: controller.signal });
    if (!response.ok) throw new Error(`Không đọc được thông tin nhân sự (HTTP ${response.status}).`);
    const body = await response.text();
    if (request.kind === 'lookup') {
      const parsed = JSON.parse(body);
      if (!Array.isArray(parsed.items)) throw new Error('Nguồn thông tin nhân sự không đúng định dạng.');
      const seen = new Set<string>();
      const users: LookupUsers = parsed.items.map((item: { ID?: unknown; placeholder?: unknown }) => {
        const id = String(item?.ID ?? '');
        if (!request.ids.includes(id) || seen.has(id)) {
          throw new Error('Không đối chiếu được ID thông tin nhân sự.');
        }
        seen.add(id);
        // Some real accounts have no placeholder at all, not a malformed ID.
        if (item.placeholder == null) return { id, department: '', position: '' };
        if (typeof item.placeholder !== 'string') throw new Error('Cấu trúc vị trí/phòng ban không hợp lệ.');
        const parts = plain(item.placeholder).split('||').map(plain);
        if (parts.length > 2) throw new Error('Cấu trúc vị trí/phòng ban đã thay đổi.');
        return { id, department: parts[0] || '', position: parts[1] || '' };
      });
      return { ok: true, data: users };
    }
    const doc = new DOMParser().parseFromString(body, 'text/html');
    const tables = doc.querySelectorAll('elem[name="table"][data-props]');
    if (tables.length !== 1) throw new Error('Không nhận diện được danh sách tài khoản. Kiểm tra quyền đọc người dùng.');
    const props = JSON.parse(tables[0].getAttribute('data-props') ?? '');
    const paging = props?.page;
    if (request.kind === 'organization') {
      if (props.object !== 'admin-department' || !Array.isArray(props.rows) || !Number.isSafeInteger(paging?.total) || paging.total < 1 || paging.total > 5000 ||
        paging.current !== 1 || paging.range < paging.total) throw new Error('Cây cơ cấu nhân sự chưa đầy đủ.');
      const units: Organization = [];
      const seen = new Set<string>();
      const visit = (row: any, company = '', companyCode = '', depth = 0, companyGroup = '') => {
        const id = String(row?.ID ?? '');
        if (!/^\d+$/.test(id) || !Number.isSafeInteger(Number(id)) || Number(id) <= 0 || seen.has(id) || depth > 64 || seen.size >= 5000 || typeof row.vals?.title !== 'string') {
          throw new Error('Cây cơ cấu nhân sự không hợp lệ.');
        }
        seen.add(id);
        const label = typeof row.fields?.title === 'string' ? plain(row.fields.title) : '';
        const code = /\[([^\[\]]+)\]/.exec(label)?.[1]?.trim() ?? '';
        if (row.rows !== undefined && !Array.isArray(row.rows)) throw new Error('Nhánh cơ cấu không hợp lệ.');
        // The source has un-coded wrappers above TDK/TDU. Grouping follows their verified child codes,
        // without inventing a source company code or joining by duplicate company names.
        if (['COMPANY', 'BRANCH'].includes(row.vals.type) && !code) {
          const childCodes = (row.rows ?? []).map((child: any) => typeof child.fields?.title === 'string' ?
            /\[([^\[\]]+)\]/.exec(plain(child.fields.title))?.[1]?.trim() : '');
          if (childCodes.includes('TDK') && !childCodes.includes('TDU')) companyGroup = 'TDK';
          if (childCodes.includes('TDU') && !childCodes.includes('TDK')) companyGroup = 'TDU';
        }
        if (['TDK', 'TDU'].includes(code)) companyGroup = code;
        if (['COMPANY', 'BRANCH'].includes(row.vals.type)) { company = plain(row.vals.title); companyCode = code; }
        units.push({ id, company, companyCode, departmentCode: code, ...(companyGroup ? { companyGroup } : {}) });
        (row.rows ?? []).forEach((child: unknown) => visit(child, company, companyCode, depth + 1, companyGroup));
      };
      props.rows.forEach((row: unknown) => visit(row));
      if (units.length !== paging.total) throw new Error('Cây cơ cấu nhân sự thiếu đơn vị.');
      return { ok: true, data: units };
    }
    if (props.object !== 'admin-user' || !Array.isArray(props.rows) || paging?.current !== request.page ||
      paging.range !== 50 || paging.rq !== 'p' || !Number.isSafeInteger(paging.total) || paging.total < 0 || paging.total > 5000) {
      throw new Error('Phân trang danh sách tài khoản không hợp lệ.');
    }
    if (paging.params?.tab !== request.tab && props.vars?.tab !== request.tab) {
      throw new Error('Không xác minh được danh sách tài khoản hoạt động/đã khóa.');
    }
    const expected = Math.max(0, Math.min(50, paging.total - (request.page - 1) * 50));
    if (props.rows.length !== expected || (request.page > 1 && expected === 0)) throw new Error('Danh sách tài khoản chưa đủ dữ liệu.');
    const users = props.rows.map((row: { ID?: unknown; vals?: { username?: unknown; department_id?: unknown }; fields?: { department_id?: unknown } }) => {
      const id = String(row?.ID ?? '');
      if (!/^\d+$/.test(id) || !Number.isSafeInteger(Number(id)) || Number(id) <= 0 ||
        typeof row.vals?.username !== 'string' || !row.vals.username.trim()) {
        throw new Error('Tài khoản thiếu ID hoặc tên đăng nhập để ghép dữ liệu.');
      }
      // vals.department_id is a numeric ID; fields contains its displayed name.
      let department = '';
      if (typeof row.fields?.department_id === 'string') {
        department = plain(row.fields.department_id);
      }
      const departmentId = String(row.vals.department_id ?? '');
      return { id, account: row.vals.username.trim(), department,
        ...(/^\d+$/.test(departmentId) && Number(departmentId) > 0 ? { departmentId } : {}) };
    });
    if (new Set(users.map((user: { id: string }) => user.id)).size !== users.length) throw new Error('ID tài khoản bị trùng.');
    return { ok: true, data: { total: paging.total, users } };
  } catch (error) {
    return { ok: false, error: controller.signal.aborted ? '1Office không phản hồi thông tin nhân sự trong 30 giây. Hãy thử lại.' :
      error instanceof Error && !['SyntaxError', 'TypeError'].includes(error.name) ? error.message : 'Không đọc được thông tin nhân sự. Kiểm tra phiên đăng nhập và định dạng nguồn.' };
  } finally { clearTimeout(timeout); }
}

export async function loadAccessPersonnel(
  tabId: number, accounts: string[] | null, signal: AbortSignal,
  progress: (loaded: number, total: number) => void,
  expectedOrigin = 'https://office.meygroup.vn',
): Promise<AccessPersonnel> {
  if (signal.aborted) throw new Error('Đã hủy bổ sung thông tin nhân sự.');
  const read = async (request: Request) => {
    if (signal.aborted) throw new Error('Đã hủy bổ sung thông tin nhân sự.');
    const api = (globalThis as any).chrome;
    if (!api?.scripting?.executeScript) throw new Error('Không có quyền đọc tab nguồn.');
    const results = await api.scripting.executeScript({ target: { tabId }, func: readAccessPersonnelSource, args: [request, oneOfficeOrigin(expectedOrigin)] });
    if (signal.aborted) throw new Error('Đã hủy bổ sung thông tin nhân sự.');
    const result = results[0]?.result as Result | undefined;
    if (!result?.ok) throw new Error(result?.error || 'Tab nguồn không trả về thông tin nhân sự.');
    return result.data;
  };
  const directory: DirectoryPage['users'] = [];
  const active = new Set<string>();
  for (const tab of ['active', 'block'] as const) {
    const first = await read({ kind: 'directory', tab, page: 1 }) as DirectoryPage;
    directory.push(...first.users);
    if (tab === 'active') first.users.forEach((user) => active.add(user.account));
    for (let page = 2; page <= Math.ceil(first.total / 50); page += 4) {
      const batch = Array.from({ length: Math.min(4, Math.ceil(first.total / 50) - page + 1) }, (_, i) => page + i);
      const results = await Promise.allSettled(batch.map((p) => read({ kind: 'directory', tab, page: p })));
      for (const result of results) {
        if (result.status === 'rejected') throw result.reason;
        const data = result.value as DirectoryPage;
        if (data.total !== first.total) throw new Error('Danh sách người dùng thay đổi khi đối chiếu. Hãy thu thập lại.');
        directory.push(...data.users);
        if (tab === 'active') data.users.forEach((user) => active.add(user.account));
      }
    }
    const final = await read({ kind: 'directory', tab, page: 1 }) as DirectoryPage;
    if (final.total !== first.total || JSON.stringify(final.users) !== JSON.stringify(first.users)) {
      throw new Error('Danh sách người dùng thay đổi khi đối chiếu. Hãy thu thập lại.');
    }
  }
  const byAccount = new Map<string, string>();
  const departmentById = new Map<string, string>();
  const departmentIds = new Map<string, string>();
  const ambiguous = new Set<string>();
  const seenIds = new Set<string>();
  for (const user of directory) {
    if (seenIds.has(user.id)) throw new Error('Danh sách tài khoản có ID trùng giữa các trang. Không ghép thông tin nhân sự.');
    seenIds.add(user.id);
    if (byAccount.has(user.account)) ambiguous.add(user.account);
    byAccount.set(user.account, user.id);
    departmentById.set(user.id, user.department);
    if (user.departmentId) departmentIds.set(user.id, user.departmentId);
  }
  const requested = accounts ?? [...byAccount.keys()];
  const activeAccounts = [...active].filter((account) => !ambiguous.has(account));
  const selected = [...new Set(requested)].filter((account) => active.has(account) && byAccount.has(account) && !ambiguous.has(account));
  const ids = selected.map((account) => byAccount.get(account)!);
  const details = new Map<string, LookupUsers[number]>();
  const warnings: string[] = [];
  let organization: Organization = [];
  try { organization = await read({ kind: 'organization' }) as Organization; }
  catch (error) {
    if (signal.aborted) throw error;
    warnings.push(`Chưa đối chiếu được công ty theo cơ cấu: ${error instanceof Error ? error.message : 'Nguồn không phản hồi'}`);
  }
  const units = new Map(organization.map((unit) => [unit.id, unit]));
  progress(0, ids.length);
  for (let index = 0; index < ids.length; index += 400) {
    const batches: string[][] = [];
    for (let offset = index; offset < Math.min(ids.length, index + 400); offset += 100) batches.push(ids.slice(offset, offset + 100));
    const results = await Promise.allSettled(batches.map((batch) => read({ kind: 'lookup', ids: batch })));
    if (signal.aborted) throw new Error('Đã hủy bổ sung thông tin nhân sự.');
    for (let batch = 0; batch < results.length; batch++) {
      const result = results[batch];
      if (result.status === 'rejected') {
        warnings.push(`Không tra được vị trí cho lô ${batches[batch].length} tài khoản: ${result.reason instanceof Error ? result.reason.message : 'Nguồn không phản hồi'}. Giữ phòng ban đã đọc từ danh sách người dùng.`);
        continue;
      }
      for (const user of result.value as LookupUsers) details.set(user.id, user);
    }
    progress(details.size, ids.length);
  }
  const users = selected.flatMap((account) => {
    const id = byAccount.get(account)!;
    const detail = details.get(id);
    const unit = units.get(departmentIds.get(id) ?? '');
    return [{ account, position: detail?.position ?? '', department: detail?.department || departmentById.get(id) || '',
      ...(unit ? { company: unit.company, companyCode: unit.companyCode, departmentCode: unit.departmentCode,
        ...(unit.companyGroup ? { companyGroup: unit.companyGroup } : {}) } : {}) }];
  });
  const missing = new Set(requested).size - users.length;
  if (missing) warnings.push(`${missing} tài khoản bị loại vì đã khóa, không có trong danh sách active hoặc tên tài khoản trùng. Không ghép theo họ tên.`);
  const noPosition = users.filter((user) => !user.position).length;
  const noDepartment = users.filter((user) => !user.department).length;
  if (noPosition || noDepartment) warnings.push(`Nguồn nhân sự chưa có vị trí cho ${noPosition} tài khoản, chưa có phòng ban cho ${noDepartment} tài khoản đã ghép.`);
  const noCompany = users.filter((user) => !user.companyCode && !user.companyGroup).length;
  if (noCompany) warnings.push(`${noCompany} tài khoản chưa xác minh được công ty theo ID phòng ban; không tự đưa vào bộ lọc công ty.`);
  return { collectedAt: new Date().toISOString(), users, warnings, activeAccounts };
}
