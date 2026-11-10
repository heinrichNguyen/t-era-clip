import { getOneOfficeTabOrigin, oneOfficeOrigin } from './origins';

export interface AccessRecord {
  id: string;
  account: string;
  name: string;
  timeRaw: string;
  timeIso: string;
}

export interface AccessPage {
  total: number;
  start: number;
  end: number;
  records: AccessRecord[];
}

export type AccessPageResult = { ok: true; data: AccessPage } | { ok: false; error: string };

export interface AccessReport {
  scope?: string;
  cutoff: string;
  collectedAt: string;
  recordCount: number;
  oldestTime: string;
  users: Array<AccessRecord & { eventCount: number }>;
  snapshot?: AccessSnapshot;
  personnel?: AccessPersonnel;
}

export interface AccessPersonnel {
  activeAccounts?: string[];
  collectedAt: string;
  users: Array<{ account: string; position: string; department: string; company?: string; companyCode?: string; companyGroup?: string; departmentCode?: string }>;
  warnings: string[];
}

export interface AccessSnapshot {
  scope?: string;
  collectedAt: string;
  recordCount: number;
  oldestTime: string;
  records: AccessRecord[];
  personnel?: AccessPersonnel;
}

export function reportFromAccessSnapshot(snapshot: AccessSnapshot, cutoff: string): AccessReport {
  const end = cutoffEnd(cutoff);
  const users = new Map<string, AccessRecord & { eventCount: number }>();
  for (const record of snapshot.records) {
    const time = Date.parse(record.timeIso);
    if (!Number.isFinite(time)) throw new Error('Nhật ký có thời gian không hợp lệ.');
    if (time >= end) continue;
    const previous = users.get(record.account);
    const latest = !previous || time > Date.parse(previous.timeIso) ||
      (time === Date.parse(previous.timeIso) && Number(record.id) > Number(previous.id)) ? record : previous;
    users.set(record.account, { ...latest, eventCount: (previous?.eventCount ?? 0) + 1 });
  }
  return { cutoff, collectedAt: snapshot.collectedAt, recordCount: snapshot.recordCount,
    oldestTime: snapshot.oldestTime, snapshot, personnel: snapshot.personnel, scope: snapshot.scope,
    users: [...users.values()].sort((a, b) => a.account.localeCompare(b.account, 'vi')) };
}

export function isAccessHistoryUrl(value: string): boolean {
  try {
    const url = new URL(value);
    return (oneOfficeOrigin(value) === url.origin) &&
      /^\/apps\/admin-reconciliation-histories(?:\/list)?\/?$/.test(url.pathname);
  } catch { return false; }
}

export function cutoffEnd(value: string): number {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) throw new Error('Hãy chọn ngày chốt hợp lệ.');
  const date = new Date(`${value}T00:00:00+07:00`);
  if (!Number.isFinite(date.getTime()) || new Date(date.getTime() + 7 * 3600000).toISOString().slice(0, 10) !== value) {
    throw new Error('Ngày chốt không hợp lệ.');
  }
  return date.getTime() + 86400000;
}

// This function is injected into the source tab, so all helpers must be self-contained.
export function readAccessHistoryPage(page: number): Promise<AccessPage>;
export function readAccessHistoryPage(page: number, captureErrors: true, searchAccount?: string | null, expectedOrigin?: string): Promise<AccessPageResult>;
export async function readAccessHistoryPage(page: number, captureErrors = false, searchAccount?: string | null, expectedOrigin = 'https://office.meygroup.vn'): Promise<AccessPage | AccessPageResult> {
  const read = async (): Promise<AccessPage> => {
  if (!['https://office.meygroup.vn', 'https://testoffice.meygroup.vn', 'https://testoffice.meygroup.vn:9443'].includes(expectedOrigin) || location.origin !== expectedOrigin ||
      !/^\/apps\/admin-reconciliation-histories(?:\/list)?\/?$/.test(location.pathname)) {
    throw new Error('Tab nguồn không còn ở Lịch sử truy cập 1Office.');
  }
  if (!Number.isInteger(page) || page < 1 || page > 2000) throw new Error('Trang nhật ký không hợp lệ.');
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 30000);
  try {
    const url = new URL('/apps/admin-reconciliation-histories/list', location.origin);
    url.searchParams.set('v', 'list');
    url.searchParams.set('p', String(page));
    url.searchParams.set('_json', '1');
    url.searchParams.set('inlineLogin', '1');
    if (searchAccount != null) {
      if (!searchAccount.trim() || searchAccount.length > 200) throw new Error('Tài khoản tìm kiếm không hợp lệ.');
      url.searchParams.set('s', searchAccount);
    }
    let response: Response;
    try {
      response = await fetch(url.href, { credentials: 'include', redirect: 'error', signal: controller.signal });
    } catch {
      throw new Error(controller.signal.aborted
        ? `Trang ${page}: 1Office không phản hồi trong 30 giây. Hãy thử lại.`
        : `Trang ${page}: không tải được nhật ký (lỗi kết nối hoặc chuyển hướng). Kiểm tra phiên đăng nhập 1Office.`);
    }
    if (!response.ok) throw new Error(`Không tải được trang ${page} (HTTP ${response.status}).`);
    const body = await response.text();
    if (/^\s*[\[{]/.test(body)) {
      throw new Error(`Trang ${page}: 1Office trả về dữ liệu dạng JSON, không phải bảng HTML mà bộ đọc hiện hỗ trợ.`);
    }
    const clean = (node: Element | null): string => (node?.textContent ?? '').replace(/\s+/g, ' ').trim();
    const findTable = (doc: Document): HTMLTableElement | undefined =>
      Array.from(doc.querySelectorAll('table')).find((candidate) => {
        const headers = Array.from(candidate.querySelector('tr')?.children ?? []).map(clean);
        return ['Tài khoản', 'Họ và tên', 'Thời gian'].every((label) => headers.includes(label));
      });
    const doc = new DOMParser().parseFromString(body, 'text/html');
    const parseRecord = (id: string, account: string, name: string, timeRaw: string): AccessRecord => {
      const parts = /^(\d{2}):(\d{2}):(\d{2}) (\d{2})\/(\d{2})\/(\d{4})$/.exec(timeRaw);
      if (!account || !parts || !/^\d+$/.test(id) || !Number.isSafeInteger(Number(id))) {
        throw new Error('Bản ghi nhật ký không đúng định dạng; đã dừng để tránh báo cáo sai.');
      }
      const [, hh, mm, ss, dd, month, yyyy] = parts;
      const local = `${yyyy}-${month}-${dd}T${hh}:${mm}:${ss}`;
      const date = new Date(`${local}+07:00`);
      if (!Number.isFinite(date.getTime()) || new Date(date.getTime() + 7 * 3600000).toISOString().slice(0, 19) !== local) {
        throw new Error('Thời gian nhật ký không hợp lệ.');
      }
      return { id, account, name, timeRaw, timeIso: date.toISOString() };
    };
    // Parse the site's serialized table as data only. Never run its event scripts.
    const serialized = Array.from(doc.querySelectorAll('elem[name="table"][data-props]'));
    if (serialized.length) {
      if (serialized.length !== 1) throw new Error('Nguồn trả về nhiều bảng nhật ký; không thể xác nhận dữ liệu.');
      let props;
      try { props = JSON.parse(serialized[0].getAttribute('data-props') ?? ''); }
      catch { throw new Error('Cấu trúc bảng nhật ký không hợp lệ.'); }
      const paging = props?.page;
      if (searchAccount != null && paging?.params?.s !== searchAccount && props?.vars?.s !== searchAccount) {
        throw new Error('1Office không xác nhận bộ lọc tài khoản. Không đọc toàn bộ nhật ký thay thế.');
      }
      if (props?.object !== 'admin-reconciliation-histories' || !Array.isArray(props.rows) ||
          !Array.isArray(props.headers) || !['username', 'fullname', 'date'].every((key) =>
            props.headers.some((header: { name?: unknown }) => header?.name === key)) ||
          !Number.isSafeInteger(paging?.total) || paging.total < 0 ||
          paging.current !== page || paging.range !== 50 || paging.rq !== 'p') {
        throw new Error('Thông tin bảng hoặc phân trang nhật ký không đúng định dạng đã xác minh.');
      }
      const records = props.rows.map((row: { ID?: unknown; vals?: Record<string, unknown> }) => {
        const values = row?.vals;
        if ((typeof row?.ID !== 'number' && typeof row?.ID !== 'string') ||
            !values || !['username', 'fullname', 'date'].every((key) => typeof values[key] === 'string')) {
          throw new Error('Bản ghi nhật ký thiếu thông tin bắt buộc.');
        }
        return parseRecord(String(row.ID), (values.username as string).trim(),
          (values.fullname as string).trim(), (values.date as string).trim());
      });
      const total = paging.total;
      const start = total ? (page - 1) * 50 + 1 : 0;
      const end = Math.min(page * 50, total);
      if (records.length !== (total ? end - start + 1 : 0) ||
          (total === 0 && page !== 1) || new Set(records.map((row: AccessRecord) => row.id)).size !== records.length) {
        throw new Error(`Trang ${page} chưa trả đủ bản ghi hoặc có ID trùng.`);
      }
      return { total, start, end, records };
    }
    if (searchAccount != null) throw new Error('Không xác minh được bộ lọc tài khoản trên nguồn nhật ký.');
    const table = findTable(doc);
    if (!table && doc.querySelector('input[type="password"]')) {
      throw new Error(`Trang ${page}: 1Office trả về màn hình đăng nhập. Hãy đăng nhập lại trên tab nguồn.`);
    }
    if (!table) {
      throw new Error(`Trang ${page}: không nhận diện được dữ liệu bảng nhật ký. Không mở trang ngầm để tránh quá tải trình duyệt.`);
    }
    const headerRow = table.querySelector('tr');
    const headers = Array.from(headerRow?.children ?? []).map(clean);
    const paginationText = [doc.body.textContent ?? '', ...Array.from(doc.querySelectorAll('[title]')).map((node) => node.getAttribute('title') ?? '')].join(' ');
    const count = /Hiển thị\s*([\d,.]+)\s*-\s*([\d,.]+)\s*\/\s*([\d,.]+)\s*bản ghi/i.exec(paginationText);
    if (!count) throw new Error('Không đọc được tổng số bản ghi nhật ký; không thể xác nhận đủ dữ liệu.');
    const [start, end, total] = count.slice(1).map((value) => Number(value.replace(/[,.]/g, '')));
    if (![start, end, total].every((value) => Number.isSafeInteger(value) && value >= 0)) {
      throw new Error('Bộ đếm nhật ký không hợp lệ.');
    }
    const records: AccessRecord[] = [];
    for (const row of Array.from(table.querySelectorAll('tbody tr'))) {
      if (row === headerRow) continue;
      const id = row.querySelector<HTMLInputElement>('input[name="ID[]"]')?.value;
      if (!id) {
        if (row.querySelector('td') && total > 0) throw new Error('Bản ghi nhật ký thiếu ID ổn định.');
        continue;
      }
      const cells = Array.from(row.querySelectorAll(':scope > td'));
      const account = clean(cells[headers.indexOf('Tài khoản')]);
      const name = clean(cells[headers.indexOf('Họ và tên')]);
      const timeRaw = clean(cells[headers.indexOf('Thời gian')]);
      records.push(parseRecord(id, account, name, timeRaw));
    }
    if (total > 0 && (start !== (page - 1) * 50 + 1 || end !== Math.min(page * 50, total) || records.length !== end - start + 1)) {
      throw new Error(`Trang ${page} không đúng phạm vi 50 bản ghi hoặc chưa tải đủ.`);
    }
    if (total === 0 && (page !== 1 || records.length !== 0)) throw new Error('Số liệu nhật ký không nhất quán.');
    return { start, end, total, records };
  } finally { clearTimeout(timeout); }
  };
  try {
    const data = await read();
    return captureErrors ? { ok: true, data } : data;
  } catch (error) {
    if (!captureErrors) throw error;
    return { ok: false, error: error instanceof Error ? error.message : `Không thể đọc trang nhật ký ${page}.` };
  }
}

export async function collectAccessReport(
  cutoff: string,
  readPage: (page: number) => Promise<AccessPage>,
  signal: AbortSignal,
  progress: (loaded: number, total: number) => void,
): Promise<AccessReport> {
  cutoffEnd(cutoff);
  const records = new Map<string, AccessRecord>();
  let oldest: AccessRecord | undefined;
  const checkCancelled = () => { if (signal.aborted) throw new Error('Đã hủy thu thập. Chưa có báo cáo để xuất.'); };
  checkCancelled();
  const first = await readPage(1);
  if (!Number.isSafeInteger(first.total) || first.total < 0) throw new Error('Tổng nhật ký không hợp lệ.');
  let pages = Math.max(1, Math.ceil(first.total / 50));
  const descending = (rows: AccessRecord[]) => rows.every((row, index) =>
    /^\d+$/.test(row.id) && Number.isSafeInteger(Number(row.id)) &&
    (index === 0 || Number(row.id) < Number(rows[index - 1].id)));
  const anchored = first.records.length > 0 && descending(first.records);
  const anchor = Number(first.records[0]?.id ?? 0);
  let lowestId = Infinity;
  if (pages > 2000) throw new Error('Nhật ký vượt giới hạn an toàn 100.000 bản ghi của bản thử nghiệm.');
  const accept = (data: AccessPage, page: number) => {
    checkCancelled();
    if ((!anchored && data.total !== first.total) || data.total < first.total || (data.total > 0 &&
      (data.start !== (page - 1) * 50 + 1 || data.end !== Math.min(page * 50, data.total) || data.records.length !== data.end - data.start + 1))) {
      throw new Error('Phân trang hoặc tổng nhật ký đã thay đổi trong khi thu thập. Hãy thử lại.');
    }
    pages = Math.max(pages, Math.ceil(data.total / 50));
    if (pages > 2000) throw new Error('Nhật ký vượt giới hạn an toàn 100.000 bản ghi.');
    if (anchored && !descending(data.records)) throw new Error('Thứ tự ID nhật ký thay đổi. Không thể đối chiếu dữ liệu.');
    for (const record of data.records) {
      if (anchored && Number(record.id) > anchor) continue;
      const existing = records.get(record.id);
      if (existing) {
        if (!anchored || JSON.stringify(existing) !== JSON.stringify(record)) {
          throw new Error('Có bản ghi trùng hoặc bị sửa giữa các trang. Không thể đối chiếu nhật ký.');
        }
        continue;
      }
      if (anchored && Number(record.id) >= lowestId) throw new Error('Nhật ký bị đổi thứ tự trong khi quét.');
      lowestId = Number(record.id);
      records.set(record.id, record);
      const time = Date.parse(record.timeIso);
      if (!Number.isFinite(time)) throw new Error('Nhật ký có thời gian không hợp lệ.');
      if (!oldest || time < Date.parse(oldest.timeIso)) oldest = record;
    }
    if (records.size > first.total) throw new Error('Số bản ghi tại mốc bắt đầu không nhất quán.');
    progress(records.size, first.total);
  };
  accept(first, 1);
  // A rolling, bounded window avoids waiting for every page in a fixed batch.
  const pending = new Map<number, Promise<PromiseSettledResult<AccessPage>>>();
  let next = 2;
  const fill = () => {
    while (!signal.aborted && pending.size < 6 && next <= pages) {
      const number = next++;
      pending.set(number, Promise.resolve().then(() => readPage(number)).then(
        (value) => ({ status: 'fulfilled' as const, value }),
        (reason) => ({ status: 'rejected' as const, reason }),
      ));
    }
  };
  try {
    fill();
    for (let page = 2; page <= pages && records.size < first.total; page++) {
      checkCancelled();
      const result = await pending.get(page)!;
      pending.delete(page);
      checkCancelled();
      if (result.status === 'rejected') throw result.reason;
      accept(result.value, page);
      fill();
    }
  } finally {
    await Promise.all(pending.values());
  }
  checkCancelled();
  const final = await readPage(1);
  checkCancelled();
  if (records.size !== first.total || final.total < first.total || (!anchored &&
    (final.total !== first.total || JSON.stringify(final.records) !== JSON.stringify(first.records)))) {
    throw new Error('Nhật ký thay đổi hoặc thiếu dữ liệu khi đối chiếu cuối. Không thể xuất báo cáo đầy đủ.');
  }
  if (anchored) {
    // Verify that growth is explained by newer IDs, not a missing/replaced history.
    const growth = final.total - first.total;
    if (growth > 950) throw new Error('Nhật ký phát sinh quá nhanh để đối chiếu an toàn. Hãy thử lúc ít truy cập.');
    const head: AccessRecord[] = [];
    for (let page = 1; page <= Math.ceil((growth + first.records.length) / 50); page++) {
      checkCancelled();
      const data = page === 1 ? final : await readPage(page);
      if (data.total !== final.total || !descending(data.records) || data.start !== (page - 1) * 50 + 1 ||
        data.end !== Math.min(page * 50, data.total) || data.records.length !== data.end - data.start + 1) {
        throw new Error('Nhật ký tiếp tục thay đổi khi đối chiếu cuối. Hãy thử lại.');
      }
      head.push(...data.records);
    }
    const newer = head.filter((record) => Number(record.id) > anchor);
    const original = head.filter((record) => Number(record.id) <= anchor).slice(0, first.records.length);
    if (newer.length !== growth || JSON.stringify(original) !== JSON.stringify(first.records)) {
      throw new Error('Không xác nhận được mốc nhật ký ban đầu. Không xuất dữ liệu thiếu.');
    }
  }
  checkCancelled();
  return reportFromAccessSnapshot({ collectedAt: new Date().toISOString(), recordCount: records.size,
    oldestTime: oldest?.timeRaw ?? '', records: [...records.values()] }, cutoff);
}

export async function loadAccessPageFromTab(tabId: number, page: number, searchAccount?: string, expectedOrigin = 'https://office.meygroup.vn'): Promise<AccessPage> {
  const chromeApi = (globalThis as any).chrome;
  if (!chromeApi?.scripting?.executeScript) throw new Error('Không có quyền đọc tab nguồn. Hãy tải lại extension.');
  const results = await chromeApi.scripting.executeScript({ target: { tabId }, func: readAccessHistoryPage,
    args: [page, true, searchAccount ?? null, oneOfficeOrigin(expectedOrigin)] });
  const result = results[0]?.result as AccessPageResult | undefined;
  if (!result) throw new Error(`Không nhận được kết quả từ bộ đọc trang ${page}. Tải lại extension, mở lại báo cáo từ tab Lịch sử truy cập.`);
  if (!result.ok) throw new Error(result.error || `Không thể đọc trang nhật ký ${page}.`);
  if (!result.data || !Array.isArray(result.data.records) || !Number.isSafeInteger(result.data.total)) {
    throw new Error(`Kết quả bộ đọc trang ${page} không hợp lệ. Không thể xác nhận đủ nhật ký.`);
  }
  return result.data;
}

export async function openAccessReview(tabId: number, expectedOrigin?: string): Promise<void> {
  const chromeApi = (globalThis as any).chrome;
  if (!chromeApi?.tabs?.create || !chromeApi?.runtime?.getURL) throw new Error('Không thể mở trang báo cáo.');
  const origin = await getOneOfficeTabOrigin(tabId, expectedOrigin);
  await chromeApi.tabs.create({ url: chromeApi.runtime.getURL(`access-review.html?sourceTabId=${tabId}&sourceOrigin=${encodeURIComponent(origin)}`) });
}
