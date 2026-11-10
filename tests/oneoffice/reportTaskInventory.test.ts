import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { collectOneOfficeReportInventory, readOneOfficeReportFromPage, readOneOfficeReportIdentityFromPage, readReportInjectionResult } from '../../src/oneoffice/reportTaskInventory';
import { isOneOfficeReportPageUrl } from '../../src/oneoffice/collectionClient';

const reportPath = '/apps/report-report/view?ID=74&from=01/10/2026&to=05/11/2026&department_ids[]=424';
const headers = ['No.', 'Mã công việc', 'Tên công việc', 'Người thực hiện', 'Phòng ban thực hiện', 'Dự án', 'Thảo luận'];
const cells = ['1', 'VPTGD.18526', 'Báo cáo tuần', 'Người A', 'Phòng A', 'Dự án A', 'Nội dung tóm tắt, không dùng để xuất'];
function report(extraHeaders: string[] = [], extraCells: string[] = []) {
  document.body.innerHTML = `<table _n="table"><tr><td></td></tr><tr>${[...headers, ...extraHeaders].map(h => `<td>${h}</td>`).join('')}</tr><tr>${[...cells, ...extraCells].map(c => `<td>${c}</td>`).join('')}</tr><tr>${headers.map(() => '<td></td>').join('')}</tr></table>`;
}
const options = () => ({ allowedOrigin: window.location.origin, sourceUrl: window.location.href });
beforeEach(() => {
  window.history.replaceState({}, '', reportPath);
  vi.spyOn(HTMLElement.prototype, 'getClientRects').mockReturnValue([{}] as any);
  report();
});
afterEach(() => { vi.restoreAllMocks(); vi.unstubAllGlobals(); vi.useRealTimers(); });

describe('report table snapshot', () => {
  it('reads the filtered table without interpreting code suffixes as IDs or using summary comments', () => {
    const snapshot = readOneOfficeReportFromPage(options());
    expect(snapshot.rows).toHaveLength(1);
    expect(snapshot.rows[0]).toMatchObject({ explicitUrl: null, task: { internalId: '', code: 'VPTGD.18526',
      name: 'Báo cáo tuần', no: '1', projectName: 'Dự án A', departmentNames: ['Phòng A'] } });
    expect(JSON.stringify(snapshot)).not.toContain('Nội dung tóm tắt');
  });
  it('accepts only explicit task identity, not the report ID', () => {
    report(['ID công việc'], ['89876']);
    expect(readOneOfficeReportFromPage(options()).rows[0].explicitUrl).toContain('ID=89876');
    report(['ID'], ['74']);
    expect(readOneOfficeReportFromPage(options()).rows[0].explicitUrl).toBeNull();
  });
  it('rejects filter drift and ambiguous tables but retains rows with missing codes', () => {
    expect(() => readOneOfficeReportFromPage({ ...options(), sourceUrl: 'https://example.com/' })).toThrow(/bộ lọc/);
    document.querySelectorAll('tr')[2].cells[1].textContent = '';
    expect(readOneOfficeReportFromPage(options()).rows[0].task).toMatchObject({ internalId: '', collectionKey: 'report-row:0', identityError: expect.stringContaining('Chưa lấy được thảo luận') });
    report(); document.body.append(document.querySelector('table')!.cloneNode(true));
    expect(() => readOneOfficeReportFromPage(options())).toThrow(/duy nhất/);
  });
  it('rejects foreign task links as identity sources', () => {
    document.querySelectorAll('tr')[2].cells[2].innerHTML = '<a href="https://evil.example/apps/work-task-task/view?ID=89876">Báo cáo tuần</a>';
    expect(readOneOfficeReportFromPage(options()).rows[0].explicitUrl).toBeNull();
  });
  it('returns a serializable error when injected into Chrome instead of losing the original exception', () => {
    document.body.innerHTML = '';
    const snapshot = readOneOfficeReportFromPage({ ...options(), captureErrors: true });
    expect(snapshot.rows).toEqual([]);
    expect(snapshot.error).toContain('nhận diện duy nhất');
    expect(() => readReportInjectionResult([{ result: JSON.parse(JSON.stringify(snapshot)) }])).toThrow(/nhận diện duy nhất/);
  });
  it('allows an empty code only with explicit identity, never by guessing a name match', () => {
    report(['ID công việc'], ['89876']);
    document.querySelectorAll('tr')[2].cells[1].textContent = '';
    expect(readOneOfficeReportFromPage(options()).rows[0].explicitUrl).toContain('ID=89876');
  });
  it('only recognizes allowlisted report view URLs with one ID', () => {
    expect(isOneOfficeReportPageUrl(`https://office.meygroup.vn${reportPath}`)).toBe(true);
    expect(isOneOfficeReportPageUrl('https://testoffice.meygroup.vn:9443/apps/report-report/view?ID=74')).toBe(true);
    for (const url of ['https://evil.example/apps/report-report/view?ID=74', 'https://office.meygroup.vn/apps/report-report/edit?ID=74',
      'https://office.meygroup.vn/apps/report-report/view?ID=74&ID=75']) expect(isOneOfficeReportPageUrl(url)).toBe(false);
  });
});

describe('identity verification', () => {
  it('preserves detail lookup errors across the injection boundary', async () => {
    const result = await readOneOfficeReportIdentityFromPage({ ...options(), expectedUrl: 'https://example.com/',
      code: 'A', name: 'A', search: false, captureErrors: true });
    expect(result.error).toContain('chuyển trang');
    expect(() => readReportInjectionResult([{result}])).toThrow(/chuyển trang/);
  });
  it('uses the real search link ID, then requires exact code on the detail page', async () => {
    vi.useFakeTimers();
    window.history.replaceState({}, '', '/apps/work-task-task/list?s=VPTGD.18526');
    document.body.innerHTML = '<div class="table-page-title">Hiển thị 1 - 1 / 1 bản ghi</div><table><tr class="table-row-body"><td><a class="table-main-title" href="/apps/work-normal-normal/view?ID=89876">Báo cáo tuần</a></td></tr></table>';
    const base = { allowedOrigin: window.location.origin, expectedUrl: window.location.href, code: 'VPTGD.18526', name: 'Báo cáo tuần' };
    const search = readOneOfficeReportIdentityFromPage({ ...base, search: true });
    await vi.advanceTimersByTimeAsync(600);
    expect(await search).toEqual({ url: `${window.location.origin}/apps/work-normal-normal/view?ID=89876` });
    window.history.replaceState({}, '', '/apps/work-normal-normal/view?ID=89876');
    document.body.innerHTML = '<div class="header-right">Báo cáo tuần</div><div class="detail-field"><div class="detail-field-label">Mã công việc</div><div class="detail-field-content">VPTGD.18526</div></div><div class="cards comments" comments-id="work-normal-normal@89876"></div>';
    const verified = readOneOfficeReportIdentityFromPage({ ...base, expectedUrl: window.location.href, search: false });
    await vi.advanceTimersByTimeAsync(600);
    expect((await verified).url).toContain('ID=89876');
  });
  it('fails rather than choosing from multiple search results', async () => {
    vi.useFakeTimers();
    document.body.innerHTML = '<div class="table-page-title">Hiển thị 1 - 2 / 2 bản ghi</div>';
    const pending = expect(readOneOfficeReportIdentityFromPage({ ...options(), expectedUrl: window.location.href,
      code: 'VPTGD.18526', name: 'Báo cáo tuần', search: true })).rejects.toThrow(/Không xác minh duy nhất/);
    await vi.advanceTimersByTimeAsync(15100); await pending;
  });
});

describe('report orchestration', () => {
  it('retains the entire report without navigation when all rows have no code or ID', async () => {
    document.querySelectorAll('tr')[2].cells[1].textContent = '';
    const snapshot = readOneOfficeReportFromPage(options());
    const sourceUrl = 'https://office.meygroup.vn/apps/report-report/view?ID=74';
    const update = vi.fn(); const executeScript = vi.fn(async () => [{result: snapshot}]);
    vi.stubGlobal('chrome', { tabs: {get: async () => ({url: sourceUrl}), update}, scripting: {executeScript} });
    const result = await collectOneOfficeReportInventory(5, 'https://office.meygroup.vn', sourceUrl);
    expect(result.tasks).toHaveLength(1);
    expect(result.tasks[0]).toMatchObject({internalId: '', collectionKey: 'report-row:0', identityError: expect.any(String)});
    expect(result.collection.complete).toBe(true);
    expect(result.collection.warnings.join(' ')).toContain('1 công việc chưa xác minh');
    expect(executeScript).toHaveBeenCalledTimes(1); expect(update).not.toHaveBeenCalled();
  });
  it.each([null, [], [{result: null}], [{result: undefined}]])('rejects absent injection data without reading rows or navigating: %j', async (response) => {
    const sourceUrl = 'https://office.meygroup.vn/apps/report-report/view?ID=74';
    const update = vi.fn();
    vi.stubGlobal('chrome', { tabs: { get: vi.fn(async () => ({url: sourceUrl})), update },
      scripting: { executeScript: vi.fn(async () => response) } });
    await expect(collectOneOfficeReportInventory(5, 'https://office.meygroup.vn', sourceUrl)).rejects.toThrow(/không trả về dữ liệu/);
    expect(update).not.toHaveBeenCalled();
  });
  it('shows the original page error before starting any task navigation', async () => {
    const sourceUrl = 'https://office.meygroup.vn/apps/report-report/view?ID=74';
    const update = vi.fn();
    vi.stubGlobal('chrome', { tabs: { get: vi.fn(async () => ({url: sourceUrl})), update },
      scripting: { executeScript: vi.fn(async () => [{result: {title: '', rows: [], error: 'Dòng 113 chưa có mã'}}]) } });
    await expect(collectOneOfficeReportInventory(5, 'https://office.meygroup.vn', sourceUrl)).rejects.toThrow('Dòng 113 chưa có mã');
    expect(update).not.toHaveBeenCalled();
  });
  it('only freezes the exact filtered report in the popup, without navigating or resolving any task', async () => {
    const snapshot = readOneOfficeReportFromPage(options());
    const origin = window.location.origin; const sourceUrl = window.location.href;
    let currentUrl = sourceUrl; const listeners = new Set<(id: number, change: {status: string}) => void>();
    const executeScript = vi.fn().mockResolvedValueOnce([{ result: snapshot }])
      .mockResolvedValueOnce([{ result: { url: `${origin}/apps/work-normal-normal/view?ID=89876` } }])
      .mockResolvedValueOnce([{ result: { url: `${origin}/apps/work-normal-normal/view?ID=89876`, projectInternalId: '130' } }]);
    // Use the real allowlisted origin in the orchestration boundary.
    const allowed = 'https://office.meygroup.vn';
    currentUrl = sourceUrl.replace(origin, allowed);
    snapshot.rows[0].explicitUrl = null;
    executeScript.mockReset().mockResolvedValueOnce([{ result: snapshot }])
      .mockResolvedValueOnce([{ result: { url: `${allowed}/apps/work-normal-normal/view?ID=89876` } }])
      .mockResolvedValueOnce([{ result: { url: `${allowed}/apps/work-normal-normal/view?ID=89876`, projectInternalId: '130' } }]);
    const update = vi.fn((_id: number, input: {url: string}, done: () => void) => {
      currentUrl = input.url; done(); listeners.forEach(listener => listener(5, {status: 'complete'}));
    });
    vi.stubGlobal('chrome', { scripting: { executeScript }, runtime: {}, tabs: {
      get: vi.fn(async () => ({url: currentUrl})), update,
      onUpdated: { addListener: (l: any) => listeners.add(l), removeListener: (l: any) => listeners.delete(l) },
    } });
    const result = await collectOneOfficeReportInventory(5, allowed, sourceUrl.replace(origin, allowed));
    expect(result.project.type).toBe('report'); expect(result.collection.complete).toBe(true);
    expect(result.tasks[0]).toMatchObject({ internalId: '', collectionKey: 'report-row:0' });
    expect(result.project.url).toBe(sourceUrl.replace(origin, allowed));
    expect(executeScript).toHaveBeenCalledTimes(1);
    expect(update).not.toHaveBeenCalled();
    expect(listeners.size).toBe(0);
  });
});
