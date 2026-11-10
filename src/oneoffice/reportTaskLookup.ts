import type { OneOfficeProjectTaskInventoryItem } from './types';
import { getOneOfficeTabOrigin } from './origins';
import { readReportInjectionResult } from './reportTaskInventory';

interface IdentityWidget { tabId: number; widgetId: number }
type LookupResult = OneOfficeProjectTaskInventoryItem & { identityWidget?: IdentityWidget };
const identityWidgets = new Map<string, IdentityWidget>();
const identityWarmups = new Map<string, Promise<void>>();
const discoveryAttempted = new Set<string>();

// Injected read-only worker. Parse server markup as inert data, never run its scripts.
export async function lookupReportTaskFromPage(options: {sourceUrl: string; task: OneOfficeProjectTaskInventoryItem; widgetHints?: Record<string, IdentityWidget>}): Promise<LookupResult> {
  const task = options.task;
  const source = new URL(options.sourceUrl);
  const clean = (value: string | null | undefined) => (value ?? '').replace(/\s+/g, ' ').trim();
  // Compare only bare fractional code representations, with no floating-point conversion.
  const fractionalCode = (value: string): string | null => {
    const match = /^(?:0)?\.(\d+)$/.exec(value);
    return match && /[1-9]/.test(match[1]) ? match[1].replace(/0+$/, '') : null;
  };
  const labelText = (root: Node): string => Array.from(root.childNodes).map(node => {
    if (node.nodeType === 3) return node.textContent ?? '';
    if (node.nodeName === 'TEMPLATE') return labelText((node as HTMLTemplateElement).content);
    if (['SCRIPT', 'STYLE', 'IFRAME', 'OBJECT'].includes(node.nodeName)) return '';
    return labelText(node);
  }).join('');
  // 1Office nests server-rendered elements in inert template.content fragments.
  const select = (root: Document | DocumentFragment, selector: string): Element[] => {
    const matches: Element[] = [];
    const fragments: Array<Document | DocumentFragment> = [root];
    for (let index = 0; index < fragments.length; index++) {
      if (index > 10000) throw new Error('Cấu trúc phản hồi quá phức tạp.');
      const fragment = fragments[index];
      matches.push(...fragment.querySelectorAll(selector));
      for (const template of fragment.querySelectorAll('template')) fragments.push(template.content);
    }
    return matches;
  };
  try {
    if (location.href !== source.href || source.pathname !== '/apps/report-report/view') throw new Error('Tab báo cáo hoặc bộ lọc đã thay đổi.');
    const readText = async (url: URL): Promise<string> => {
      if (url.origin !== source.origin || url.username || url.password || location.href !== source.href) throw new Error('Nguồn đọc không còn khớp báo cáo.');
      const controller = new AbortController(); const timeout = setTimeout(() => controller.abort(), 30000);
      try {
        const response = await fetch(url.href, {method: 'GET', credentials: 'include', redirect: 'error', signal: controller.signal,
          headers: {'X-Requested-With': 'XMLHttpRequest', Accept: 'text/html, application/json'}});
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        const body = await response.text();
        if (body.length > 8_000_000) throw new Error('Phản hồi quá lớn.');
        if (location.href !== source.href) throw new Error('Bộ lọc báo cáo đã thay đổi.');
        return body;
      } finally { clearTimeout(timeout); }
    };
    const read = async (url: URL): Promise<Document> => {
      let body = await readText(url);
      if (body.trimStart().startsWith('{')) {
        const data = JSON.parse(body);
        if (typeof data.content !== 'string') throw new Error('Phản hồi không có nội dung bảng.');
        body = data.content;
      }
      const doc = new DOMParser().parseFromString(body, 'text/html');
      if (select(doc, 'input[type="password"]').length) throw new Error('Phiên đăng nhập đã hết.');
      return doc;
    };
    const validTaskUrl = (href: string): URL => {
      const url = new URL(href, source.origin);
      if (url.origin !== source.origin || !/^\/apps\/work-(normal-normal|task-task)\/view$/.test(url.pathname)
        || url.searchParams.getAll('ID').length !== 1 || !/^[1-9]\d*$/.test(url.searchParams.get('ID') ?? '')) throw new Error('Liên kết công việc không hợp lệ.');
      return url;
    };
    const linkFrom = (html: string): {href: string; name: string} | null => {
      const doc = new DOMParser().parseFromString(html, 'text/html');
      const link = select(doc, 'a.table-main-title[href]')[0] ?? select(doc, 'a[href]')[0];
      if (link) return {href: link.getAttribute('href')!, name: clean(labelText(link))};
      const elements = select(doc, 'elem[name="a"][data-atts]');
      const main = elements.filter(element => clean(JSON.parse(element.getAttribute('data-atts')!).class).split(' ').includes('table-main-title'));
      for (const element of main.length ? main : elements) {
        const attrs = JSON.parse(element.getAttribute('data-atts')!);
        const props = JSON.parse(element.getAttribute('data-props') || '{}');
        const children = JSON.parse(element.getAttribute('data-children') || 'null');
        const content = typeof props.content === 'string' ? props.content
          : Array.isArray(children) && children.every(value => typeof value === 'string') ? children.join('') : null;
        const name = content !== null ? clean(labelText(new DOMParser().parseFromString(content, 'text/html').body)) : clean(labelText(element));
        if (typeof attrs.href === 'string' && name) return {href: attrs.href, name};
      }
      return null;
    };
    let url: URL | undefined;
    let verifiedLookupCode = task.code ?? '';
    if (task.reportLookupUrl) url = validTaskUrl(task.reportLookupUrl);
    else {
      if (!task.code) throw new Error('Báo cáo chưa cung cấp ID/liên kết để nối đúng công việc.');
      // Approved representation aliases only; query results must still supply the real ID.
      const fraction = fractionalCode(task.code);
      const queries = fraction === null ? [task.code]
        : [...new Set([task.code, task.code.startsWith('0.') ? task.code.slice(1) : task.code, `.${fraction}`])];
      for (const [queryIndex, query] of queries.entries()) {
        const search = new URL('/apps/work-task-task/list', source.origin);
        search.searchParams.set('v', 'list'); search.searchParams.set('tab', 'all_all'); search.searchParams.set('s', query);
        search.searchParams.set('_json', '1');
        const doc = await read(search);
        const links: Array<{href: string; name: string}> = [];
        let total: number | null = null;
        let loadedRows = 0;
        const count = /\/\s*([\d,.]+)\s*bản ghi/i.exec(clean(select(doc, '.table-page-title')[0]?.textContent));
        if (count) {
          total = Number(count[1].replace(/[,.]/g, ''));
          loadedRows = select(doc, 'tr.table-row-body').length;
          for (const link of select(doc, 'tr.table-row-body a.table-main-title[href]')) links.push({href: link.getAttribute('href')!, name: clean(labelText(link))});
        } else {
          const tables = select(doc, 'elem[name="table"][data-props]').map(el => JSON.parse(el.getAttribute('data-props')!))
            .filter(props => props.object === 'work-task-task' || props.object === 'work-normal-normal');
          if (tables.length !== 1) throw new Error('Chưa đọc được bảng tra ID bằng request ngầm.');
          const props = tables[0];
          if (props.page?.params?.s !== query && props.vars?.s !== query) throw new Error('Chưa xác nhận được truy vấn tra ID.');
          total = props.page?.total;
          if (!Array.isArray(props.rows)) throw new Error('Thiếu các dòng kết quả tra ID.');
          loadedRows = props.rows.length;
          for (const row of props.rows) {
            if (typeof row.vals?.title !== 'string') throw new Error('Thiếu định danh trong kết quả tra ID.');
            const link = linkFrom(typeof row.fields?.title === 'string' ? row.fields.title : row.vals.title);
            if (link) {
              if (validTaskUrl(link.href).searchParams.get('ID') !== String(row.ID)) throw new Error('ID và liên kết không khớp.');
              if (typeof row.fields?.title === 'string' && clean(row.vals.title) !== link.name) throw new Error('Tên và liên kết trong kết quả tra ID không khớp.');
              if (typeof row.url === 'string' && validTaskUrl(row.url.startsWith('apps/') ? `/${row.url}` : row.url).href !== validTaskUrl(link.href).href) throw new Error('Các liên kết định danh không khớp.');
              links.push(link);
            }
          }
        }
        if (total === 0 && queryIndex < queries.length - 1) continue;
        if (!Number.isSafeInteger(total) || total === null || total <= 0) throw new Error(`Tra ID trả ${total ?? 'không rõ số'} kết quả, chưa tìm được công việc.`);
        if (loadedRows !== total) throw new Error('Kết quả tra ID còn phân trang hoặc chưa đầy đủ. Chưa chọn ứng viên.');
        const matching = links.filter(link => link.name === clean(task.name));
        if (matching.length !== 1) throw new Error(`Đọc được ${matching.length} liên kết khớp tên trong ${total} kết quả, cần đúng 1 liên kết. Chưa xác minh mã chi tiết.`);
        url = validTaskUrl(matching[0].href);
        verifiedLookupCode = query;
        break;
      }
    }
    if (!url) throw new Error('Chưa xác minh được liên kết công việc.');
    const codes: string[] = [];
    const readCodes = (doc: Document) => {
      for (const field of select(doc, '.detail-field')) {
        if (clean(field.querySelector('.detail-field-label')?.textContent) === 'Mã công việc') codes.push(clean(field.querySelector('.detail-field-content')?.textContent));
      }
      for (const field of select(doc, 'elem[name="detail.field"][data-props]')) {
        const props = JSON.parse(field.getAttribute('data-props')!);
        if (props.label === 'Mã công việc' && typeof props.content === 'string') codes.push(clean(new DOMParser().parseFromString(props.content, 'text/html').body.textContent));
      }
    };
    const object = url.pathname.split('/')[2];
    const taskId = url.searchParams.get('ID')!;
    let identityWidget: IdentityWidget | undefined;
    const readIdentityWidget = async (plan: IdentityWidget): Promise<boolean> => {
      const widgetUrl = new URL('/features/core-app-features-board-objects-widgets/data', source.origin);
      Object.entries({ID: String(plan.widgetId), module: 'work', mode: 'view', object, board_type: 'detail', tab: String(plan.tabId),
        name: `visual-${plan.tabId}`, object_id: taskId, _json: '1', inlineLogin: '1'})
        .forEach(([key, value]) => widgetUrl.searchParams.set(key, value));
      const widgetDoc = await read(widgetUrl);
      const widgets = select(widgetDoc, 'elem[name="visual.widget"][data-props]').map(el => JSON.parse(el.getAttribute('data-props')!))
        .filter(p => p.name === 'Work-Task-TaskNormalInformation');
      if (widgets.length === 0) return false;
      if (widgets.length !== 1) throw new Error('Khối mã công việc không duy nhất.');
      const widget = widgets[0];
      if (String(widget.postId) !== taskId || widget.board_id !== plan.widgetId || widget.url !== `/apps/${object}/view/widget`
        || typeof widget.body !== 'string' || widget.isAjax !== false) throw new Error('Khối mã công việc không khớp ID hoặc cấu trúc.');
      readCodes(new DOMParser().parseFromString(widget.body, 'text/html'));
      if (codes.length === 0) return false;
      identityWidget = plan;
      return true;
    };
    const hint = options.widgetHints?.[object];
    const validHint = hint && Number.isSafeInteger(hint.tabId) && hint.tabId > 0 && Number.isSafeInteger(hint.widgetId) && hint.widgetId > 0;
    // Reuse only the widget address; its returned task ID and code are checked every time.
    if (validHint) await readIdentityWidget(hint);
    if (codes.length === 0) {
      const detailUrl = new URL(url.href);
      detailUrl.searchParams.set('_json', '1');
      const detail = await read(detailUrl);
      readCodes(detail);
      if (codes.length === 0) {
        const boards = select(detail, 'elem[name="board"][data-props]').map(el => JSON.parse(el.getAttribute('data-props')!))
          .filter(p => p.object === object && String(p.post?.ID) === taskId && p.module === 'work' && p.page === 'detail' && p.mode === 'view');
        if (boards.length !== 1) throw new Error('Chưa xác nhận được khung thông tin đúng ID công việc.');
        const board = boards[0];
        const tabId = board.curentTabId;
        if (!Number.isSafeInteger(tabId) || tabId <= 0 || !Array.isArray(board.tabs)
          || board.tabs.filter((tab: any) => tab.ID === tabId && tab.object === object && tab.module === 'work').length !== 1) throw new Error('Chưa xác nhận được tab thông tin công việc.');
        let candidates: number[] = [];
        const widgetsUrl = new URL('/features/core-app-features-board-objects-tabs/widgets', source.origin);
        Object.entries({ID: String(tabId), module: 'work', object, object_id: taskId, _json: '1', inlineLogin: '1'})
          .forEach(([key, value]) => widgetsUrl.searchParams.set(key, value));
        const data = JSON.parse(await readText(widgetsUrl));
        if (data.error !== false || !data.widgets || typeof data.widgets !== 'object' || Array.isArray(data.widgets)) throw new Error('Không đọc được danh sách khối thông tin.');
        const pending: any[] = Object.values(data.widgets).flat();
        for (let index = 0; index < pending.length; index++) {
          if (index >= 64) throw new Error('Quá nhiều khối thông tin để xác minh.');
          const item = pending[index];
          if (!item || item.boardId !== tabId || item.module !== 'work') throw new Error('Khối thông tin không khớp tab công việc.');
          if (Array.isArray(item.items)) pending.push(...item.items);
          if (item.type === 'system' && Number.isSafeInteger(item.ID) && item.ID > 0) candidates.push(item.ID);
        }
        candidates = [...new Set(candidates)];
        for (const widgetId of candidates) {
          if (validHint && hint.tabId === tabId && hint.widgetId === widgetId) continue;
          if (await readIdentityWidget({tabId, widgetId})) break;
        }
      }
    }
    if (codes.length !== 1) throw new Error('Request chi tiết chưa trả đúng một mã công việc.');
    const originalCode = task.code ?? '';
    const fraction = fractionalCode(originalCode);
    const equivalentFraction = fraction !== null && fraction === fractionalCode(codes[0]);
    if (codes[0] !== originalCode && !equivalentFraction) throw new Error(`Mã chi tiết "${codes[0]}" khác mã đối chiếu "${verifiedLookupCode}" (báo cáo: "${originalCode}"). Không khớp quy tắc dạng mã được duyệt.`);
    const identityWarning = codes[0] !== originalCode
      ? `Mã báo cáo "${originalCode}" khác mã chi tiết "${codes[0]}". Đã đối chiếu theo quy tắc mã dạng thập phân được duyệt (0 đầu và 0 cuối phần thập phân), khớp tên và xác minh ID; giữ mã báo cáo trong Excel, không sửa 1Office.`
      : undefined;
    return {...task, internalId: url.searchParams.get('ID')!, url: url.href, identityError: undefined, identityWidget, identityWarning};
  } catch (error) {
    return {...task, internalId: '', url: '', identityError: `Chưa lấy được thảo luận: ${error instanceof Error ? error.message : 'Không xác minh được ID qua request ngầm.'} Không mở trang công việc.`};
  }
}

export async function resolveReportTask(tabId: number, sourceUrl: string, task: OneOfficeProjectTaskInventoryItem): Promise<OneOfficeProjectTaskInventoryItem> {
  const allowedOrigin = new URL(sourceUrl).origin;
  await getOneOfficeTabOrigin(tabId, allowedOrigin);
  const api = (globalThis as any).chrome;
  const resolve = async (): Promise<OneOfficeProjectTaskInventoryItem> => {
    const widgetHints = Object.fromEntries(['work-task-task', 'work-normal-normal'].flatMap(object => {
      const hint = identityWidgets.get(`${allowedOrigin}/${object}`);
      return hint ? [[object, hint]] : [];
    }));
    const {identityWidget, ...result} = readReportInjectionResult<LookupResult>(await api.scripting.executeScript({target: {tabId},
      func: lookupReportTaskFromPage, args: [{sourceUrl, task, widgetHints}]}));
    if (identityWidget && !result.identityError) identityWidgets.set(`${allowedOrigin}/${new URL(result.url).pathname.split('/')[2]}`, identityWidget);
    return result;
  };
  // Share the first discovery instead of making every concurrent row scan all widgets.
  while (identityWarmups.has(allowedOrigin)) await identityWarmups.get(allowedOrigin);
  if (discoveryAttempted.has(allowedOrigin)) return resolve();
  const pending = resolve();
  const warmup = pending.then(() => {}, () => {});
  identityWarmups.set(allowedOrigin, warmup);
  try { return await pending; }
  finally {
    discoveryAttempted.add(allowedOrigin);
    if (identityWarmups.get(allowedOrigin) === warmup) identityWarmups.delete(allowedOrigin);
  }
}
