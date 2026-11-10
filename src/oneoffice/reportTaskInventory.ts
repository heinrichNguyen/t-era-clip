import type { OneOfficeProjectInventory, OneOfficeProjectTaskInventoryItem } from './types';
import { getOneOfficeTabOrigin } from './origins';

export interface ReportTaskRow {
  task: OneOfficeProjectTaskInventoryItem;
  explicitUrl: string | null;
}

// Self-contained: Chrome serializes this function into the source tab.
export function readOneOfficeReportFromPage(options: { allowedOrigin: string; sourceUrl: string; captureErrors?: boolean }): {
  title: string; rows: ReportTaskRow[]; error?: string;
} {
  try {
  const current = new URL(window.location.href);
  if (current.origin !== options.allowedOrigin || current.href !== options.sourceUrl
    || current.pathname !== '/apps/report-report/view' || !/^\d+$/.test(current.searchParams.get('ID') ?? '')) {
    throw new Error('Báo cáo hoặc bộ lọc đã thay đổi. Mở lại extension từ đúng báo cáo.');
  }
  const text = (value: string | null | undefined) => (value ?? '').replace(/\s+/g, ' ').trim();
  const tables = Array.from(document.querySelectorAll<HTMLTableElement>('table[_n="table"]'))
    .filter(table => table.getClientRects().length > 0);
  const candidates = tables.flatMap(table => Array.from(table.rows).flatMap(header => {
    const labels = Array.from(header.cells).map(cell => text(cell.getAttribute('ov') ?? cell.textContent));
    return labels.includes('Mã công việc') && labels.includes('Tên công việc') ? [{ table, header, labels }] : [];
  }));
  if (candidates.length !== 1) throw new Error('Chưa nhận diện duy nhất bảng công việc trong báo cáo.');
  const { table, header, labels } = candidates[0];
  if (new Set(labels).size !== labels.length) throw new Error('Báo cáo có tiêu đề cột trùng, chưa thể đối chiếu an toàn.');
  const rows: ReportTaskRow[] = [];
  for (const row of Array.from(table.rows).slice(header.rowIndex + 1)) {
    if (!text(row.textContent)) continue;
    if (row.cells.length !== labels.length || Array.from(row.cells).some(cell => cell.colSpan !== 1 || cell.rowSpan !== 1)) {
      throw new Error('Bảng báo cáo có dòng gộp hoặc cấu trúc khác dự kiến. Chưa thu thập.');
    }
    const read = (label: string) => text(row.cells[labels.indexOf(label)]?.textContent);
    const name = read('Tên công việc');
    const code = read('Mã công việc');
    if (!name) throw new Error(`Dòng ${row.rowIndex + 1} thiếu tên công việc. Chưa thể xác minh ID.`);
    const nameCell = row.cells[labels.indexOf('Tên công việc')];
    const taskLinks = Array.from(nameCell.querySelectorAll<HTMLAnchorElement>('a[href]')).flatMap(link => {
      const url = new URL(link.getAttribute('href')!, current.href);
      return url.origin === current.origin && /^\/apps\/work-(task-task|normal-normal)\/view$/.test(url.pathname)
        && /^\d+$/.test(url.searchParams.get('ID') ?? '') ? [url] : [];
    });
    const explicitId = read('ID công việc');
    if (explicitId && !/^[1-9]\d*$/.test(explicitId)) throw new Error('Cột ID công việc có giá trị không hợp lệ.');
    if (taskLinks.length > 1 || (taskLinks[0] && explicitId && taskLinks[0].searchParams.get('ID') !== explicitId)) {
      throw new Error('ID và liên kết công việc trong báo cáo không khớp.');
    }
    const explicitUrl = taskLinks[0]?.href ?? (explicitId ? `${current.origin}/apps/work-task-task/view?ID=${explicitId}` : null);
    rows.push({ explicitUrl, task: {
      sourceOrder: rows.length, internalId: '', no: read('No.'), code, name, url: '', depth: 0,
      collectionKey: `report-row:${rows.length}`,
      ...(!code && !explicitUrl ? { identityError: 'Chưa lấy được thảo luận: báo cáo không có mã hoặc liên kết/ID để xác định đúng công việc. Dòng công việc vẫn được giữ để xuất.' } : {}),
      projectName: read('Dự án'),
      parentName: read('Công việc cha'),
      parentNo: null, parentInternalId: null,
      assigneeNames: read('Người thực hiện').split(',').map(text).filter(Boolean),
      assignerNames: read('Người giao việc').split(',').map(text).filter(Boolean),
      departmentNames: read('Phòng ban thực hiện').split(',').map(text).filter(Boolean),
      taskStatus: read('Trạng thái'), baselineStart: read('Bắt đầu (KH gốc)'), baselineEnd: read('Kết thúc (KH gốc)'),
      plannedStart: read('Bắt đầu dự kiến'), plannedEnd: read('Kết thúc dự kiến'), actualStart: read('Bắt đầu thực tế'),
      actualEnd: read('Kết thúc thực tế'), overdueDays: read('Quá hạn (ngày)'),
      completedOverdueDays: read('Số ngày hoàn thành quá hạn'), labels: read('Nhãn').split(',').map(text).filter(Boolean),
    } });
  }
  if (!rows.length) throw new Error('Báo cáo chưa có công việc để thu thập.');
  return { title: text(document.querySelector('.main-content')?.querySelector('h1')?.textContent) || document.title.replace(/^\(\d+\)\s*/, ''), rows };
  } catch (error) {
    if (!options.captureErrors) throw error;
    return { title: '', rows: [], error: error instanceof Error ? error.message : 'Không đọc được bảng công việc trong báo cáo.' };
  }
}

export async function readOneOfficeReportIdentityFromPage(options: {
  allowedOrigin: string; expectedUrl: string; code: string; name: string; search: boolean; captureErrors?: boolean;
}): Promise<{ url: string; projectInternalId?: string; error?: string }> {
  try {
  const text = (value: string | null | undefined) => (value ?? '').replace(/\s+/g, ' ').trim();
  const start = Date.now();
  let previous = '';
  let stableSince = 0;
  while (Date.now() - start < 15000) {
    if (window.location.origin !== options.allowedOrigin || window.location.href !== options.expectedUrl) {
      throw new Error('Tab nguồn đã chuyển trang trong lúc xác minh công việc.');
    }
    let result = '';
    if (options.search) {
      const count = /\/\s*([\d,.]+)\s*bản ghi/i.exec(text(document.querySelector('.table-page-title')?.textContent));
      const links = Array.from(document.querySelectorAll<HTMLAnchorElement>('tr.table-row-body a.table-main-title[href]'));
      if (count && Number(count[1].replace(/[,.]/g, '')) === 1 && links.length === 1 && text(links[0].textContent) === options.name) {
        const url = new URL(links[0].getAttribute('href')!, options.allowedOrigin);
        if (url.origin === options.allowedOrigin && /^\/apps\/work-(normal-normal|task-task)\/view$/.test(url.pathname)
          && /^[1-9]\d*$/.test(url.searchParams.get('ID') ?? '')) result = url.href;
      }
    } else {
      const fields = Array.from(document.querySelectorAll('.detail-field')).filter(field =>
        text(field.querySelector('.detail-field-label')?.textContent) === 'Mã công việc');
      const url = new URL(options.expectedUrl);
      const id = url.searchParams.get('ID');
      const object = url.pathname.includes('work-normal-normal') ? 'work-normal-normal' : 'work-task-task';
      const hasDiscussion = document.querySelector(`.cards.comments[comments-id="${object}@${id}"]`);
      const heading = text(document.querySelector('.header-right')?.textContent);
      if (fields.length === 1 && text(fields[0].querySelector('.detail-field-content')?.textContent) === options.code
        && heading.includes(options.name) && hasDiscussion) result = url.href;
    }
    if (result) {
      if (previous !== result) { previous = result; stableSince = Date.now(); }
      else if (Date.now() - stableSince >= 400) {
        const projectField = Array.from(document.querySelectorAll('.detail-field')).find(field =>
          text(field.querySelector('.detail-field-label')?.textContent) === 'Dự án');
        const href = projectField?.querySelector('a[href]')?.getAttribute('href');
        const project = href ? new URL(href, options.allowedOrigin) : null;
        const projectInternalId = project?.origin === options.allowedOrigin && project.pathname === '/apps/work-project-project/view'
          && /^[1-9]\d*$/.test(project.searchParams.get('ID') ?? '') ? project.searchParams.get('ID')! : undefined;
        return { url: result, ...(!options.search && projectInternalId ? { projectInternalId } : {}) };
      }
    } else { previous = ''; stableSince = 0; }
    await new Promise(resolve => setTimeout(resolve, 100));
  }
  throw new Error(`Không xác minh duy nhất mã và tên: ${options.code}. Kiểm tra quyền xem, mã trùng hoặc bổ sung liên kết/ID công việc vào báo cáo.`);
  } catch (error) {
    if (!options.captureErrors) throw error;
    return { url: '', error: error instanceof Error ? error.message : 'Không xác minh được ID công việc.' };
  }
}

export function readReportInjectionResult<T>(results: Array<{result?: T | null; error?: {message?: string}}> | null | undefined): T {
  const entry = results?.[0];
  if (entry?.error) throw new Error(entry.error.message || 'Không chạy được bước đọc báo cáo.');
  if (!entry || entry.result == null) throw new Error('Bước đọc báo cáo không trả về dữ liệu. Tải lại báo cáo, chờ bảng hiển thị xong và mở lại tiện ích. Chưa thu thập thảo luận.');
  const failure = (entry.result as {error?: unknown}).error;
  if (typeof failure === 'string' && failure) throw new Error(failure);
  return entry.result;
}

export async function collectOneOfficeReportInventory(tabId: number, expectedOrigin: string, sourceUrl: string): Promise<OneOfficeProjectInventory> {
  const api = (globalThis as any).chrome;
  await getOneOfficeTabOrigin(tabId, expectedOrigin);
  const snapshot = readReportInjectionResult<ReturnType<typeof readOneOfficeReportFromPage>>(
    await api.scripting.executeScript({ target: { tabId }, func: readOneOfficeReportFromPage,
      args: [{ allowedOrigin: expectedOrigin, sourceUrl, captureErrors: true }] }));
  if (!Array.isArray(snapshot.rows) || !snapshot.rows.length) throw new Error('Báo cáo chưa trả về danh sách công việc hợp lệ. Chưa thu thập.');
  const capturedAt = new Date().toISOString();
  // The popup only freezes the filtered rows. Network work starts on the review page.
  const tasks = snapshot.rows.map(row => ({ ...row.task, ...(row.explicitUrl ? {reportLookupUrl: row.explicitUrl} : {}) }));
  return { schemaVersion: 1, source: '1office-report-dom', capturedAt,
    project: { type: 'report', internalId: new URL(sourceUrl).searchParams.get('ID')!, code: null,
      name: `Danh sách báo cáo: ${snapshot.title}`, url: sourceUrl },
    collection: { complete: true, expectedTaskCount: tasks.length, loadedTaskCount: tasks.length,
      warnings: ['Chỉ thu thập danh sách trong báo cáo tại thời điểm kiểm tra, không mở rộng toàn bộ dự án. Ngày báo cáo không tự giới hạn ngày bình luận.',
        ...(tasks.some(task => task.identityError) ? [`${tasks.filter(task => task.identityError).length} công việc chưa xác minh được ID: vẫn giữ đầy đủ trong sheet Công việc, chưa lấy được thảo luận. Không coi là không có bình luận.`] : [])] }, tasks };
}
