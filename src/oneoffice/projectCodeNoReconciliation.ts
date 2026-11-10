import type { ImportReview } from './projectImportReview';
import type { OneOfficeProjectInventory } from './types';
import type { TestTaskDetails } from './testTaskDetails';
import { TEST_BASELINE_ORIGIN, TEST_BASELINE_PROJECT } from './testBaselineScope';
import { oneOfficeOrigin } from './origins';

export interface CodeNoOptions {
  idColumn: string;
  oldCodeColumn: string;
  codeColumn: string;
  noColumn: string;
  mappings: string;
  approvedFile: boolean;
  confirmedProject: boolean;
  noMeaning: '' | 'custom-no' | 'outline';
}
export interface CodeNoRow {
  row: number; name: string; internalId: string | null; match: 'id' | 'old-code' | 'target-code' | 'unresolved';
  beforeCode: string | null; beforeNo: string | null; afterCode: string; afterNo: string;
  state: 'changed' | 'same' | 'review'; issues: string[];
}
export interface CodeNoOperation { internalId: string; field: 'code' | 'sourceNo'; before: string; after: string; temporary: boolean }
export interface CodeNoPreview {
  mode: 'code-no-preview'; writeEnabled: false; snapshotAt: string | null;
  rows: CodeNoRow[]; errors: string[]; unchangedOutsideFile: number; operations: CodeNoOperation[];
}
const cell = (value: unknown) => typeof value === 'string' || typeof value === 'number' ? String(value).trim() : '';

// Identity and target values are independent. This function never creates, deletes or writes a task.
export function reconcileProjectCodeNo(report: ImportReview, inventory: OneOfficeProjectInventory | null, options: CodeNoOptions,
  details: Record<string, TestTaskDetails> = {}): CodeNoPreview {
  const errors: string[] = [];
  if (!inventory?.collection.complete || inventory.collection.expectedTaskCount !== inventory.tasks.length || inventory.collection.loadedTaskCount !== inventory.tasks.length) errors.push('Cần danh sách dự án đầy đủ để đối chiếu.');
  if (!options.approvedFile) errors.push('Cần xác nhận đây là file được phê duyệt làm nguồn mã và No. đích.');
  if (!options.confirmedProject) errors.push('Cần xác nhận đúng dự án đang đối chiếu.');
  if (options.noMeaning !== 'custom-no') errors.push(options.noMeaning === 'outline' ? 'Số thứ tự trên cây/Gantt chưa được hỗ trợ cập nhật ở luồng này. Không thay bằng trường No. tùy chỉnh.' : 'Cần xác nhận No. đích là trường No. tùy chỉnh, không phải số thứ tự sinh trên cây/Gantt.');
  if (!options.codeColumn || !options.noColumn || options.codeColumn === options.noColumn) errors.push('Chọn hai cột khác nhau cho mã đầy đủ và No. được duyệt.');
  if (inventory) {
    try { const project = new URL(inventory.project.url); oneOfficeOrigin(project.href);
      if (!/^\d+$/.test(inventory.project.internalId) || project.pathname !== '/apps/work-project-project/view' || project.searchParams.getAll('ID').length !== 1 || project.searchParams.get('ID') !== inventory.project.internalId) errors.push('Định danh dự án không khớp URL nguồn.');
    } catch { errors.push('Nguồn dự án không thuộc môi trường 1Office được hỗ trợ.'); }
  }
  const columns = new Set(report.rows.flatMap((row) => Object.keys(row.values)));
  if ([options.idColumn, options.oldCodeColumn, options.codeColumn, options.noColumn].some((column) => column && !columns.has(column))) errors.push('Cột đã chọn không có trong sheet đang dùng.');
  const ids = new Map<string, NonNullable<typeof inventory>['tasks'][number]>();
  const codes = new Map<string, string[]>();
  const current = new Map<string, { code: string | null; no: string | null }>();
  const sourceRows = new Set(report.rows.map((row) => row.row));
  if (sourceRows.size !== report.rows.length) errors.push('Dòng Excel bị lặp trong bản nguồn.');
  for (const task of inventory?.tasks ?? []) {
    if (!/^\d+$/.test(task.internalId) || ids.has(task.internalId)) errors.push('Danh sách có ID không hợp lệ hoặc bị trùng.');
    try {
      const url = new URL(task.url); const project = new URL(inventory!.project.url); oneOfficeOrigin(url.href);
      if (url.origin !== project.origin || !['/apps/work-task-task/view', '/apps/work-task-task/quickview'].includes(url.pathname) || url.searchParams.getAll('ID').length !== 1 || url.searchParams.get('ID') !== task.internalId) errors.push('URL công việc không khớp ID hoặc khác môi trường dự án.');
    } catch { errors.push('URL dự án hoặc công việc không hợp lệ.'); }
    ids.set(task.internalId, task);
    let code = task.code === null || task.code === undefined ? null : task.code;
    let no = typeof task.sourceNo === 'string' ? task.sourceNo : null;
    const evidence = details[task.internalId];
    // Only the verified project-203 reader can supply detail evidence in this workflow.
    if (evidence) {
      let valid = false;
      try { const url = new URL(evidence.sourceUrl); valid = new URL(inventory!.project.url).origin === TEST_BASELINE_ORIGIN && inventory!.project.internalId === TEST_BASELINE_PROJECT &&
        evidence.identityVerified && evidence.taskId === task.internalId && evidence.evidenceSource === 'widget-response' && url.origin === TEST_BASELINE_ORIGIN && url.pathname === '/apps/work-task-task/view' && url.searchParams.get('ID') === task.internalId; } catch { /* Invalid evidence stays unknown. */ }
      if (!valid) errors.push(`ID ${task.internalId}: nguồn chi tiết không đúng phạm vi xác minh.`);
      else {
        if (evidence.code.state === 'value' || evidence.code.state === 'empty') code = evidence.code.displayValue;
        if (evidence.no.state === 'value' || evidence.no.state === 'empty') no = evidence.no.displayValue;
      }
    }
    current.set(task.internalId, { code, no });
    if (code) codes.set(code, [...(codes.get(code) ?? []), task.internalId]);
  }
  const mapping = new Map<number, string>();
  if (options.mappings.length > 1_000_000) errors.push('Bảng ánh xạ ID quá dài.');
  else for (const line of options.mappings.split(/\r?\n/).filter((line) => line.trim())) {
    const fields = line.trim().split(/\s+/); const row = Number(fields[0]); const id = fields[1];
    if (fields.length !== 2 || !/^\d+$/.test(fields[0]) || !/^\d+$/.test(id ?? '') || !sourceRows.has(row) || !ids.has(id)) { errors.push('Ánh xạ cần hai cột: dòng Excel và ID hiện hữu trong đúng dự án.'); continue; }
    if (mapping.has(row)) errors.push(`Dòng ${row}: có nhiều ánh xạ ID.`);
    mapping.set(row, id);
  }
  const rows: CodeNoRow[] = report.rows.map((source) => {
    const issues: string[] = [];
    const afterCode = cell(source.values[options.codeColumn]); const afterNo = cell(source.values[options.noColumn]);
    const columnId = options.idColumn ? cell(source.values[options.idColumn]) : '';
    const manualId = mapping.get(source.row);
    const oldCode = options.oldCodeColumn ? cell(source.values[options.oldCodeColumn]) : '';
    const result: CodeNoRow = { row: source.row, name: source.name, internalId: null, match: 'unresolved', beforeCode: null, beforeNo: null, afterCode, afterNo, state: 'review', issues };
    if (afterCode.length === 0 || afterCode.length > 160 || /[\s\x00-\x1f\x7f]/.test(afterCode) || /^\d+(?:\.\d+)*$/.test(afterCode)) issues.push('Thiếu mã đầy đủ hoặc mã đích không hợp lệ (không dùng mã số cục bộ). Không tự sinh lại mã đã duyệt.');
    if (!/^[1-9]\d*(?:\.[1-9]\d*)*$/.test(afterNo)) issues.push('No. đích cần là số phân cấp như 1.2.3, không có hậu tố chữ hoặc ô trống.');
    if (typeof source.values[options.noColumn] === 'number' && !Number.isInteger(source.values[options.noColumn])) issues.push('No. phân cấp phải lưu dạng văn bản trong Excel để không mất phần số như 1.10 thành 1.1.');
    if (options.idColumn && typeof source.values[options.idColumn] === 'number' && !Number.isSafeInteger(source.values[options.idColumn])) issues.push('ID dạng số trong Excel không còn chính xác, cần lưu ID dạng văn bản.');
    if (manualId && columnId && manualId !== columnId) issues.push('ID nhập tay khác ID trong file.');
    const requestedId = manualId || columnId;
    if (requestedId) {
      if (!/^\d+$/.test(requestedId) || !ids.has(requestedId)) issues.push('ID không thuộc dự án hiện tại.');
      else { result.internalId = requestedId; result.match = 'id'; }
    } else {
      const candidates = codes.get(oldCode || afterCode) ?? [];
      if (candidates.length === 1) { result.internalId = candidates[0]; result.match = oldCode ? 'old-code' : 'target-code'; }
      else issues.push(candidates.length > 1 ? 'Mã hiện tại bị trùng, cần xác nhận ID.' : 'Chưa xác định ID. Không tự coi đây là việc mới; tạo/import trước rồi ghép ID, hoặc xác nhận lại công việc cũ.');
    }
    if (result.internalId) {
      const before = current.get(result.internalId)!; result.beforeCode = before.code; result.beforeNo = before.no;
      if (oldCode && before.code !== oldCode) issues.push('Mã cũ trong file không khớp hiện trạng của ID đã chọn.');
      if (before.code === null) issues.push('Chưa đọc được mã hiện tại; không coi là mã trống.');
      if (before.no === null) issues.push('Chưa đọc được trường No. hiện tại; không dùng số thứ tự trên cây thay thế.');
    }
    return result;
  });
  for (const key of ['internalId', 'afterCode', 'afterNo'] as const) {
    const grouped = new Map<string, CodeNoRow[]>();
    for (const row of rows) if (row[key]) grouped.set(row[key]!, [...(grouped.get(row[key]!) ?? []), row]);
    for (const group of grouped.values()) if (group.length > 1) for (const row of group) row.issues.push(key === 'internalId' ? 'Một ID bị ghép vào nhiều dòng file.' : key === 'afterCode' ? 'Mã đích trùng giữa các dòng file.' : 'No. đích trùng giữa các dòng file.');
  }
  const changedIds = new Set(rows.flatMap((row) => row.internalId ? [row.internalId] : []));
  for (const [id, before] of current) if (!changedIds.has(id)) {
    if (before.code === null || before.no === null) errors.push('Chưa đủ mã/No. của các công việc ngoài file để kiểm tra trùng toàn dự án.');
    for (const row of rows) {
      if (before.code && before.code === row.afterCode) row.issues.push(`Mã đích đang thuộc ID ${id} ngoài file, không được thay đổi công việc đó.`);
      if (before.no && before.no === row.afterNo) row.issues.push(`No. đích đang thuộc ID ${id} ngoài file, không được thay đổi công việc đó.`);
    }
  }
  for (const row of rows) if (row.internalId && !row.issues.length) row.state = row.beforeCode === row.afterCode && row.beforeNo === row.afterNo ? 'same' : 'changed';
  const operations: CodeNoOperation[] = [];
  if (!errors.length && rows.length && rows.every((row) => row.state !== 'review')) {
    for (const field of ['code', 'sourceNo'] as const) {
      const beforeKey = field === 'code' ? 'beforeCode' : 'beforeNo'; const afterKey = field === 'code' ? 'afterCode' : 'afterNo';
      const positions = new Map([...current].map(([id, value]) => [id, field === 'code' ? value.code : value.no]));
      const owners = new Map<string, Set<string>>();
      for (const [id, value] of positions) if (value) { const set = owners.get(value) ?? new Set<string>(); set.add(id); owners.set(value, set); }
      const reserved = new Set([...positions.values(), ...rows.map((row) => row[afterKey])]);
      const pending = rows.filter((row) => row[beforeKey] !== row[afterKey]).map((row) => ({ id: row.internalId!, target: row[afterKey] }));
      let sequence = 0;
      while (pending.length) {
        const free = pending.findIndex((item) => { const set = owners.get(item.target); return !set?.size || set.size === 1 && set.has(item.id); });
        const index = free >= 0 ? free : 0; const item = pending[index];
        let target = item.target;
        if (free < 0) { do { sequence++; target = field === 'code' ? `TERA_TMP_${sequence}` : `990000.${sequence}`; } while (reserved.has(target)); reserved.add(target); }
        operations.push({ internalId: item.id, field, before: positions.get(item.id)!, after: target, temporary: free < 0 });
        const old = positions.get(item.id); if (old) owners.get(old)?.delete(item.id);
        const holders = owners.get(target) ?? new Set<string>(); holders.add(item.id); owners.set(target, holders);
        positions.set(item.id, target);
        if (free >= 0) pending.splice(index, 1);
      }
    }
  }
  return { mode: 'code-no-preview', writeEnabled: false, snapshotAt: inventory?.capturedAt ?? null, rows, errors: [...new Set(errors)], unchangedOutsideFile: (inventory?.tasks.length ?? 0) - changedIds.size, operations };
}
