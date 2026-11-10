import type { ImportCell, ImportReview } from './projectImportReview';
import type { NormalizedImportRow, proposeImportNormalization } from './projectImportNormalization';
import type { OneOfficeProjectInventory, OneOfficeProjectTaskInventoryItem } from './types';

export const IMPORT_DIFF_FIELDS = [
  { id: 'name', label: 'Tên công việc', suggestedColumn: 'Nội dung công việc', date: false },
  { id: 'sourceNo', label: 'No. (không đổi mã hoặc cha/con)', suggestedColumn: 'No.', date: false },
  { id: 'plannedStart', label: 'Bắt đầu kế hoạch', suggestedColumn: 'Bắt đầu', date: true },
  { id: 'plannedEnd', label: 'Kết thúc kế hoạch', suggestedColumn: 'Kết thúc', date: true },
  { id: 'baselineStart', label: 'Bắt đầu kế hoạch gốc', suggestedColumn: 'Bắt đầu (KH gốc)', date: true },
  { id: 'baselineEnd', label: 'Kết thúc kế hoạch gốc', suggestedColumn: 'Kết thúc (KH gốc)', date: true },
] as const;
export type ImportDiffField = typeof IMPORT_DIFF_FIELDS[number]['id'];
export interface ImportDiffOptions {
  columns: Partial<Record<ImportDiffField, string>>;
  confirmPlanDates: boolean;
  confirmProject: boolean;
}
export interface ImportFieldDiff {
  field: ImportDiffField; column: string; before: string; beforeKnown: boolean; after: string;
  state: 'changed' | 'same' | 'keep' | 'unknown' | 'invalid' | 'draft'; note: string;
}
export interface ImportRowDiff {
  row: number; code: string; name: string; internalId: string | null;
  state: 'changed' | 'same' | 'review' | 'draft' | 'blocked' | 'excluded';
  issues: string[]; fields: ImportFieldDiff[];
}

const blank = (value: ImportCell | undefined) => value == null || (typeof value === 'string' && ['', 'NA', 'N/A'].includes(value.trim().toUpperCase()));
const display = (value: ImportCell | undefined) => value instanceof Date ? (Number.isFinite(value.getTime()) ? value.toISOString() : '(ngày không hợp lệ)') : value == null ? '' : String(value).trim();

// Dates are compared as calendar dates only. Reject time-bearing or ambiguous data.
export function importCalendarDate(value: ImportCell | undefined): string | null {
  let year: number, month: number, day: number;
  if (value instanceof Date) {
    if (!Number.isFinite(value.getTime()) || value.getUTCHours() || value.getUTCMinutes() || value.getUTCSeconds() || value.getUTCMilliseconds()) return null;
    year = value.getUTCFullYear(); month = value.getUTCMonth() + 1; day = value.getUTCDate();
  } else {
    if (typeof value !== 'string') return null;
    const raw = value.trim();
    const iso = /^(\d{4})-(\d{2})-(\d{2})(?:[T ]00:00:00(?:\.000)?Z?)?$/.exec(raw);
    const local = /^(\d{1,2})\/(\d{1,2})\/(\d{4})(?: 00:00:00)?$/.exec(raw);
    if (iso) [, year, month, day] = iso.map(Number);
    else if (local) { day = Number(local[1]); month = Number(local[2]); year = Number(local[3]); }
    else return null;
  }
  if (year < 1900 || year > 9999) return null;
  const check = new Date(Date.UTC(year, month - 1, day));
  if (check.getUTCFullYear() !== year || check.getUTCMonth() + 1 !== month || check.getUTCDate() !== day) return null;
  return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
}

export function buildImportDiff(report: ImportReview, proposal: ReturnType<typeof proposeImportNormalization>, inventory: OneOfficeProjectInventory | null, options: ImportDiffOptions) {
  const errors = [...proposal.errors];
  const allowed = new Set<string>(IMPORT_DIFF_FIELDS.map((f) => f.id));
  for (const key of Object.keys(options.columns)) {
    if (!allowed.has(key)) errors.push('Có trường không được phép so sánh trong bước này.');
    else if (!options.columns[key as ImportDiffField]) errors.push('Chọn cột Excel cho từng trường đã bật.');
  }
  const selected = IMPORT_DIFF_FIELDS.filter((f) => Boolean(options.columns[f.id]));
  if (!selected.length) errors.push('Chọn ít nhất một trường và cột Excel tương ứng.');
  const uses = new Map<string, string>();
  for (const field of selected) {
    const column = options.columns[field.id]!;
    if (uses.has(column)) errors.push(`Cột “${column}” đang được ghép cho nhiều trường. Cần xác nhận lại ý nghĩa dữ liệu.`);
    uses.set(column, field.id);
    if (field.date && /thực tế|actual|MQH|tiến độ|trạng thái/i.test(column)) errors.push(`Không dùng cột “${column}” để cập nhật ngày kế hoạch.`);
  }
  if (selected.some((f) => f.date) && !options.confirmPlanDates) errors.push('Cần xác nhận các cột ngày đã chọn là ngày kế hoạch, không phải ngày thực tế.');
  const hasExisting = proposal.rows.some((r) => r.kind === 'existing');
  if (hasExisting && (!inventory?.collection.complete || !options.confirmProject)) errors.push('Cần danh sách đầy đủ và xác nhận đúng dự án 1Office để so sánh công việc cũ.');
  const sourceRows = new Map(report.rows.map((r) => [r.row, r]));
  const idMap = new Map<string, OneOfficeProjectTaskInventoryItem[]>();
  const codeMap = new Map<string, OneOfficeProjectTaskInventoryItem[]>();
  const noMap = new Map<string, OneOfficeProjectTaskInventoryItem[]>();
  for (const task of inventory?.tasks || []) {
    const ids = idMap.get(task.internalId) || []; ids.push(task); idMap.set(task.internalId, ids);
    if (task.code) { const codes = codeMap.get(task.code) || []; codes.push(task); codeMap.set(task.code, codes); }
    if (task.sourceNo) { const numbers = noMap.get(task.sourceNo) || []; numbers.push(task); noMap.set(task.sourceNo, numbers); }
  }
  const proposalByCode = new Map<string, NormalizedImportRow[]>();
  for (const row of proposal.rows) if (row.code) { const items = proposalByCode.get(row.code) || []; items.push(row); proposalByCode.set(row.code, items); }
  const rows: ImportRowDiff[] = proposal.rows.map((row) => {
    const output: ImportRowDiff = { row: row.row, code: row.code, name: row.name, internalId: row.internalId, state: 'blocked', issues: [...row.issues], fields: [] };
    if (row.kind === 'excluded') { output.state = 'excluded'; return output; }
    if (row.kind === 'blocked' || errors.length) { if (errors.length) output.issues.push('Cần sửa thiết lập trước khi so sánh.'); return output; }
    const source = sourceRows.get(row.row);
    if (!source || (proposalByCode.get(row.code)?.length || 0) !== 1) { output.issues.push('Không xác định duy nhất dòng nguồn và mã đề xuất.'); return output; }
    let task: OneOfficeProjectTaskInventoryItem | undefined;
    if (row.kind === 'existing') {
      const candidates = idMap.get(row.internalId || '') || [];
      if (candidates.length !== 1 || candidates[0].code !== row.code || codeMap.get(row.code)?.length !== 1) { output.issues.push('ID và mã không còn khớp duy nhất danh sách 1Office.'); return output; }
      task = candidates[0];
      if (row.parentCode) {
        const parent = proposalByCode.get(row.parentCode);
        if (!parent || parent.length !== 1 || parent[0].kind !== 'existing' || !parent[0].internalId || task.parentInternalId !== parent[0].internalId) output.issues.push('Cha trong file khác hoặc chưa đối chiếu được với cha trên 1Office. Không tự chuyển công việc.');
      } else if (task.parentInternalId) output.issues.push('Cha trên 1Office nằm ngoài phạm vi mã của file, chưa xác nhận quan hệ cha/con.');
      else if (task.parentInternalId !== null) output.issues.push('Chưa có dữ liệu xác nhận cha trên 1Office.');
    } else if (codeMap.has(row.code)) { output.issues.push('Mã bản nháp đã tồn tại trên 1Office.'); return output; }
    for (const field of selected) {
      const column = options.columns[field.id]!;
      const value = source.values[column];
      const before = task?.[field.id];
      const delta: ImportFieldDiff = { field: field.id, column, before: task ? display(before) : '(bản nháp, chưa có giá trị đối chiếu)', beforeKnown: Boolean(task && before !== undefined && (!field.date || !blank(before))), after: display(value), state: 'unknown', note: '' };
      output.fields.push(delta);
      if (!Object.prototype.hasOwnProperty.call(source.values, column)) { delta.note = 'Không có cột đã chọn trong dòng nguồn.'; continue; }
      if (blank(value)) { delta.state = 'keep'; delta.note = 'Ô nguồn trống hoặc NA: giữ nguyên, không đề xuất xóa.'; continue; }
      const afterValue = field.date ? importCalendarDate(value) : typeof value === 'string' || typeof value === 'number' ? display(value) : null;
      if (afterValue === null || (field.id === 'sourceNo' && !/^\d+(?:\.\d+)*$/.test(afterValue))) { delta.state = 'invalid'; delta.note = field.date ? 'Ngày không hợp lệ, không rõ định dạng hoặc có giờ khác 00:00:00.' : field.id === 'sourceNo' ? 'No. cần dạng số phân cấp như 1.2.3.' : 'Giá trị không phải tên công việc dạng văn bản.'; continue; }
      delta.after = afterValue;
      if (!task) { delta.state = 'draft'; delta.note = 'Chỉ là dữ liệu bản nháp, chưa tạo công việc.'; continue; }
      if (before === undefined || (field.date && blank(before))) { delta.note = 'Danh sách hiện có chưa xác nhận giá trị cũ. Không tự coi là ô trống.'; continue; }
      const beforeValue = field.date ? importCalendarDate(before) : display(before);
      if (beforeValue === null) { delta.note = 'Không đọc được ngày cũ để so sánh chắc chắn.'; continue; }
      delta.before = beforeValue;
      delta.state = beforeValue === afterValue ? 'same' : 'changed';
      if (field.id === 'sourceNo' && delta.state === 'changed' && noMap.get(afterValue)?.some((other) => other.internalId !== task.internalId)) output.issues.push('No. đích đang thuộc công việc khác. Cần kiểm tra phương án đổi số, chưa lập lệnh ghi.');
    }
    for (const [start, end] of [['plannedStart','plannedEnd'], ['baselineStart','baselineEnd']] as const) {
      if (!selected.some((field) => field.id === start || field.id === end)) continue;
      const effective = (key: ImportDiffField) => {
        const delta = output.fields.find((f) => f.field === key);
        if (delta?.state === 'invalid' || delta?.state === 'unknown') return null;
        if (delta && delta.state !== 'keep') return importCalendarDate(delta.after);
        return importCalendarDate(task?.[key]);
      };
      const first = effective(start); const last = effective(end);
      if (first && last && first > last) output.issues.push('Ngày bắt đầu sau ngày kết thúc trong phương án dự kiến.');
      if ((!first || !last) && output.fields.some((field) => (field.field === start || field.field === end) && (field.state === 'changed' || field.state === 'draft'))) output.issues.push('Chưa đủ hai mốc ngày để kiểm tra thứ tự bắt đầu và kết thúc.');
    }
    output.state = output.issues.length || output.fields.some((f) => f.state === 'unknown' || f.state === 'invalid') ? 'review' : row.kind === 'proposed' ? 'draft' : output.fields.some((f) => f.state === 'changed') ? 'changed' : 'same';
    return output;
  });
  return { rows, errors: [...new Set(errors)], snapshotAt: inventory?.capturedAt || null };
}
