import type { OneOfficeProjectInventory } from './types';

export type ImportCell = string | number | boolean | Date | null;
export interface ImportSheet { sheet: string; data: ImportCell[][] }
export interface ImportReviewRow {
  row: number;
  no: string;
  originalWbs: string;
  code: string;
  name: string;
  level: number | null;
  parentCandidate: string;
  values: Record<string, ImportCell>;
  issues: string[];
  match: 'blocked' | 'matched' | 'unmatched' | 'unavailable';
  internalId: string | null;
}
export interface ImportReview {
  sheet: string;
  rows: ImportReviewRow[];
  levels: Record<number, number>;
  warnings: string[];
}
const clean = (v: ImportCell | undefined) => v == null ? '' : String(v).trim();
const REQUIRED = ['No.', 'Mã gói thầu/Mã CV', 'Nội dung công việc'];
const LIMIT = 10000;

// This is a file outline proposal, never a mutation of 1Office parent IDs.
export function reviewProjectImport(sheet: ImportSheet, inventory: OneOfficeProjectInventory | null): ImportReview {
  const header = (sheet.data[0] || []).map(clean);
  if (REQUIRED.some((name) => !header.includes(name))) throw new Error('Thiếu cột No., Mã gói thầu/Mã CV hoặc Nội dung công việc ở dòng đầu.');
  const nonblankHeaders = header.filter(Boolean);
  if (new Set(nonblankHeaders).size !== nonblankHeaders.length) throw new Error('Tiêu đề cột bị trùng, cần kiểm tra file.');
  if (sheet.data.length > LIMIT + 1 || header.length > 100) throw new Error('Chỉ hỗ trợ tối đa 10.000 dòng và 100 cột mỗi sheet.');
  const at = (cells: ImportCell[], name: string) => cells[header.indexOf(name)] ?? null;
  const rows: ImportReviewRow[] = [];
  for (let i = 1; i < sheet.data.length; i++) {
    const cells = sheet.data[i];
    if (cells.every((v) => v == null || v === '')) continue;
    const no = clean(at(cells, 'No.'));
    const outlineValid = /^\d+[a-z]?(?:\.\d+[a-z]?)*$/i.test(no);
    const code = clean(at(cells, 'Mã gói thầu/Mã CV'));
    const name = clean(at(cells, 'Nội dung công việc'));
    const issues: string[] = [];
    if (!name) issues.push('Thiếu tên công việc');
    if (!outlineValid) issues.push('No. chưa nhận diện được cấu trúc');
    else if (/[a-z]/i.test(no)) issues.push('No. có hậu tố chữ, cần duyệt đánh lại số');
    if (!code) issues.push('Thiếu mã công việc');
    else if (/^\d+(?:\.\d+)*$/.test(code)) issues.push('Mã dạng số cục bộ, cần xác nhận mã ổn định');
    const values: Record<string, ImportCell> = Object.create(null);
    header.forEach((key, j) => { if (key) values[key] = cells[j] ?? null; });
    rows.push({ row: i + 1, no, originalWbs: clean(at(cells, 'Text5 (WBS gốc)')), code, name,
      level: outlineValid ? no.split('.').length : null,
      parentCandidate: outlineValid && no.includes('.') ? no.slice(0, no.lastIndexOf('.')) : '',
      values, issues, match: 'unavailable', internalId: null });
  }
  const group = <T,>(items: T[], key: (item: T) => string) => {
    const result = new Map<string, T[]>();
    items.forEach((item) => { const k = key(item); const values = result.get(k); if (values) values.push(item); else result.set(k, [item]); });
    return result;
  };
  const byNo = group(rows, (r) => r.no);
  const byCode = group(rows, (r) => r.code);
  const sourceCodes = group(inventory?.tasks || [], (r) => r.code || '');
  const levels: Record<number, number> = {};
  for (const row of rows) {
    if (row.level) levels[row.level] = (levels[row.level] || 0) + 1;
    if ((byNo.get(row.no)?.length || 0) > 1) row.issues.push('No. bị trùng trong file');
    if (row.code && (byCode.get(row.code)?.length || 0) > 1) row.issues.push('Mã bị trùng trong file');
    if (row.parentCandidate) {
      const parent = byNo.get(row.parentCandidate);
      if (!parent || parent.length !== 1) row.issues.push('Chưa xác định duy nhất dòng cha theo No.');
      else if (parent[0].row > row.row) row.issues.push('Dòng cha theo No. nằm sau công việc, cần kiểm tra cấp');
    }
    const candidates = sourceCodes.get(row.code) || [];
    if (row.code && candidates.length > 1) row.issues.push('Mã bị trùng trong danh sách 1Office');
    if (row.issues.length) row.match = 'blocked';
    else if (!inventory?.collection.complete) row.match = 'unavailable';
    else if (candidates.length === 1) { row.match = 'matched'; row.internalId = candidates[0].internalId; }
    else row.match = 'unmatched';
  }
  const warnings = ['Chỉ kiểm tra và đề xuất đối chiếu mã. Chưa xác nhận tương đương các trường, chưa tạo công việc hoặc lập lệnh ghi.',
    'Cấp và dòng cha là gợi ý từ No. trong file, không xác nhận hoặc thay đổi quan hệ cha/con trên 1Office.',
    'Không dùng tên hoặc No. để tự ghép công việc. Mã không khớp chưa đủ chứng minh là công việc mới.'];
  if (!inventory?.collection.complete) warnings.push('Chưa có danh sách 1Office đầy đủ, chưa đối chiếu mã với hệ thống.');
  if (header.includes('Bắt đầu thực tế')) warnings.push('Bắt đầu thực tế: cần xác nhận nguồn ngày. Bộ đọc lấy giá trị có sẵn trong Excel, không xác minh công thức và không tính lại công thức.');
  if (header.includes('MQH')) warnings.push('MQH được giữ để tham khảo, chưa quy đổi hoặc ghi liên kết phụ thuộc lên 1Office.');
  return { sheet: sheet.sheet, rows, levels, warnings };
}
