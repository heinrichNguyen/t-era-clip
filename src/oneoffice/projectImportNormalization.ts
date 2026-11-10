import type { ImportReview } from './projectImportReview';
import type { OneOfficeProjectInventory } from './types';
import { indexErpPackages, type ProjectErpCatalog } from './projectErpCatalog';

export interface NormalizationOptions {
  mode: 'initial' | 'update';
  prefix: string;
  firstTaskLevel: number;
  packageLevel: number;
  mappings: string;
  newRows: string;
  erpPackages?: Record<number, string>;
}
export interface NormalizedImportRow {
  row: number; no: string; sourceCode: string; name: string;
  group: string; code: string; parentCode: string; internalId: string | null;
  kind: 'excluded' | 'blocked' | 'existing' | 'proposed';
  issues: string[];
  erpCode?: string;
}
const validCode = (code: string) => code.length > 0 && code.length <= 160 && !/[\s\x00-\x1f\x7f]/.test(code) && !/(^\.|\.$|\.\.)/.test(code);
const numeric = (code: string) => /^\d+(?:\.\d+)*$/.test(code);

export function proposeImportNormalization(report: ImportReview, inventory: OneOfficeProjectInventory | null, options: NormalizationOptions, catalog: ProjectErpCatalog | null = null) {
  const errors: string[] = [];
  const prefix = options.prefix.trim();
  const erpIndex = catalog ? indexErpPackages(catalog) : null;
  if (catalog?.projects.length && catalog.projects.filter((p) => p.code === prefix).length !== 1) errors.push('Mã dự án chưa khớp duy nhất Danh mục hạng mục ERP của sheet đã chọn.');
  if (!validCode(prefix) || numeric(prefix)) errors.push('Nhập tiền tố mã dự án hợp lệ, không chỉ gồm số hoặc chứa khoảng trắng.');
  if (!['initial', 'update'].includes(options.mode)) errors.push('Chế độ đề xuất không hợp lệ.');
  if (!Number.isInteger(options.firstTaskLevel) || !Number.isInteger(options.packageLevel) || options.firstTaskLevel < 1 || options.packageLevel <= options.firstTaskLevel || options.packageLevel > 9) errors.push('Cấp gói thầu phải lớn hơn cấp hạng mục, trong khoảng 1 đến 9.');
  if (options.mappings.length > 1000000 || options.newRows.length > 100000) errors.push('Danh sách thiết lập quá dài.');
  const byRow = new Map(report.rows.map((r) => [r.row, r]));
  for (const [key, code] of Object.entries(options.erpPackages || {})) {
    if (!/^\d+$/.test(key) || byRow.get(Number(key))?.level !== options.packageLevel || !erpIndex || erpIndex.get(code)?.length !== 1) errors.push('Lựa chọn mã ERP không hợp lệ hoặc không còn thuộc danh mục đang dùng.');
  }
  const mapping = new Map<number, string>();
  for (const line of options.mappings.trim().split(/\r?\n/).filter(Boolean)) {
    const fields = line.trim().split(/\s+/);
    const row = Number(fields[0]);
    if (fields.length !== 2 || !/^\d+$/.test(fields[0]) || !byRow.has(row) || !validCode(fields[1])) { errors.push('Ghép mã cần đúng hai cột: dòng Excel và mã đầy đủ hợp lệ.'); continue; }
    if (mapping.has(row)) errors.push(`Dòng ${row}: ghép mã nhiều lần.`);
    mapping.set(row, fields[1]);
  }
  const newRows = new Set<number>();
  for (const token of options.newRows.trim().split(/[\s,;]+/).filter(Boolean)) {
    if (!/^\d+$/.test(token) || !byRow.has(Number(token))) errors.push(`Dòng đánh dấu mới không hợp lệ: ${token.slice(0, 30)}.`);
    else newRows.add(Number(token));
  }
  const tasksByCode = new Map<string, NonNullable<typeof inventory>['tasks']>();
  for (const task of inventory?.tasks || []) {
    if (!task.code) continue;
    const items = tasksByCode.get(task.code) || []; items.push(task); tasksByCode.set(task.code, items);
  }
  const noCount = new Map<string, number>();
  report.rows.forEach((r) => noCount.set(r.no, (noCount.get(r.no) || 0) + 1));
  const rows: NormalizedImportRow[] = [];
  const byNo = new Map<string, NormalizedImportRow>();
  const reserved = new Set(tasksByCode.keys());
  // Reserve explicit codes before allocating child suffixes, including later rows.
  for (const r of report.rows) {
    if (mapping.has(r.row)) reserved.add(mapping.get(r.row)!);
    if (r.code === prefix || r.code.startsWith(prefix + '.')) reserved.add(r.code);
  }
  const counters = new Map<string, number>();
  for (const source of report.rows) {
    const row: NormalizedImportRow = { row: source.row, no: source.no, sourceCode: source.code, name: source.name,
      group: '', code: '', parentCode: '', internalId: null, kind: 'blocked', issues: [] };
    rows.push(row);
    if (source.level !== null && source.level < options.firstTaskLevel) {
      row.group = 'Ngoài phạm vi cấp công việc'; row.kind = 'excluded';
      if (mapping.has(row.row) || newRows.has(row.row)) { row.kind = 'blocked'; row.issues.push('Thiết lập mã cho dòng ngoài phạm vi, cần chọn lại cấp.'); }
      continue;
    }
    row.group = source.level === options.firstTaskLevel ? 'Hạng mục' : source.level !== null && source.level < options.packageLevel ? 'Công việc trung gian' : source.level === options.packageLevel ? 'Gói thầu / công việc cùng cấp' : source.level === options.packageLevel + 1 ? 'Segment' : 'Công việc con';
    if (source.level === null) row.issues.push('Không nhận diện được cấp từ No.');
    if (!source.name) row.issues.push('Thiếu tên công việc.');
    if (noCount.get(source.no) !== 1) row.issues.push('No. trùng trong file.');
    if (errors.length) { row.issues.push('Cần sửa thiết lập trước khi đề xuất.'); continue; }
    if (source.level === null) continue;
    if (source.level > options.firstTaskLevel) {
      const parent = byNo.get(source.parentCandidate);
      if (!parent || parent.kind === 'blocked' || !parent.code) row.issues.push('Chưa có mã cha hợp lệ ở dòng trước.');
      else row.parentCode = parent.code;
    }
    const sourceMatches = tasksByCode.get(source.code) || [];
    if (options.mode === 'initial' && sourceMatches.length) row.issues.push('Mã nguồn đã có trên 1Office, cần dùng chế độ cập nhật thay vì cấp lại mã.');
    const explicitNew = newRows.has(source.row);
    if (options.mode === 'update' && !inventory?.collection.complete) row.issues.push('Cần danh sách 1Office đầy đủ để bảo toàn mã cũ.');
    const mapped = mapping.get(source.row);
    const allocating = options.mode === 'initial' || explicitNew;
    // Treat each supplied component as opaque: dots inside an ERP code are not levels.
    const baseCode = source.level === options.firstTaskLevel ? prefix : row.parentCode;
    const selectedErp = options.erpPackages?.[source.row];
    if (erpIndex && source.level === options.packageLevel) {
      const full = mapped || source.code;
      const component = erpIndex.has(full) ? full : baseCode && full.startsWith(baseCode + '.') ? full.slice(baseCode.length + 1) : full;
      row.erpCode = selectedErp || component;
      if (erpIndex.get(row.erpCode)?.length !== 1) row.issues.push('Mã gói thầu chưa khớp duy nhất danh mục ERP. Tìm và chọn mã, không tự ghép theo tên.');
      if (selectedErp && !allocating) row.issues.push('Không dùng lựa chọn ERP để thay mã công việc cũ. Giữ mã/ID 1Office và kiểm tra đối chiếu riêng.');
      if (mapped && selectedErp && mapped !== `${baseCode}.${selectedErp}`) row.issues.push('Mã đầy đủ nhập thủ công khác mã ERP đã chọn. Cần bỏ hoặc sửa một trong hai thiết lập.');
    }
    if (mapped) row.code = mapped;
    else if (options.mode === 'update' && sourceMatches.length === 1 && !explicitNew) row.code = source.code;
    else if (allocating && selectedErp && erpIndex?.get(selectedErp)?.length === 1 && baseCode) row.code = `${baseCode}.${selectedErp}`;
    else if (options.mode === 'initial' || explicitNew) {
      if (source.code === prefix || source.code.startsWith(prefix + '.')) row.code = source.code;
      else if (source.level <= options.packageLevel) {
        if (!source.code || numeric(source.code)) row.issues.push('Cần mã hạng mục/gói thầu, hoặc nhập mã đầy đủ thủ công.');
        else if (baseCode) row.code = `${baseCode}.${source.code}`;
      } else if (row.parentCode) {
        let sequence = counters.get(row.parentCode) || 0;
        do { sequence++; row.code = `${row.parentCode}.${String(sequence).padStart(2, '0')}`; } while (reserved.has(row.code));
        counters.set(row.parentCode, sequence);
      }
    } else row.issues.push('Chưa ghép được mã cũ: nhập mã đầy đủ hoặc xác nhận đây là việc mới.');
    if (allocating && row.code && baseCode && !row.code.startsWith(baseCode + '.')) {
      row.issues.push('Mã đầy đủ phải nối tiếp mã dự án/hạng mục/công việc cha. Không tự sửa mã đầy đủ đã nhập.');
    }
    if (row.code && !validCode(row.code)) row.issues.push('Mã đề xuất không hợp lệ.');
    const matches = tasksByCode.get(row.code) || [];
    if (matches.length > 1) row.issues.push('Mã trùng trong danh sách 1Office.');
    if (options.mode === 'update' && sourceMatches.length === 1 && row.code !== source.code) row.issues.push('Không được đổi mã công việc đã khớp với 1Office.');
    if ((options.mode === 'initial' || explicitNew) && matches.length) row.issues.push('Mã đã có trên 1Office, không được coi là việc mới.');
    if (options.mode === 'update' && !explicitNew && row.code && matches.length !== 1) row.issues.push('Mã ghép chưa xác định duy nhất công việc hiện có.');
    if (row.code) reserved.add(row.code);
    if (!row.code && !row.issues.length) row.issues.push('Chưa đề xuất được mã.');
    if (!row.issues.length) {
      row.kind = options.mode === 'update' && !explicitNew ? 'existing' : 'proposed';
      if (row.kind === 'existing') row.internalId = matches[0].internalId;
    }
    byNo.set(row.no, row);
  }
  const proposedCodes = new Map<string, NormalizedImportRow[]>();
  for (const row of rows) { if (row.code) { const items = proposedCodes.get(row.code) || []; items.push(row); proposedCodes.set(row.code, items); } }
  for (const items of proposedCodes.values()) if (items.length > 1) for (const row of items) { row.kind = 'blocked'; row.internalId = null; row.issues.push('Mã đề xuất trùng giữa nhiều dòng.'); }
  // Parents precede children. Propagate a late-discovered code collision downstream.
  for (const row of rows) if (row.parentCode && proposedCodes.get(row.parentCode)?.some((parent) => parent.kind === 'blocked')) {
    row.kind = 'blocked'; row.internalId = null; row.issues.push('Mã cha cần xử lý, chưa dùng đề xuất cấp con.');
  }
  return { rows, errors: [...new Set(errors)], warnings: [
    'Chỉ là đề xuất mã và nhóm cấp theo cấu hình, chưa duyệt import và không tạo lệnh ghi.',
    'No. giữ nguyên để truy vết. Quan hệ cha/con lấy từ cấu trúc file để duyệt, không thay đổi quan hệ hiện có trên 1Office.',
    'Sinh mã tự động chỉ dành cho lần lập mã đầu hoặc những dòng bạn xác nhận là mới. Không dùng mã sinh lại để nhận diện công việc cũ.',
    'Mã mới ghép theo dự án → hạng mục → gói thầu → công việc cha/con. Dấu chấm trong mã danh mục không dùng để đếm cấp.',
    catalog ? 'Đã kiểm tra mã gói thầu theo danh mục ERP được nạp. Danh mục hạng mục ERP dùng đối chiếu mã dự án, không xác minh hạng mục tiến độ hay số thứ tự segment. Gợi ý theo tên không chứng minh cùng công việc.' : 'Chưa nạp danh mục ERP: mã thành phần chưa được xác minh theo danh mục.',
    'Chưa chuyển đổi ngày, trạng thái, người thực hiện hoặc MQH. Không mặc định giá trị trống là yêu cầu xóa dữ liệu.',
  ] };
}
