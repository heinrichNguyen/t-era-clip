import type { OneOfficeProjectInventory } from './types';

export interface NoUpdateRow {
  internalId: string;
  code: string;
  name: string;
  before: string;
  after: string;
}

// This module only plans changes. It must not send requests to 1Office.
export function buildNoUpdatePlan(inventory: OneOfficeProjectInventory, text: string) {
  const errors: string[] = [];
  const rows: NoUpdateRow[] = [];
  const additions: { code: string; no: string }[] = [];
  const tasksByCode = new Map<string, typeof inventory.tasks>();
  const ids = new Set<string>();
  for (const task of inventory.tasks) {
    if (!/^\d+$/.test(task.internalId) || ids.has(task.internalId)) {
      errors.push('Danh sách nguồn có ID công việc trống, không hợp lệ hoặc trùng.');
    }
    ids.add(task.internalId);
    const code = task.code?.trim();
    if (code) tasksByCode.set(code, [...(tasksByCode.get(code) ?? []), task]);
  }
  if (!inventory.collection.complete) {
    errors.push('Danh sách nguồn chưa được xác nhận đầy đủ. Cần kiểm tra lại trước khi ghi dữ liệu.');
  }
  const supplied = new Map<string, string>();
  const lines = text.split(/\r?\n/).filter((line) => line.trim());
  if (lines.length === 0) errors.push('Hãy dán hai cột: mã công việc và No. mới, không kèm tiêu đề.');
  if (lines.length > 10000) errors.push('Chỉ hỗ trợ tối đa 10.000 dòng trong một bản đối chiếu.');
  for (const [index, line] of lines.slice(0, 10000).entries()) {
    const cells = line.split('\t').map((cell) => cell.trim());
    const [code, no] = cells;
    if (cells.length !== 2 || !code || !/^[1-9]\d*(?:\.[1-9]\d*)*$/.test(no)) {
      errors.push(`Dòng ${index + 1}: cần đúng hai cột, No. dạng 1, 1.2 hoặc 1.2.3 (không giới hạn cấp).`);
      continue;
    }
    if (supplied.has(code)) {
      errors.push(`Dòng ${index + 1}: mã ${code} bị lặp trong dữ liệu mới.`);
      continue;
    }
    supplied.set(code, no);
    const matches = tasksByCode.get(code) ?? [];
    if (matches.length === 0) {
      additions.push({ code, no });
      errors.push(`Mã ${code} chưa có trong dự án. Chưa hỗ trợ tạo công việc mới.`);
    } else if (matches.length > 1) {
      errors.push(`Mã ${code} trùng trong danh sách nguồn, không thể xác định đúng công việc.`);
    } else {
      const task = matches[0];
      rows.push({ internalId: task.internalId, code, name: task.name, before: typeof task.sourceNo === 'string' ? task.sourceNo : task.no, after: no });
    }
  }
  // Validate the final numbering, including tasks absent from the submitted list.
  const replacement = new Map(rows.map((row) => [row.internalId, row.after]));
  const finalNumbers = new Map<string, string[]>();
  for (const task of inventory.tasks) {
    const no = replacement.get(task.internalId) ?? task.sourceNo ?? task.no;
    if (!no) continue;
    finalNumbers.set(no, [...(finalNumbers.get(no) ?? []), task.code || task.internalId]);
  }
  for (const addition of additions) {
    finalNumbers.set(addition.no, [...(finalNumbers.get(addition.no) ?? []), addition.code]);
  }
  for (const [no, codes] of finalNumbers) {
    if (codes.length > 1) errors.push(`No. ${no} bị trùng sau đối chiếu: ${codes.join('; ')}.`);
  }
  return {
    projectId: inventory.project.internalId,
    capturedAt: inventory.capturedAt,
    rows,
    additions,
    changes: rows.filter((row) => row.before !== row.after),
    unchanged: rows.filter((row) => row.before === row.after).length,
    untouched: inventory.tasks.length - rows.length,
    errors: [...new Set(errors)],
  };
}
