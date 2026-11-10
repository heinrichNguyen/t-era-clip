import React, { useEffect, useRef, useState } from 'react';
import { reconcileProjectCodeNo, type CodeNoOptions, type CodeNoPreview } from '../oneoffice/projectCodeNoReconciliation';
import type { ImportReview } from '../oneoffice/projectImportReview';
import type { OneOfficeProjectInventory } from '../oneoffice/types';
import { collectTestTaskDetails, type TestTaskDetails } from '../oneoffice/testTaskDetails';
import { TEST_BASELINE_ORIGIN, TEST_BASELINE_PROJECT } from '../oneoffice/testBaselineScope';

const initial: CodeNoOptions = { idColumn: '', oldCodeColumn: '', codeColumn: 'Mã gói thầu/Mã CV', noColumn: 'No.', mappings: '', approvedFile: false, confirmedProject: false, noMeaning: '' };
export function ProjectCodeNoReconciliation({ report, inventory, disabled, sourceTabId, onBusyChange }: {
  report: ImportReview; inventory: OneOfficeProjectInventory | null; disabled: boolean; sourceTabId?: number; onBusyChange?: (busy: boolean) => void;
}) {
  const [options, setOptions] = useState(initial);
  const [preview, setPreview] = useState<CodeNoPreview | null>(null);
  const [evidence, setEvidence] = useState<Record<string, TestTaskDetails>>({});
  const [busy, setBusy] = useState(false); const [message, setMessage] = useState(''); const [page, setPage] = useState(0);
  const disposed = useRef(false); const running = useRef(false); const stop = useRef(false);
  useEffect(() => { disposed.current = false; return () => { disposed.current = true; stop.current = true; }; }, []);
  const update = (value: Partial<CodeNoOptions>) => { setOptions((old) => ({ ...old, ...value })); setPreview(null); setPage(0); };
  const columns = [...new Set(report.rows.flatMap((row) => Object.keys(row.values)))];
  let canRead = false;
  try { canRead = !!inventory && inventory.project.internalId === TEST_BASELINE_PROJECT && new URL(inventory.project.url).origin === TEST_BASELINE_ORIGIN && Number.isSafeInteger(sourceTabId) && sourceTabId! > 0; } catch { /* Keep malformed contexts read-disabled. */ }
  const compare = () => { setPreview(reconcileProjectCodeNo(report, inventory, options, evidence)); setPage(0); };
  const read = async () => {
    if (running.current || disabled || !canRead || !inventory) return;
    running.current = true; stop.current = false; setBusy(true); setPreview(null); onBusyChange?.(true);
    const next = { ...evidence };
    const mapping = reconcileProjectCodeNo(report, inventory, options, evidence);
    const mapped = new Set(mapping.rows.flatMap((row) => row.internalId ? [row.internalId] : []));
    const tasks = [...inventory.tasks.filter((task) => mapped.has(task.internalId)), ...inventory.tasks.filter((task) => !mapped.has(task.internalId))]
      .filter((task) => {
        const item = next[task.internalId];
        const knownCode = item?.identityVerified && ['value', 'empty'].includes(item.code.state) || typeof task.code === 'string';
        const knownNo = item?.identityVerified && ['value', 'empty'].includes(item.no.state) || typeof task.sourceNo === 'string';
        return !knownCode || !knownNo;
      }).slice(0, 20);
    try {
      let failed = 0; let completed = 0;
      for (const [index, task] of tasks.entries()) {
        if (disposed.current || stop.current) break;
        setMessage(`Đang đọc mã/No. ${index + 1}/${tasks.length}, ID ${task.internalId}. Không đọc thảo luận, không ghi.`);
        try { next[task.internalId] = await collectTestTaskDetails(sourceTabId!, task.internalId); if (!next[task.internalId].identityVerified) failed++; }
        catch { failed++; }
        completed++;
      }
      if (!disposed.current) { setEvidence(next); setPreview(null); setMessage(`Đã thử đọc ${completed}/${tasks.length} công việc trong đợt này, ${failed} yêu cầu lỗi hoặc chưa xác minh. Bấm xem bảng đối chiếu để dùng kết quả. Chưa ghi dữ liệu.`); }
    } finally { running.current = false; if (!disposed.current) { setBusy(false); onBusyChange?.(false); } }
  };
  const exportPreview = () => {
    if (!preview) return;
    const data = { schemaVersion: 1, project: inventory?.project, sheet: report.sheet, exportedAt: new Date().toISOString(), options, evidence, ...preview };
    const url = URL.createObjectURL(new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' }));
    const a = document.createElement('a'); a.href = url; a.download = `tera-clip-code-no-preview-${inventory?.project.internalId ?? 'local'}.json`; a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  };
  const lastPage = Math.max(0, Math.ceil((preview?.rows.length ?? 0) / 50) - 1);
  return <section className="mt-6 space-y-4 rounded-2xl border-2 border-blue-600 bg-blue-50 p-5" aria-label="Đối chiếu mã và No. theo ID">
    <h2 className="text-xl font-bold">Đối chiếu mã và No. theo file được duyệt</h2>
    <p className="text-sm">ID xác định công việc, mã và No. là giá trị được thay đổi. Việc đã tạo thủ công hoặc import cũng dùng ID hiện hữu để đặt mã. Chưa có ID thì cần ghép lại hoặc tạo/import trước, không tự tạo hay xóa công việc.</p>
    <div className="grid gap-3 sm:grid-cols-2">{([
      ['codeColumn', 'Cột mã đầy đủ được duyệt'], ['noColumn', 'Cột No. được duyệt'], ['idColumn', 'Cột ID 1Office (nếu có)'], ['oldCodeColumn', 'Cột mã cũ để nhận diện (nếu có)'],
    ] as const).map(([key, label]) => <label key={key} className="text-sm font-semibold">{label}<select aria-label={label} className="mt-1 block w-full rounded-lg border p-2" disabled={disabled || busy} value={options[key]} onChange={(event) => update({ [key]: event.target.value })}><option value="">Chưa chọn</option>{columns.map((column) => <option key={column} value={column}>{column}</option>)}</select></label>)}</div>
    <label className="block text-sm font-semibold">Ánh xạ ID xác nhận thủ công (dòng Excel và ID)<textarea aria-label="Ánh xạ dòng Excel và ID" value={options.mappings} disabled={disabled || busy} onChange={(event) => update({ mappings: event.target.value })} placeholder={'24\t7698\n25\t7699'} className="mt-1 block min-h-24 w-full rounded-lg border p-2 font-mono" /></label>
    <p className="text-xs text-slate-600">Nếu không có ID, chỉ dùng mã cũ khớp duy nhất để ghép ban đầu. Không ghép theo tên hoặc No.; mã đích khác mã cũ không làm mất ID đã xác nhận. File phải chứa mã đầy đủ, không phải mã gói thầu thành phần.</p>
    <p className="text-xs text-slate-600">Bảng dùng bản hiện trạng đã thu thập và các đợt đọc bổ sung, không phải ảnh chụp đồng thời. Trước khi bật ghi cần đọc lại hiện trạng và xác minh cách cập nhật cả hai trường.</p>
    <label className="block text-sm">Ý nghĩa trường No.<select aria-label="Ý nghĩa trường No." value={options.noMeaning} disabled={disabled || busy} onChange={(event) => update({ noMeaning: event.target.value as CodeNoOptions['noMeaning'] })} className="ml-2 rounded-lg border p-2"><option value="">Chưa xác nhận</option><option value="custom-no">Trường No. tùy chỉnh trong chi tiết công việc</option><option value="outline">Số thứ tự trên cây/Gantt (chưa hỗ trợ ở luồng này)</option></select></label>
    <label className="block text-sm"><input type="checkbox" disabled={disabled || busy} checked={options.approvedFile} onChange={(event) => update({ approvedFile: event.target.checked })} /> Tôi xác nhận file này được duyệt làm nguồn mã và No. đích.</label>
    <label className="block text-sm"><input type="checkbox" disabled={disabled || busy} checked={options.confirmedProject} onChange={(event) => update({ confirmedProject: event.target.checked })} /> Tôi xác nhận đúng dự án: {inventory?.project.name ?? '(chưa có danh sách hệ thống)'}.</label>
    <div className="flex flex-wrap gap-3"><button type="button" disabled={disabled || busy} onClick={compare} className="rounded-xl bg-blue-700 px-4 py-3 text-sm font-semibold text-white disabled:opacity-50">Xem bảng mã và No. trước/sau</button>
      {canRead && <button type="button" disabled={disabled || busy} onClick={() => void read()} className="rounded-xl border border-blue-700 bg-white px-4 py-3 text-sm text-blue-700 disabled:opacity-50">Đọc hiện trạng mã/No. còn thiếu (tối đa 20 công việc)</button>}
      {busy && <button type="button" onClick={() => { stop.current = true; }} className="text-sm text-red-700">Dừng sau yêu cầu đọc hiện tại</button>}</div>
    {!canRead && <p className="text-xs text-amber-900">Chưa có bộ đọc bổ sung mã/No. cho ngữ cảnh này. Dữ liệu thiếu sẽ chặn phương án, không lấy số thứ tự cây làm No. gốc.</p>}
    {message && <p role="status" className="text-sm">{message}</p>}
    {preview && <>
      <p className="font-semibold">{preview.rows.filter((row) => row.state === 'changed').length} dòng thay đổi · {preview.rows.filter((row) => row.state === 'same').length} không đổi · {preview.rows.filter((row) => row.state === 'review').length} cần xác nhận · {preview.unchangedOutsideFile} công việc ngoài file giữ nguyên</p>
      {preview.errors.length > 0 && <ul role="alert" className="list-inside list-disc text-sm text-red-700">{preview.errors.map((error) => <li key={error}>{error}</li>)}</ul>}
      <div className="overflow-auto"><table className="w-full text-left text-sm"><thead><tr>{['Dòng / công việc', 'ID / cách ghép', 'Mã hiện tại → mã duyệt', 'No. hiện tại → No. duyệt', 'Kết quả'].map((label) => <th key={label} className="border-b p-2">{label}</th>)}</tr></thead><tbody>{preview.rows.slice(page * 50, (page + 1) * 50).map((row) => <tr key={row.row}>
        <td className="border-b p-2">{row.row}: {row.name}</td><td className="border-b p-2">{row.internalId ?? '(chưa ghép ID)'}<br />{({ id: 'ID đã chọn', 'old-code': 'Mã cũ duy nhất', 'target-code': 'Mã đang khớp duy nhất', unresolved: 'Cần xác nhận' })[row.match]}</td>
        <td className="break-all border-b p-2">{row.beforeCode === null ? '(chưa đọc được)' : row.beforeCode || '(trống)'} → {row.afterCode || '(thiếu)'}</td>
        <td className="border-b p-2">{row.beforeNo === null ? '(chưa đọc được)' : row.beforeNo || '(trống)'} → {row.afterNo || '(thiếu)'}</td>
        <td className="border-b p-2">{({ changed: 'Dự kiến thay đổi', same: 'Không đổi', review: 'Cần xác nhận' })[row.state]}<p className="text-amber-900">{row.issues.join('; ')}</p></td>
      </tr>)}</tbody></table></div>
      <div className="flex gap-3 text-sm"><button type="button" disabled={page === 0} onClick={() => setPage(page - 1)}>Trang trước</button><span>{page + 1}/{lastPage + 1}</span><button type="button" disabled={page >= lastPage} onClick={() => setPage(page + 1)}>Trang sau</button></div>
      <p className="text-sm">{preview.operations.length ? `Phương án gồm ${preview.operations.length} bước dự kiến, ${preview.operations.filter((step) => step.temporary).length} bước dùng giá trị tạm để phá vòng hoán đổi.` : 'Chưa có phương án thay đổi hoặc còn điều kiện cần xử lý.'} Đây là bản xem trước, chưa phải phê duyệt ghi. Không thay đổi ID, cha/con, thảo luận, tiến độ hoặc báo cáo.</p>
      <button type="button" disabled={busy || disabled} onClick={exportPreview} className="rounded-xl border border-blue-700 px-4 py-2 text-sm text-blue-700">Tải bảng đối chiếu JSON (chưa ghi)</button>
    </>}
  </section>;
}
