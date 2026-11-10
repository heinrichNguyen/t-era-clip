import React, { useState } from 'react';
import { buildImportDiff, IMPORT_DIFF_FIELDS, type ImportDiffField } from '../oneoffice/projectImportDiff';
import type { ImportReview } from '../oneoffice/projectImportReview';
import type { proposeImportNormalization } from '../oneoffice/projectImportNormalization';
import type { OneOfficeProjectInventory } from '../oneoffice/types';

const rowLabels = { changed: 'Có thay đổi dự kiến', same: 'Không đổi trong các trường đã chọn', review: 'Cần kiểm tra thêm', draft: 'Bản nháp việc mới, chưa tạo', blocked: 'Chưa đủ điều kiện so sánh', excluded: 'Ngoài phạm vi' };
const fieldLabels = { changed: 'Thay đổi dự kiến', same: 'Không đổi', keep: 'Giữ nguyên', unknown: 'Chưa đủ dữ liệu', invalid: 'Dữ liệu cần kiểm tra', draft: 'Dữ liệu bản nháp' };

export function ProjectImportDiff({ report, proposal, inventory, disabled }: { report: ImportReview; proposal: ReturnType<typeof proposeImportNormalization>; inventory: OneOfficeProjectInventory | null; disabled: boolean }) {
  const headers = Object.keys(report.rows[0]?.values || {});
  const [selected, setSelected] = useState<ImportDiffField[]>([]);
  const [columns, setColumns] = useState<Partial<Record<ImportDiffField, string>>>(() => Object.fromEntries(IMPORT_DIFF_FIELDS.map((f) => [f.id, headers.includes(f.suggestedColumn) ? f.suggestedColumn : ''])));
  const [confirmPlanDates, setConfirmPlanDates] = useState(false);
  const [confirmProject, setConfirmProject] = useState(false);
  const [show, setShow] = useState(false);
  const [page, setPage] = useState(0);
  const [filter, setFilter] = useState('all');
  const invalidate = () => { setShow(false); setPage(0); };
  const result = show ? buildImportDiff(report, proposal, inventory, { columns: Object.fromEntries(selected.map((id) => [id, columns[id] || ''])), confirmPlanDates, confirmProject }) : null;
  const rows = result?.rows.filter((r) => filter === 'all' || r.state === filter) || [];
  const lastPage = Math.max(0, Math.ceil(rows.length / 25) - 1);
  const current = Math.min(page, lastPage);
  return <section className="mt-6 rounded-xl border border-slate-200 bg-white p-4">
    <h3 className="text-lg font-semibold">Bước 3: So sánh dữ liệu trước và sau</h3>
    <p className="mt-2 text-sm text-slate-600">Chọn trường cần xem và cột tương ứng trong Excel. Đây là so sánh với danh sách đã thu thập, không phải đọc lại trực tiếp và chưa phải phê duyệt ghi.</p>
    <p className="mt-2 rounded-lg bg-blue-50 p-3 text-sm text-blue-900">Ô Excel trống hoặc NA: giữ nguyên. Không đề xuất xóa dữ liệu. Thảo luận, tiến độ hoàn thành, kết quả báo cáo, tệp đính kèm và ngày thực tế không nằm trong danh sách trường được cập nhật ở bước này.</p>
    <fieldset disabled={disabled} className="mt-4 space-y-3">
      {IMPORT_DIFF_FIELDS.map((f) => <div key={f.id} className="grid items-center gap-2 rounded-lg border p-3 md:grid-cols-2">
        <label className="text-sm font-semibold"><input type="checkbox" checked={selected.includes(f.id)} onChange={(e) => { setSelected((old) => e.target.checked ? [...old, f.id] : old.filter((key) => key !== f.id)); invalidate(); }} /> {f.label}</label>
        <div><label htmlFor={`diff-column-${f.id}`} className="text-xs text-slate-600">Cột Excel dùng cho {f.label.toLowerCase()}</label><select id={`diff-column-${f.id}`} disabled={!selected.includes(f.id)} className="mt-1 w-full rounded border p-2 text-sm" value={columns[f.id] || ''} onChange={(e) => { setColumns((old) => ({ ...old, [f.id]: e.target.value })); setConfirmPlanDates(false); invalidate(); }}><option value="">Chọn cột Excel</option>{headers.map((h) => <option key={h} value={h}>{h}</option>)}</select></div>
      </div>)}
      {selected.some((key) => IMPORT_DIFF_FIELDS.some((f) => f.id === key && f.date)) && <label className="block text-sm"><input type="checkbox" checked={confirmPlanDates} onChange={(e) => { setConfirmPlanDates(e.target.checked); invalidate(); }} /> Tôi xác nhận các cột ngày đã chọn là ngày kế hoạch hoặc kế hoạch gốc, không phải ngày thực tế.</label>}
      {proposal.rows.some((r) => r.kind === 'existing') && <label className="block text-sm"><input type="checkbox" checked={confirmProject} onChange={(e) => { setConfirmProject(e.target.checked); invalidate(); }} /> Tôi xác nhận file này thuộc đúng dự án đang đối chiếu: {inventory?.project.name || '(chưa có danh sách)'}.</label>}
    </fieldset>
    <button type="button" disabled={disabled} className="mt-4 rounded-lg bg-blue-700 px-4 py-2 font-semibold text-white disabled:opacity-50" onClick={() => { setShow(true); setPage(0); }}>Xem bảng so sánh (không ghi)</button>
    {result && <div className="mt-4">
      <p className="text-sm text-slate-600">Thời điểm danh sách 1Office: {result.snapshotAt || 'chưa có'}. Giá trị thiếu từ danh sách không được coi là ô trống đã xác minh.</p>
      {!!result.errors.length && <ul role="alert" className="mt-3 list-inside list-disc text-red-700">{result.errors.map((error) => <li key={error}>{error}</li>)}</ul>}
      <div className="mt-3 grid gap-2 sm:grid-cols-2 lg:grid-cols-3">{Object.entries(rowLabels).map(([state, label]) => <p key={state} className="rounded-lg bg-slate-50 p-3 text-sm"><strong>{result.rows.filter((r) => r.state === state).length}</strong> {label}</p>)}</div>
      <label className="mt-4 block text-sm">Lọc kết quả<select aria-label="Lọc kết quả so sánh" className="ml-2 rounded border p-2" value={filter} onChange={(e) => { setFilter(e.target.value); setPage(0); }}><option value="all">Tất cả</option>{Object.entries(rowLabels).map(([state, label]) => <option key={state} value={state}>{label}</option>)}</select></label>
      <div className="mt-3 overflow-auto"><table className="w-full text-left text-sm"><thead><tr>{['Công việc', 'Trường / cột Excel', 'Trước (danh sách 1Office)', 'Sau (dự kiến từ file)', 'Kết quả'].map((h) => <th key={h} className="border-b p-2">{h}</th>)}</tr></thead><tbody>
        {rows.slice(current * 25, (current + 1) * 25).map((r) => <React.Fragment key={r.row}>
          <tr className="bg-slate-50"><td colSpan={5} className="border-b p-3"><strong>Dòng Excel {r.row}: {r.code || '(chưa có mã)'}</strong> · {r.name}<p>{rowLabels[r.state]}{r.internalId ? `, ID ${r.internalId}` : ''}</p>{r.issues.length > 0 && <p className="mt-1 text-amber-900">{r.issues.join('; ')}</p>}</td></tr>
          {r.fields.map((f) => <tr key={f.field}><td className="border-b p-2">{r.code}</td><td className="border-b p-2">{IMPORT_DIFF_FIELDS.find((definition) => definition.id === f.field)?.label}<p className="text-slate-500">Cột: {f.column}</p></td><td className="border-b p-2">{f.before || (f.beforeKnown ? '(trống)' : '(chưa xác nhận giá trị cũ)')}</td><td className="border-b p-2">{f.state === 'keep' ? 'Giữ nguyên giá trị trên 1Office' : f.after}</td><td className="border-b p-2">{fieldLabels[f.state]}<p className="text-amber-900">{f.note}</p></td></tr>)}
        </React.Fragment>)}
      </tbody></table></div>
      <div className="mt-3 flex flex-wrap items-center gap-4"><button type="button" disabled={current === 0} onClick={() => setPage(current - 1)}>Trang so sánh trước</button><span>Trang {current + 1}/{lastPage + 1}, {rows.length} công việc</span><button type="button" disabled={current >= lastPage} onClick={() => setPage(current + 1)}>Trang so sánh sau</button></div>
      <p className="mt-3 text-sm text-amber-900">Chưa có nút áp dụng. Việc mới chỉ là bản nháp; công việc không có trong file được giữ nguyên. Mã, ID và quan hệ cha/con không được sửa ở bước này.</p>
    </div>}
  </section>;
}
