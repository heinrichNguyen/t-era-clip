import React, { useEffect, useRef, useState } from 'react';
import { reviewProjectImport, type ImportSheet } from '../oneoffice/projectImportReview';
import type { OneOfficeProjectInventory } from '../oneoffice/types';
import { ProjectImportNormalization } from './ProjectImportNormalization';
import { ProjectCodeNoReconciliation } from '../import-review/ProjectCodeNoReconciliation';

const labels = { blocked: 'Cần xử lý', matched: 'Khớp mã, cần duyệt trường', unmatched: 'Chưa có mã tương ứng', unavailable: 'Chưa đối chiếu 1Office' };

export function ProjectImportReview({ inventory, disabled, primaryReconciliation = false, sourceTabId, onBusyChange }: { inventory: OneOfficeProjectInventory | null; disabled: boolean; primaryReconciliation?: boolean; sourceTabId?: number; onBusyChange?: (busy: boolean) => void }) {
  const [sheets, setSheets] = useState<ImportSheet[]>([]);
  const [selected, setSelected] = useState(0);
  const [fileName, setFileName] = useState('');
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);
  const [page, setPage] = useState(0);
  const [onlyIssues, setOnlyIssues] = useState(false);
  const generation = useRef(0);
  const worker = useRef<Worker | null>(null);
  const timer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const input = useRef<HTMLInputElement>(null);
  const stop = () => { worker.current?.terminate(); worker.current = null; if (timer.current) clearTimeout(timer.current); timer.current = null; };
  const clear = () => { generation.current++; stop(); setSheets([]); setFileName(''); setError(''); setLoading(false); setPage(0); if (input.current) input.current.value = ''; };
  useEffect(() => () => { generation.current++; stop(); }, []);
  const load = async (file: File) => {
    clear();
    if (!/\.xlsx$/i.test(file.name) || file.size > 5 * 1024 * 1024) { setError('Chỉ nhận file .xlsx tối đa 5 MB.'); return; }
    setLoading(true); setFileName(file.name); setSelected(0); setOnlyIssues(false);
    const id = generation.current;
    try {
      const buffer = await file.arrayBuffer();
      if (generation.current !== id) return;
      const instance = new Worker(new URL('../oneoffice/projectImport.worker.ts', import.meta.url), { type: 'module' });
      worker.current = instance;
      const fail = () => { if (generation.current === id) { stop(); setLoading(false); setError('Không đọc được file trong 30 giây hoặc file không hợp lệ. Hãy kiểm tra và thử lại.'); } };
      timer.current = setTimeout(fail, 30000);
      instance.onerror = fail;
      instance.onmessage = (event: MessageEvent<{ sheets?: ImportSheet[]; error?: string }>) => {
        if (generation.current !== id) return;
        stop(); setLoading(false);
        if (event.data.error) setError(event.data.error);
        else setSheets(event.data.sheets || []);
      };
      instance.postMessage(buffer, [buffer]);
    } catch { if (generation.current === id) { stop(); setLoading(false); setError('Không đọc được file đã chọn.'); } }
  };
  let report: ReturnType<typeof reviewProjectImport> | null = null;
  let reviewError = '';
  if (sheets[selected]) {
    try { report = reviewProjectImport(sheets[selected], inventory); }
    catch (e) { reviewError = e instanceof Error ? e.message : 'Không kiểm tra được cấu trúc sheet.'; }
  }
  const visible = report?.rows.filter((r) => !onlyIssues || r.issues.length) || [];
  const lastPage = Math.max(0, Math.ceil(visible.length / 50) - 1);
  const current = Math.min(page, lastPage);
  return <section className="rounded-2xl border border-blue-200 bg-white p-6">
    <h2 className="text-lg font-semibold">{primaryReconciliation ? 'Bước 1: chọn file được phê duyệt để đối chiếu' : 'Kiểm tra file MS Project trước import'}</h2>
    <p className="mt-2 text-sm text-slate-600">Đọc Excel trên máy, không gửi file ra ngoài và không ghi lên 1Office. Dữ liệu chỉ giữ trong trang này. {primaryReconciliation && 'Dùng bảng mã/No. theo ID bên dưới làm luồng chính; không cần sinh lại mã qua danh mục ERP nếu file đã có mã đầy đủ được duyệt.'}</p>
    <label className="mt-4 block font-semibold" htmlFor="project-import-file">Chọn file Excel (.xlsx, tối đa 5 MB)</label>
    <input ref={input} id="project-import-file" type="file" accept=".xlsx" disabled={disabled} className="mt-2 block" onChange={(e) => { const file = e.target.files?.[0]; if (file) void load(file); }} />
    {(sheets.length > 0 || loading || fileName) && <button type="button" disabled={disabled} onClick={clear} className="mt-3 text-sm text-red-700">Xóa file và kết quả kiểm tra khỏi trang</button>}
    {loading && <p role="status" className="mt-3">Đang đọc file trên máy…</p>}
    {(error || reviewError) && <p role="alert" className="mt-3 text-red-700">{error || reviewError}</p>}
    {sheets.length > 0 && <div className="mt-4"><label htmlFor="project-import-sheet">Sheet cần kiểm tra</label><select id="project-import-sheet" className="ml-3 rounded border p-2" value={selected} disabled={disabled} onChange={(e) => { setSelected(Number(e.target.value)); setPage(0); }}>{sheets.map((sheet, i) => <option key={i} value={i}>{sheet.sheet}</option>)}</select></div>}
    {report && <div className="mt-4">
      {primaryReconciliation && <ProjectCodeNoReconciliation key={`${generation.current}:${selected}:${inventory?.project.url ?? 'local'}:${inventory?.capturedAt ?? ''}`} report={report} inventory={inventory} disabled={disabled} sourceTabId={sourceTabId} onBusyChange={onBusyChange} />}
      <details open={!primaryReconciliation} className="mt-5"><summary className="cursor-pointer font-semibold">Kiểm tra cấu trúc file và mã nguồn (bổ trợ)</summary>
      <p className="font-semibold">{fileName}: {report.rows.length} dòng, {report.rows.filter((r) => r.match === 'blocked').length} dòng cần xử lý, {report.rows.filter((r) => r.match === 'matched').length} dòng khớp mã.</p>
      <p className="mt-2 text-sm">{Object.entries(report.levels).map(([level, count]) => `Cấp ${level}: ${count}`).join(' · ')}</p>
      <ul className="mt-3 list-inside list-disc text-sm text-amber-900">{report.warnings.map((w) => <li key={w}>{w}</li>)}</ul>
      <label className="mt-4 block text-sm"><input type="checkbox" checked={onlyIssues} onChange={(e) => { setOnlyIssues(e.target.checked); setPage(0); }} /> Chỉ hiện dòng cần xử lý</label>
      <div className="mt-3 overflow-auto"><table className="w-full text-left text-sm"><thead><tr>{['Dòng Excel', 'No. nguồn', 'Mã nguồn', 'Tên công việc', 'Cấp / No. cha gợi ý', 'Đối chiếu', 'Lưu ý'].map((h) => <th className="border-b p-2" key={h}>{h}</th>)}</tr></thead>
        <tbody>{visible.slice(current * 50, (current + 1) * 50).map((r) => <tr key={r.row}><td className="border-b p-2">{r.row}</td><td className="border-b p-2">{r.no}</td><td className="border-b p-2">{r.code || '(trống)'}</td><td className="border-b p-2">{r.name}</td><td className="border-b p-2">{r.level ?? '?'} / {r.parentCandidate || '—'}</td><td className="border-b p-2">{labels[r.match]}{r.internalId ? ` (ID ${r.internalId})` : ''}</td><td className="border-b p-2">{r.issues.join('; ')}</td></tr>)}</tbody></table></div>
      <div className="mt-3 flex items-center gap-4"><button type="button" disabled={current === 0} onClick={() => setPage(current - 1)}>Trang trước</button><span>Trang {current + 1}/{lastPage + 1}, {visible.length} dòng</span><button type="button" disabled={current >= lastPage} onClick={() => setPage(current + 1)}>Trang sau</button></div>
      </details>
      {primaryReconciliation ? <details className="mt-5"><summary className="cursor-pointer font-semibold">Lập mã lần đầu / tham khảo ERP và các trường khác (bổ trợ)</summary><ProjectImportNormalization key={`${generation.current}:${selected}`} report={report} inventory={inventory} disabled={disabled} /></details> : <ProjectImportNormalization key={`${generation.current}:${selected}`} report={report} inventory={inventory} disabled={disabled} />}
    </div>}
  </section>;
}
