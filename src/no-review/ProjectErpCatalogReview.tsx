import React, { useEffect, useMemo, useRef, useState } from 'react';
import { erpPackageNameKey, indexErpPackageNames, indexErpPackages, suggestErpPackages, type ProjectErpCatalog } from '../oneoffice/projectErpCatalog';
import type { ImportReview } from '../oneoffice/projectImportReview';
import type { NormalizationOptions } from '../oneoffice/projectImportNormalization';

interface Props {
  report: ImportReview;
  options: NormalizationOptions;
  catalog: ProjectErpCatalog | null;
  disabled: boolean;
  onCatalog: (catalog: ProjectErpCatalog | null) => void;
  onLoading: (loading: boolean) => void;
  onSelect: (row: number, code: string) => void;
}

export function ProjectErpCatalogReview({ report, options, catalog, disabled, onCatalog, onLoading, onSelect }: Props) {
  const [sheet, setSheet] = useState('TDR');
  const [fileName, setFileName] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [page, setPage] = useState(0);
  const [queries, setQueries] = useState<Record<number, string>>({});
  const [onlyPending, setOnlyPending] = useState(true);
  const worker = useRef<Worker | null>(null);
  const timer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const generation = useRef(0);
  const input = useRef<HTMLInputElement>(null);
  const stop = () => { worker.current?.terminate(); worker.current = null; if (timer.current) clearTimeout(timer.current); timer.current = null; };
  useEffect(() => () => { generation.current++; stop(); }, []);
  useEffect(() => { setPage(0); setQueries({}); }, [options.packageLevel, options.mode, options.firstTaskLevel]);
  const clear = () => {
    generation.current++; stop(); setLoading(false); onLoading(false); onCatalog(null);
    setFileName(''); setError(''); setPage(0); setQueries({}); setOnlyPending(true); if (input.current) input.current.value = '';
  };
  const load = async (file: File) => {
    clear();
    if (!/\.xlsx$/i.test(file.name) || file.size > 10 * 1024 * 1024) { setError('Chỉ nhận danh mục .xlsx tối đa 10 MB.'); return; }
    if (!sheet.trim() || sheet.trim().length > 100) { setError('Nhập tên sheet cần đọc.'); return; }
    const id = generation.current;
    setFileName(file.name); setLoading(true); onLoading(true);
    const fail = (message: string) => { if (id === generation.current) { generation.current++; stop(); setLoading(false); onLoading(false); setError(message); } };
    timer.current = setTimeout(() => fail('Không đọc được danh mục trong 30 giây. Kiểm tra sheet và thử lại.'), 30000);
    try {
      const buffer = await file.arrayBuffer();
      if (id !== generation.current) return;
      const instance = new Worker(new URL('../oneoffice/projectErpCatalog.worker.ts', import.meta.url), { type: 'module' });
      worker.current = instance;
      instance.onerror = () => fail('Không đọc được danh mục. Kiểm tra file và thử lại.');
      instance.onmessage = (event: MessageEvent<{ catalog?: ProjectErpCatalog; error?: string }>) => {
        if (id !== generation.current) return;
        generation.current++; stop(); setLoading(false); onLoading(false);
        if (event.data.error || !event.data.catalog) setError(event.data.error || 'Không có danh mục hợp lệ.');
        else onCatalog(event.data.catalog);
      };
      instance.postMessage({ buffer, sheet: sheet.trim() }, [buffer]);
    } catch { fail('Không đọc được file danh mục đã chọn.'); }
  };
  const index = useMemo(() => catalog ? indexErpPackages(catalog) : new Map(), [catalog]);
  const names = useMemo(() => catalog ? indexErpPackageNames(catalog) : new Map<string, Set<string>>(), [catalog]);
  const packages = useMemo(() => report.rows.filter((r) => r.level === options.packageLevel), [report, options.packageLevel]);
  const visible = packages.filter((r) => !onlyPending || index.get(r.code)?.length !== 1 || options.erpPackages?.[r.row]);
  const lastPage = Math.max(0, Math.ceil(visible.length / 20) - 1);
  const current = Math.min(page, lastPage);
  return <section className="mt-4 rounded-xl border bg-white p-4" aria-label="Đối chiếu danh mục ERP">
    <h4 className="font-semibold">Danh mục ERP: kiểm tra và chọn mã gói thầu</h4>
    <p className="mt-1 text-sm text-slate-600">Đọc một sheet trên máy, chỉ giữ Danh mục gói thầu và Danh mục hạng mục dùng cho mã dự án. Không giữ danh mục nhân sự, tài khoản hoặc đối tượng khác. Không gửi file ra ngoài.</p>
    <fieldset disabled={disabled} className="mt-3 grid gap-3 md:grid-cols-2">
      <div><label htmlFor="erp-catalog-sheet" className="text-sm font-semibold">Tên sheet danh mục</label>
        <input id="erp-catalog-sheet" aria-describedby="erp-catalog-sheet-help" className="mt-1 w-full rounded border p-2" maxLength={100} value={sheet} onChange={(e) => { clear(); setSheet(e.target.value); }} />
        <p id="erp-catalog-sheet-help" className="text-xs text-slate-600">Mặc định TDR cho file đã cung cấp. Đổi tên sheet trước khi chọn file nếu dùng đơn vị khác.</p></div>
      <div><label htmlFor="erp-catalog-file" className="text-sm font-semibold">Chọn file danh mục ERP (.xlsx)</label>
        <input ref={input} id="erp-catalog-file" aria-describedby="erp-catalog-file-help" type="file" accept=".xlsx" className="mt-2 block w-full" onChange={(e) => { const file = e.target.files?.[0]; if (file) void load(file); }} />
        <p id="erp-catalog-file-help" className="text-xs text-slate-600">Tối đa 10 MB. Danh mục chỉ giữ trong trang này; xóa dữ liệu tạm hoặc đóng trang sẽ bỏ danh mục và lựa chọn.</p></div>
    </fieldset>
    {fileName && <button disabled={disabled} type="button" className="mt-2 text-sm text-red-700" onClick={clear}>Bỏ danh mục ERP và các lựa chọn</button>}
    {loading && <p role="status">Đang đọc sheet danh mục trên máy…</p>}
    {error && <p role="alert" className="mt-2 text-red-700">{error}</p>}
    {catalog && <>
      <p className="mt-3 text-sm font-semibold">{fileName}, sheet {catalog.sheet}: {catalog.projects.length} mã dự án, {catalog.packages.length} mã gói thầu.</p>
      <p className="mt-1 text-sm text-amber-900">Chưa xác minh mã hạng mục tiến độ và segment bằng danh mục này. Có mã trong danh mục không chứng minh gói thầu thuộc dự án hoặc công việc đã có trên 1Office. Tìm theo tên chỉ là gợi ý, luôn cần chọn nếu mã nguồn chưa khớp.</p>
      {catalog.warnings.map((warning) => <p key={warning} className="mt-1 text-sm text-amber-900">{warning}</p>)}
      <details className="mt-3 rounded-lg border p-3"><summary className="cursor-pointer text-sm font-semibold">Xem đối chiếu và chọn mã cho {packages.length} dòng cấp gói thầu</summary>
      <label className="mt-2 block text-sm"><input type="checkbox" disabled={disabled} checked={onlyPending} onChange={(e) => { setOnlyPending(e.target.checked); setPage(0); }} /> Chỉ hiện mã nguồn chưa khớp hoặc đã chọn thay thế</label>
      <div className="mt-3 space-y-3">{visible.slice(current * 20, (current + 1) * 20).map((r) => {
        const selected = options.erpPackages?.[r.row] || '';
        const sourceMatches = index.get(r.code) || [];
        const query = queries[r.row] ?? r.name;
        const suggestions = suggestErpPackages(catalog, query);
        const selectedEntry = index.get(selected)?.[0];
        const sameNameCount = (entry: typeof catalog.packages[number]) => names.get(erpPackageNameKey(entry))?.size || 1;
        const ambiguous = suggestions.entries.some((entry) => sameNameCount(entry) > 1);
        return <div key={r.row} className="rounded-lg border p-3">
          <p className="text-sm font-semibold">Dòng Excel {r.row}: {r.name}</p>
          <p className="text-xs text-slate-600">Mã nguồn: {r.code || '(trống)'}. {sourceMatches.length === 1 ? 'Mã nguồn có trong danh mục ERP (chưa xác minh tên/phạm vi dự án).' : 'Chưa khớp duy nhất mã nguồn với danh mục ERP.'}</p>
          <label htmlFor={`erp-query-${r.row}`} className="mt-2 block text-sm">Tìm mã hoặc tên gói thầu cho dòng {r.row}</label>
          <input id={`erp-query-${r.row}`} aria-describedby={`erp-pick-help-${r.row}`} disabled={disabled} maxLength={300} className="mt-1 w-full rounded border p-2 text-sm" value={query} onChange={(e) => setQueries((previous) => ({ ...previous, [r.row]: e.target.value }))} />
          <label htmlFor={`erp-pick-${r.row}`} className="mt-2 block text-sm">Chọn mã ERP cho dòng {r.row}</label>
          <select id={`erp-pick-${r.row}`} aria-describedby={`erp-pick-help-${r.row}`} disabled={disabled || options.mode === 'update' && !options.newRows.split(/[\s,;]+/).includes(String(r.row))} className="mt-1 w-full rounded border bg-white p-2 text-sm" value={selected} onChange={(e) => onSelect(r.row, e.target.value)}>
            <option value="">Không chọn thay thế, dùng mã trong file nếu khớp</option>
            {selectedEntry && !suggestions.entries.some((entry) => entry.code === selected) && <option value={selected}>{selectedEntry.code}: {selectedEntry.name}{sameNameCount(selectedEntry) > 1 ? ' (nhiều mã cùng mô tả)' : ''}</option>}
            {suggestions.entries.map((entry) => <option key={`${entry.row}:${entry.code}`} value={entry.code} disabled={index.get(entry.code)?.length !== 1}>{entry.code}: {entry.name}{index.get(entry.code)?.length !== 1 ? ' (mã trùng, cần kiểm tra)' : sameNameCount(entry) > 1 ? ` (${sameNameCount(entry)} mã cùng mô tả)` : ''}</option>)}
          </select>
          <p id={`erp-pick-help-${r.row}`} className="mt-1 text-xs text-slate-600">{suggestions.total} gợi ý, hiển thị tối đa 20. Tìm cụ thể hơn để thu hẹp. Lựa chọn chỉ dùng cho lập mã lần đầu hoặc dòng đã xác nhận mới; không đổi mã/ID công việc cũ.</p>
          {(ambiguous || selectedEntry && sameNameCount(selectedEntry) > 1) && <p className="mt-2 rounded bg-amber-50 p-2 text-sm text-amber-900">{selected ? 'Mã được chọn có nhiều mã cùng mô tả. Chỉ giữ lựa chọn sau khi đã xác minh.' : 'Có nhiều mã cùng mô tả: chờ xác nhận với kế toán trưởng nếu chưa phân biệt được.'} Chúng có thể khác giá, nhà thầu hoặc phạm vi; file danh mục hiện không có các thông tin này. Không tự chọn mã đầu tiên. Có thể để trống và tiếp tục kiểm tra dòng khác; dòng này và các công việc con vẫn cần xử lý.</p>}
        </div>;
      })}</div>
      <div className="mt-3 flex items-center gap-3"><button type="button" disabled={disabled || !current} onClick={() => setPage(current - 1)}>Nhóm trước</button><span>Nhóm {current + 1}/{lastPage + 1} ({visible.length} dòng đang lọc)</span><button type="button" disabled={disabled || current >= lastPage} onClick={() => setPage(current + 1)}>Nhóm sau</button></div>
      </details>
    </>}
  </section>;
}
