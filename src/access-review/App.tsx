import React, { useEffect, useRef, useState } from 'react';
import { collectAccessReport, cutoffEnd, loadAccessPageFromTab, reportFromAccessSnapshot, type AccessReport, type AccessSnapshot } from '../oneoffice/accessHistory';
import { downloadAccessWorkbook } from '../oneoffice/export/accessXlsx';
import { useTemporaryData } from '../oneoffice/temporaryData';
import { ClearTemporaryDataButton, TemporaryDataCleared } from '../shared/TemporaryDataControls';
import { TadtLogo } from '../shared/TadtLogo';
import { loadAccessPersonnel } from '../oneoffice/accessPersonnel';
import { activeAccessSnapshot, collectScopedAccessReport, selectAccessAccounts, type AccessFilters } from '../oneoffice/accessScope';
import { CompanyScopePicker } from './CompanyScopePicker';
import { companyScopeLabel } from '../oneoffice/companyScopes';
import { ONEOFFICE_ALLOWED_ORIGIN } from '../oneoffice/origins';

export function AccessReviewApp() {
  const [sourceOrigin] = useState(() => new URLSearchParams(location.search).get('sourceOrigin') ?? ONEOFFICE_ALLOWED_ORIGIN);
  const [cutoff, setCutoff] = useState(() => new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Ho_Chi_Minh', year: 'numeric', month: '2-digit', day: '2-digit' }).format(new Date()));
  const [busy, setBusy] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [progress, setProgress] = useState({ loaded: 0, total: 0 });
  const [report, setReport] = useState<AccessReport | null>(null);
  const [message, setMessage] = useState('');
  const [stage, setStage] = useState<'history' | 'personnel'>('history');
  const [filters, setFilters] = useState<AccessFilters>({ company: '', position: '', department: '' });
  const filtered = Object.values(filters).some((value) => value.trim());
  const changeFilter = (key: keyof AccessFilters, value: string) => {
    setFilters((current) => ({ ...current, [key]: value }));
    snapshot.current = null; setReport(null); setMessage('');
  };
  const controller = useRef<AbortController | null>(null);
  const snapshot = useRef<AccessSnapshot | null>(null);
  const temporary = useTemporaryData(() => {
    controller.current?.abort(); snapshot.current = null;
    setReport(null); setMessage(''); setProgress({ loaded: 0, total: 0 });
  });
  useEffect(() => () => controller.current?.abort(), []);
  const run = async () => {
    setReport(null); setMessage(''); setProgress({ loaded: 0, total: 0 });
    snapshot.current = null;
    try {
      cutoffEnd(cutoff);
      const tabId = Number(new URLSearchParams(location.search).get('sourceTabId'));
      if (!Number.isInteger(tabId) || tabId <= 0) throw new Error('Hãy mở báo cáo từ extension trên tab Lịch sử truy cập.');
      setBusy(true); controller.current = new AbortController();
      let result: AccessReport;
      if (filtered) {
        setStage('personnel');
        const personnel = await loadAccessPersonnel(tabId, null, controller.current.signal,
          (loaded, total) => setProgress({ loaded, total }), sourceOrigin);
        const accounts = selectAccessAccounts(personnel, filters);
        setStage('history');
        result = await collectScopedAccessReport(cutoff, accounts, (page, account) => loadAccessPageFromTab(tabId, page, account, sourceOrigin),
          controller.current.signal, (loaded, total) => setProgress({ loaded, total }),
          `Nhật ký của ${accounts.length} tài khoản hiện còn active, khớp cơ cấu hiện tại. Phạm vi công ty: ${companyScopeLabel(filters.company)}; vị trí: ${filters.position || 'Tất cả'}; phòng ban/mã: ${filters.department || 'Tất cả'}.`);
        result.personnel = personnel;
        if (result.snapshot) result.snapshot.personnel = personnel;
      } else {
        setStage('history');
        result = await collectAccessReport(cutoff, (page) => loadAccessPageFromTab(tabId, page, undefined, sourceOrigin), controller.current.signal,
          (loaded, total) => setProgress({ loaded, total }));
        setStage('personnel'); setProgress({ loaded: 0, total: 0 });
        const personnel = await loadAccessPersonnel(tabId, result.snapshot?.records.map((row) => row.account) ?? [],
          controller.current.signal, (loaded, total) => setProgress({ loaded, total }), sourceOrigin);
        result.personnel = personnel;
        if (result.snapshot) result.snapshot.personnel = personnel;
      }
      if (temporary.clearedRef.current) return;
      if (!result.snapshot || !result.personnel) throw new Error('Chưa có dữ liệu xác minh tài khoản active.');
      snapshot.current = activeAccessSnapshot(result.snapshot, result.personnel);
      setReport(reportFromAccessSnapshot(snapshot.current, cutoff));
    } catch (error) { setMessage(error instanceof Error ? error.message : 'Không thể thu thập nhật ký.'); }
    finally { setBusy(false); }
  };
  const download = async () => {
    if (temporary.clearedRef.current || !report || !window.confirm('Xuất dữ liệu truy cập nội bộ ra Excel? Chỉ chia sẻ với người được phép nhận báo cáo.')) return;
    setExporting(true);
    try { await downloadAccessWorkbook(report); setMessage('Đã xuất Excel.'); }
    catch { setMessage('Không thể xuất Excel. Hãy thử lại.'); }
    finally { setExporting(false); }
  };
  const refreshPersonnel = async () => {
    const current = snapshot.current;
    if (!current || busy || exporting) return;
    setBusy(true); setStage('personnel'); setMessage(''); setProgress({ loaded: 0, total: 0 });
    controller.current = new AbortController();
    try {
      const tabId = Number(new URLSearchParams(location.search).get('sourceTabId'));
      if (!Number.isInteger(tabId) || tabId <= 0) throw new Error('Không xác định được tab nguồn 1Office.');
      const personnel = await loadAccessPersonnel(tabId, current.records.map((row) => row.account),
        controller.current.signal, (loaded, total) => setProgress({ loaded, total }), sourceOrigin);
      if (temporary.clearedRef.current) return;
      const allowed = filtered ? new Set(selectAccessAccounts(personnel, filters)) : null;
      snapshot.current = activeAccessSnapshot(allowed ? { ...current, records: current.records.filter((row) => allowed.has(row.account)) } : current, personnel);
      setReport(reportFromAccessSnapshot(snapshot.current, cutoff));
      setMessage('Đã cập nhật thông tin nhân sự, không quét lại nhật ký truy cập.');
    } catch (error) { snapshot.current = null; setReport(null); setMessage(error instanceof Error ? error.message : 'Chưa cập nhật được thông tin nhân sự.'); }
    finally { setBusy(false); }
  };
  if (temporary.cleared) return <TemporaryDataCleared />;
  return <main className="min-h-screen bg-slate-50 p-6 text-slate-900"><div className="mx-auto max-w-6xl space-y-5">
    <ClearTemporaryDataButton onClick={() => void temporary.clearTemporaryData()} disabled={temporary.clearing || busy || exporting} error={temporary.clearError} />
    <header className="rounded-2xl bg-gradient-to-br from-blue-800 to-violet-700 p-6 text-white">
      <TadtLogo className="mb-4 h-16 w-28 rounded bg-white p-2 object-contain" />
      <h1 className="text-2xl font-bold">Báo cáo lịch sử truy cập</h1>
      <p className="mt-2 text-sm">Nguồn: {sourceOrigin}</p>
      <p className="mt-2 text-sm text-blue-100">Truy cập cuối cùng của từng tài khoản tính đến hết ngày chốt (UTC+7), không lấy trường đăng nhập mới nhất.</p>
    </header>
    <section className="space-y-4 rounded-2xl border border-slate-200 bg-white p-6">
      <p className="text-sm text-slate-600">{filtered ? 'Đọc danh sách nhân sự trước, rồi chỉ lấy nhật ký của tài khoản khớp bộ lọc bên dưới.' : 'Đọc toàn bộ nhật ký mà tài khoản của bạn được quyền xem.'} Không dùng bộ lọc đang có trên 1Office. Giữ tab nguồn và trang này mở. Nhật ký lớn có thể mất nhiều phút; dữ liệu chỉ giữ trong bộ nhớ trang.</p>
      <CompanyScopePicker value={filters.company} disabled={busy || exporting} onChange={(value) => changeFilter('company', value)} />
      <div className="grid gap-3 md:grid-cols-2">
        {([['position', 'Lọc theo vị trí / chức danh', 'Ví dụ: Giám đốc'], ['department', 'Lọc theo phòng ban', 'Nhập tên hoặc mã phòng ban']] as const).map(([key, label, placeholder]) =>
          <label key={key} className="text-sm font-semibold">{label}<input aria-label={label} value={filters[key]} disabled={busy || exporting}
            placeholder={placeholder} onChange={(event) => changeFilter(key, event.target.value)} className="mt-1 block w-full rounded-lg border border-slate-300 p-2" />
            <span className="mt-1 block font-normal text-slate-500">{key === 'department' ? 'Nhập một phần tên hoặc mã để lọc.' : 'Nhập một phần tên để lọc.'} Để trống nếu không cần lọc.</span></label>)}
      </div>
      <p className="text-sm text-slate-500">Các điều kiện công ty, vị trí và phòng ban kết hợp đồng thời. Chọn Toàn bộ và để trống hai ô để lấy tất cả tài khoản active. Phân nhóm theo cơ cấu và mã đơn vị đã xác minh, không suy từ tên phòng ban.</p>
      <label className="block font-semibold">Ngày chốt <input aria-label="Ngày chốt" type="date" value={cutoff} disabled={busy || exporting} onChange={(event) => {
        const value = event.target.value; setCutoff(value); setReport(null); setMessage('');
        if (snapshot.current && value) {
          try { setReport(reportFromAccessSnapshot(snapshot.current, value)); setMessage('Đã tính lại ngày chốt trên dữ liệu đã thu thập, không tải lại 1Office.'); }
          catch (error) { setMessage(error instanceof Error ? error.message : 'Ngày chốt không hợp lệ.'); }
        }
      }} className="ml-3 rounded-lg border border-slate-300 p-2" /></label>
      <div className="flex gap-3"><button disabled={busy || exporting} onClick={() => void run()} className="rounded-lg bg-blue-700 px-4 py-2 text-white disabled:opacity-50">{report ? 'Thu thập lại từ 1Office' : 'Thu thập nhật ký'}</button>
        {busy && <button onClick={() => { controller.current?.abort(); setMessage('Đang hủy, chờ nhóm trang hiện tại kết thúc (tối đa khoảng 30 giây).'); }} className="rounded-lg border px-4 py-2">Hủy</button>}
        {report && <button disabled={busy || exporting} onClick={() => void refreshPersonnel()} className="rounded-lg border border-blue-300 px-4 py-2 text-blue-800 disabled:opacity-50">Cập nhật vị trí/phòng ban</button>}
        {report && <button disabled={busy || exporting} onClick={() => void download()} className="rounded-lg bg-cyan-600 px-4 py-2 text-white disabled:opacity-50">{exporting ? 'Đang xuất...' : 'Xuất Excel'}</button>}</div>
      {busy && <p role="status">{stage === 'history' ? 'Đã đối chiếu' : 'Đang bổ sung thông tin nhân sự'} {progress.loaded.toLocaleString('vi-VN')}/{progress.total.toLocaleString('vi-VN')} {stage === 'history' && !filtered ? 'bản ghi' : 'tài khoản'}. Chưa hoàn tất.</p>}
      {message && <p role="status" className="text-amber-800">{message}</p>}
      {report && <p className="text-sm text-blue-800">Có thể đổi ngày chốt mà không quét lại. Dữ liệu lấy tại lần thu thập này, không tự cập nhật khi 1Office có bản ghi mới. Thu thập lúc: {new Date(report.collectedAt).toLocaleString('vi-VN', {timeZone:'Asia/Ho_Chi_Minh'})}.</p>}
      {report?.scope && <p className="text-sm text-blue-800">{report.scope} Nhật ký và trạng thái nhân sự không phải ảnh chụp giao dịch đồng thời của máy chủ.</p>}
      <p className="text-sm text-slate-500">Báo cáo sử dụng bản ghi từ mục Lịch sử truy cập của 1Office, tổng hợp lần truy cập cuối cùng của từng tài khoản đến hết ngày chốt.</p>
      <p className="text-sm text-blue-800">Chỉ xuất tài khoản hiện còn hoạt động (active), xác minh khi thu thập/cập nhật nhân sự. Tài khoản đã khóa hoặc chưa xác minh được bị loại, dù có nhật ký cũ. Đây không phải trạng thái tài khoản tại ngày chốt. Để kiểm tra lại trạng thái sau đó, cập nhật nhân sự hoặc thu thập lại.</p>
      {report?.personnel?.warnings.map((warning) => <p key={warning} role="status" className="text-amber-800">{warning}</p>)}
    </section>
    {report && <section className="rounded-2xl border border-slate-200 bg-white p-6">
      <h2 className="text-lg font-semibold">{report.users.length} tài khoản có bản ghi đến hết {report.cutoff}</h2>
      <p className="my-3 text-sm text-slate-500">Đã đối chiếu {report.recordCount.toLocaleString('vi-VN')} bản ghi. Bản ghi cũ nhất: {report.oldestTime || 'Không có'}. Xem trước tối đa 50 tài khoản; Excel chứa toàn bộ kết quả.</p>
      <div className="overflow-x-auto"><table className="w-full text-left text-sm"><thead><tr>{['Tài khoản', 'Họ và tên tại bản ghi', 'Vị trí hiện tại', 'Phòng ban hiện tại', 'Công ty hiện tại', 'Truy cập cuối trước ngày chốt', 'Số bản ghi đến ngày chốt'].map((title) => <th key={title} className="border-b p-3">{title}</th>)}</tr></thead>
        <tbody>{report.users.slice(0, 50).map((user) => {
          const personnel = report.personnel?.users.find((item) => item.account === user.account);
          return <tr key={user.account}><td className="p-3">{user.account}</td><td className="p-3">{user.name}</td><td className="p-3">{personnel?.position || 'Chưa có dữ liệu'}</td><td className="p-3">{personnel?.department || 'Chưa có dữ liệu'}</td><td className="p-3">{personnel?.company || 'Chưa có dữ liệu'}</td><td className="p-3">{user.timeRaw}</td><td className="p-3">{user.eventCount}</td></tr>;
        })}</tbody></table></div>
      {report.users.length === 0 && <p>Không tìm thấy bản ghi trước ngày chốt trong nhật ký hiện có.</p>}
    </section>}
  </div></main>;
}
