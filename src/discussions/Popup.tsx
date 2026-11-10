import React, { useEffect, useRef, useState } from 'react';
import { collectOneOfficeDiscussion, collectOneOfficeProjectInventory, getActiveTab, isOneOfficeProjectPageUrl, isOneOfficeReportPageUrl, isSupportedOneOfficeUrl, type ActiveTab } from '../oneoffice/collectionClient';
import { collectOneOfficeReportInventory } from '../oneoffice/reportTaskInventory';
import { openOneOfficeReview } from '../oneoffice/reviewSession';
import { openOneOfficeProjectBatchReview } from '../oneoffice/projectBatchSession';
import type { OneOfficeCollectionProgress } from '../oneoffice/collector';
import type { OneOfficeDiscussionExport, OneOfficeProjectInventory } from '../oneoffice/types';
import { useTemporaryData } from '../oneoffice/temporaryData';
import { ClearTemporaryDataButton } from '../shared/TemporaryDataControls';
import { TadtLogo } from '../shared/TadtLogo';
import { projectInventoryWarningLabel } from '../shared/projectInventoryWarning';

export function DiscussionPopup() {
  const [tab, setTab] = useState<ActiveTab | null>(null);
  const [busy, setBusy] = useState(false); const running = useRef(false); const disposed = useRef(false);
  const [message, setMessage] = useState(''); const [error, setError] = useState('');
  const [result, setResult] = useState<OneOfficeDiscussionExport | null>(null);
  const [inventory, setInventory] = useState<OneOfficeProjectInventory | null>(null);
  const temporary = useTemporaryData(() => { setResult(null); setInventory(null); setMessage(''); });
  useEffect(() => {
    disposed.current = false;
    void getActiveTab().then((value) => { if (!disposed.current && value && isSupportedOneOfficeUrl(value.url)) setTab(value); })
      .catch(() => { if (!disposed.current) setError('Không đọc được tab nguồn. Mở lại tiện ích trên trang 1Office.'); });
    return () => { disposed.current = true; };
  }, []);
  const run = async (project: boolean, report = false) => {
    if (!tab || running.current) return;
    running.current = true; setBusy(true); setError(''); setMessage('Đang chuẩn bị...'); temporary.reset();
    setResult(null); setInventory(null);
    const revision = temporary.revisionRef.current;
    const current = () => !disposed.current && revision === temporary.revisionRef.current;
    try {
      if (project || report) {
        if (report ? !isOneOfficeReportPageUrl(tab.url) : !isOneOfficeProjectPageUrl(tab.url)) throw new Error('Mở đúng trang báo cáo hoặc dự án 1Office trước khi thu thập.');
        setMessage(report ? 'Đang đọc danh sách theo bộ lọc báo cáo...' : 'Đang kiểm tra danh sách công việc...');
        const value = report ? await collectOneOfficeReportInventory(tab.id, new URL(tab.url).origin, tab.url)
          : await collectOneOfficeProjectInventory(tab.id, new URL(tab.url).origin);
        if (!current()) return;
        setInventory(value);
        if (value.collection.complete && value.tasks.length) {
          setMessage('Đang mở trang thu thập thảo luận danh sách công việc...');
          await openOneOfficeProjectBatchReview(tab, value);
        } else setMessage('Danh sách chưa đầy đủ hoặc chưa có công việc. Chưa bắt đầu thu thập toàn dự án.');
      } else {
        const value = await collectOneOfficeDiscussion(tab.id, (progress: OneOfficeCollectionProgress) => {
          if (current()) setMessage(`Đang tải thảo luận: ${progress.loadedRootCommentCount} bình luận gốc...`);
        }, new URL(tab.url).origin);
        if (!current()) return;
        setResult(value); setMessage('Đã thu thập thảo luận trang hiện tại. Kiểm tra kết quả rồi mở bản xem trước.');
      }
    } catch (e) { if (current()) setError(e instanceof Error ? e.message : 'Không thu thập được thảo luận.'); }
    finally { running.current = false; if (!disposed.current) setBusy(false); }
  };
  const review = async () => {
    if (!result || running.current) return;
    running.current = true; setBusy(true); setError('');
    try { await openOneOfficeReview(result); }
    catch (e) { if (!disposed.current) setError(e instanceof Error ? e.message : 'Không mở được bản xem trước.'); }
    finally { running.current = false; if (!disposed.current) setBusy(false); }
  };
  return <main className="w-[420px] space-y-4 bg-slate-50 p-5 text-slate-900">
    <header className="rounded-2xl bg-gradient-to-br from-blue-900 to-violet-700 p-5 text-white"><TadtLogo className="h-10 w-20 rounded bg-white p-1" /><h1 className="mt-4 text-xl font-bold">T-eraClip Thảo luận</h1><p className="mt-2 text-sm text-blue-100">Thu thập bình luận và phản hồi 1Office để xem trước và xuất file. Không chỉnh sửa dữ liệu công việc.</p></header>
    <section className="space-y-3 rounded-2xl border bg-white p-4">
      <h2 className="font-semibold">Chọn phạm vi thu thập</h2>
      <p className="break-all text-xs text-slate-500">{tab ? `Nguồn: ${tab.title || new URL(tab.url).origin}` : 'Mở công việc hoặc dự án trên 1Office, đăng nhập rồi mở lại tiện ích.'}</p>
      <button type="button" disabled={!tab || busy || isOneOfficeReportPageUrl(tab.url)} onClick={() => void run(false)} className="w-full rounded-xl bg-blue-700 px-4 py-3 text-sm font-semibold text-white disabled:opacity-50">Thu thập thảo luận trang hiện tại</button>
      <p className="text-xs text-slate-600">Dùng cho một công việc đang mở. Trên trang dự án, nút này chỉ lấy thảo luận cấp dự án, không gồm từng công việc.</p>
      <button type="button" disabled={!tab || !isOneOfficeProjectPageUrl(tab.url) || busy} onClick={() => void run(true)} className="w-full rounded-xl border border-blue-700 px-4 py-3 text-sm font-semibold text-blue-700 disabled:opacity-50">Thu thập thảo luận toàn dự án</button>
      <p className="text-xs text-slate-600">Kiểm tra danh sách → thu thập từng công việc trên trang riêng → lọc và xuất Excel.</p>
      {tab && isOneOfficeReportPageUrl(tab.url) && <>
        <button type="button" disabled={busy} onClick={() => void run(false, true)} className="w-full rounded-xl bg-blue-700 px-4 py-3 text-sm font-semibold text-white disabled:opacity-50">Thu thập thảo luận danh sách trong báo cáo</button>
        <p className="text-xs text-slate-600">Lọc phòng ban và thời gian trên báo cáo trước. Tiến độ mở ở trang riêng và thu thập bằng request ngầm, không mở từng công việc. Ngày báo cáo không tự lọc ngày bình luận. Luồng xem trước, lọc và xuất Excel giữ nguyên.</p>
      </>}
    </section>
    {message && <p role="status" className="text-sm text-blue-800">{message}</p>}
    {error && <p role="alert" className="text-sm text-red-700">{error}</p>}
    {result && <section className="space-y-2 rounded-xl border bg-white p-4 text-sm"><p className="font-semibold">{result.collection.loadedRootCommentCount} bình luận gốc · {result.collection.loadedReplyCount} phản hồi</p><p>{result.collection.complete ? 'Đã tải đủ theo bộ thu thập.' : 'Chưa đầy đủ, cần kiểm tra cảnh báo.'}</p>
      {result.collection.warnings.length > 0 && <details><summary>Cảnh báo thu thập</summary><ul className="list-inside list-disc text-xs text-amber-800">{result.collection.warnings.map((warning, i) => <li key={i}>{warning}</li>)}</ul></details>}
      <button type="button" disabled={busy} onClick={() => void review()} className="w-full rounded-xl bg-blue-700 px-4 py-3 font-semibold text-white disabled:opacity-50">Xem trước và xuất file</button></section>}
    {inventory && <p className="text-sm">Danh sách: {inventory.tasks.length} công việc. {inventory.collection.warnings.map(projectInventoryWarningLabel).join(' ')}</p>}
    <p className="text-xs text-slate-600">Chỉ dùng tài khoản được phân quyền. Không đăng nhập hộ hoặc đọc mật khẩu/cookie. File xuất có thể chứa thông tin nội bộ, chỉ chia sẻ đúng phạm vi.</p>
    <ClearTemporaryDataButton onClick={() => void temporary.clearTemporaryData()} disabled={busy || temporary.clearing} error={temporary.clearError} />
  </main>;
}
