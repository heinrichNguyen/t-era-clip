import React, { useEffect, useState } from 'react';
import { buildNoUpdatePlan } from '../oneoffice/noUpdatePlan';
import { readNoReview } from '../oneoffice/noReviewSession';
import { useTemporaryData } from '../oneoffice/temporaryData';
import type { OneOfficeProjectInventory } from '../oneoffice/types';
import { TadtLogo } from '../shared/TadtLogo';
import { inspectNoWriteJournal, probeDirectNoForms, runNoUpdatePlan, runNoWriteExperiment, type NoExperimentResult } from '../oneoffice/noWriteExperiment';
import { clearResolvedNoWriteJournal, readNoWriteJournal, type NoWriteJournal } from '../oneoffice/noWriteJournal';
import { isNoTestTask } from '../oneoffice/noTestScope';
import { ProjectImportReview } from './ProjectImportReview';
import { ProjectImportLauncher } from '../shared/ProjectImportLauncher';
import { ONEOFFICE_ALLOWED_ORIGIN } from '../oneoffice/origins';
import { isSupportedOneOfficeUrl } from '../oneoffice/collectionClient';

export function NoReviewApp({ embedded = false, inventoryContext = null, onRunningChange }: {
  embedded?: boolean; inventoryContext?: OneOfficeProjectInventory | null; onRunningChange?: (running: boolean) => void;
} = {}) {
  const [inventory, setInventory] = useState<OneOfficeProjectInventory | null>(inventoryContext);
  const writableEnvironment = !!inventory && isSupportedOneOfficeUrl(inventory.project.url)
    && new URL(inventory.project.url).origin === ONEOFFICE_ALLOWED_ORIGIN;
  const [text, setText] = useState('');
  const [plan, setPlan] = useState<ReturnType<typeof buildNoUpdatePlan> | null>(null);
  const [error, setError] = useState('');
  const [running, setRunning] = useState(false);
  useEffect(() => { onRunningChange?.(running); }, [running, onRunningChange]);
  const [experiment, setExperiment] = useState<NoExperimentResult | null>(null);
  const [planResult, setPlanResult] = useState<NoExperimentResult | null>(null);
  const [writeBlocked, setWriteBlocked] = useState(false);
  const [journal, setJournal] = useState<NoWriteJournal | null>(null);
  const [journalReady, setJournalReady] = useState(false);
  useEffect(() => {
    let disposed = false;
    void readNoWriteJournal().then((value) => {
      if (disposed) return;
      setJournal(value); setJournalReady(true);
      if (value?.status === 'pending') setWriteBlocked(true);
    }).catch((e) => { if (!disposed) { setWriteBlocked(true); setError(e.message); } });
    return () => { disposed = true; };
  }, []);
  const temporary = useTemporaryData(() => { setInventory(null); setText(''); setPlan(null); setExperiment(null); setPlanResult(null); });
  useEffect(() => {
    if (!running) return;
    const warn = (event: BeforeUnloadEvent) => { event.preventDefault(); event.returnValue = ''; };
    window.addEventListener('beforeunload', warn);
    return () => window.removeEventListener('beforeunload', warn);
  }, [running]);
  useEffect(() => {
    if (embedded) { setInventory(inventoryContext); return; }
    let disposed = false;
    const revision = temporary.revisionRef.current;
    void readNoReview().then((value) => {
      if (!disposed && revision === temporary.revisionRef.current) setInventory(value);
    }).catch((e) => { if (!disposed) setError(e instanceof Error ? e.message : 'Không thể đọc dữ liệu tạm.'); });
    return () => { disposed = true; };
  }, [embedded, inventoryContext]);
  const handleExperiment = (readOnly: boolean) => {
    if (running || (!readOnly && (writeBlocked || !journalReady))) return;
    if (!readOnly && !window.confirm('Chỉ trên dự án test ID 870: đổi No. của TEST01.3 và TEST01.INSERT, kiểm tra rồi hoàn nguyên. Thao tác sẽ gửi yêu cầu lưu lên 1Office. Tiếp tục?')) return;
    setRunning(true); setError(''); setExperiment(null);
    const revision = temporary.revisionRef.current;
    const tabId = Number(new URLSearchParams(window.location.search).get('sourceTabId'));
    void (readOnly ? probeDirectNoForms(tabId) : runNoWriteExperiment(tabId, false))
      .then((value) => { if (revision === temporary.revisionRef.current) { setExperiment(value); setPlan(null); if (!value.complete && value.saveAttempts > 0) setWriteBlocked(true); } })
      .catch((e) => { if (!readOnly) setWriteBlocked(true); setError(e instanceof Error ? e.message : 'Không nhận được kết quả thử nghiệm.'); })
      .finally(() => setRunning(false));
  };
  const allowedPlan = writableEnvironment && plan && plan.projectId === '870' && !plan.errors.length && plan.changes.length > 0
    && plan.changes.length <= 5 && plan.changes.every(isNoTestTask);
  const handleWritePlan = () => {
    if (running || writeBlocked || !journalReady || !allowedPlan || !inventory || !plan) return;
    if (!window.confirm(`Ghi ${plan.changes.length} thay đổi No. trên dự án test ID 870 theo bảng xem trước? Chỉ dùng số tạm khi cần phá vòng trùng No. Nhật ký lưu cục bộ trên máy. Kết quả được giữ lại, KHÔNG tự hoàn nguyên. Việc lưu tạo lịch sử cập nhật; nếu lỗi có thể chỉ cập nhật một phần. Tiếp tục?`)) return;
    setRunning(true); setError(''); setPlanResult(null);
    const revision = temporary.revisionRef.current;
    void runNoUpdatePlan(Number(new URLSearchParams(window.location.search).get('sourceTabId')), inventory, text)
      .then((value) => {
        if (revision !== temporary.revisionRef.current) return;
        setPlanResult(value);
        // Require a fresh inventory after any submitted plan, even a successful one.
        if (value.saveAttempts > 0) setWriteBlocked(true);
      })
      .catch((e) => { setWriteBlocked(true); setError(e instanceof Error ? e.message : 'Không nhận được kết quả ghi. Kiểm tra trạng thái, không chạy lại.'); })
      .finally(() => {
        setRunning(false);
        void readNoWriteJournal().then(setJournal).catch((e) => { setWriteBlocked(true); setError(e.message); });
      });
  };
  const handleInspectJournal = () => {
    if (running) return;
    setRunning(true); setError('');
    void inspectNoWriteJournal(Number(new URLSearchParams(window.location.search).get('sourceTabId')))
      .then((value) => { setJournal(value); setPlan(null); setPlanResult(null); setWriteBlocked(true); })
      .catch((e) => setError(e instanceof Error ? e.message : 'Chưa kiểm tra được trạng thái. Giữ nguyên nhật ký.'))
      .finally(() => setRunning(false));
  };
  const Container = embedded ? 'section' : 'main';
  const Heading = embedded ? 'h2' : 'h1';
  return <Container className={embedded ? 'text-slate-900' : 'min-h-screen bg-slate-50 p-6 pb-32 text-slate-900'}>
    <div className={embedded ? 'space-y-5' : 'mx-auto max-w-6xl space-y-5'}>
      <header className="rounded-2xl bg-gradient-to-r from-blue-800 to-violet-700 p-6 text-white">
        {!embedded && <TadtLogo className="mb-5 h-16 w-28 rounded-lg bg-white p-2 object-contain" />}
        <Heading className="text-2xl font-bold">Đối chiếu No. công việc hàng loạt</Heading>
        <p className="mt-2">{inventory?.project.name || 'Chưa có danh sách dự án'}</p>
        <p className="mt-2 text-sm">Xem trước không ghi dữ liệu. Chỉ dự án test ID 870 được ghi tối đa năm mã công việc cho phép, sau khi bạn xác nhận. Không tạo, xóa hoặc đổi mã công việc.</p>
      </header>
      {error && <p role="alert">{error}</p>}
      {journal && <section className="rounded-2xl border bg-white p-6">
        <h2 className="font-semibold">Nhật ký lượt ghi gần nhất (trên máy này)</h2>
        <p className="mt-2 text-sm">{journal.status === 'pending' ? 'Chưa đối soát, khóa lượt ghi mới.' : journal.status === 'verified' ? 'Đã xác nhận toàn bộ No. cuối cùng.' : 'Đã đọc lại: toàn bộ No. vẫn ở giá trị trước lượt ghi.'} Cập nhật: {journal.updatedAt}.</p>
        <p className="mt-2 text-sm">Nhật ký vẫn còn khi đóng trang hoặc khởi động lại Chrome. Không tự ghi tiếp hoặc hoàn nguyên. “Xóa dữ liệu tạm” không xóa nhật ký phục hồi này.</p>
        {journal.rows.map((row) => <p className="mt-2 text-sm" key={row.internalId}>{row.code}: {row.before} → {row.after}. {row.observed === null ? 'Chưa có lần đọc lại xác nhận.' : `Lần xác nhận gần nhất: ${row.observed || '(trống)'}.`} {row.state === 'temporary' ? 'Đang ở No. tạm.' : row.state === 'conflict' ? 'Giá trị khác dự kiến, cần kiểm tra thủ công.' : ''}</p>)}
        <button type="button" disabled={running} className="mt-3 rounded-lg border border-blue-700 px-4 py-3 text-blue-700" onClick={handleInspectJournal}>Kiểm tra trạng thái trên 1Office (không ghi)</button>
        {journal.status !== 'pending' && <button type="button" disabled={running} className="ml-4 text-red-700" onClick={() => {
          if (!window.confirm('Xóa nhật ký đã đối soát trên máy này? Không thay đổi dữ liệu 1Office.')) return;
          void clearResolvedNoWriteJournal().then(() => setJournal(null)).catch((e) => setError(e.message));
        }}>Xóa nhật ký đã đối soát</button>}
      </section>}
      {!inventory && <p>Hãy mở dự án, kiểm tra danh sách công việc rồi mở “Chuẩn hóa mã dự án” để dùng danh sách đó cho phần đối chiếu No.</p>}
      {inventory && <section className="rounded-2xl border bg-white p-6">
        <p className="mb-4 text-sm text-slate-600">Đối chiếu theo mã công việc, phân biệt chữ hoa/thường. Công việc không có trong dữ liệu dán được giữ nguyên. No. không được dùng để suy ra hoặc thay đổi quan hệ cha/con.</p>
        <label className="block font-semibold" htmlFor="numbering">Dán hai cột từ Excel: Mã công việc, No. mới (không kèm tiêu đề)</label>
        <textarea id="numbering" className="mt-3 min-h-44 w-full rounded-lg border p-3 font-mono" value={text}
          placeholder={'TEST01.3\t1.4\nTEST01.INSERT\t1.3'}
          disabled={running} maxLength={1000000} onChange={(e) => { setText(e.target.value); setPlan(null); }} />
        <button type="button" disabled={running} className="mt-3 rounded-lg bg-blue-700 px-5 py-3 font-semibold text-white"
          onClick={() => setPlan(buildNoUpdatePlan(inventory, text))}>Xem trước thay đổi</button>
        <button type="button" disabled={running} className="ml-4 text-sm text-red-700" onClick={() => void temporary.clearTemporaryData()}>Xóa dữ liệu tạm</button>
      </section>}
      {temporary.clearError && <p role="alert">{temporary.clearError}</p>}
      {!embedded && <ProjectImportReview key={temporary.revisionRef.current} inventory={inventory} disabled={running} />}
      {writableEnvironment && inventory?.project.internalId === '870' && <section className="rounded-2xl border border-amber-300 bg-white p-6">
        <h2 className="text-lg font-semibold">Kiểm tra biểu mẫu trên dự án test</h2>
        <p className="mt-1 text-xs text-slate-500">Bản 14: kiểm tra mở trực tiếp năm biểu mẫu, chỉ đọc, không lưu</p>
        <p className="mt-2 text-sm">TEST01.01, TEST01.2, TEST01.3, TEST01.INSERT và TEST01.4. Không dùng dữ liệu dán ở trên. Muốn ghi thay đổi, dùng bảng xem trước có nhật ký phục hồi bên dưới. Chưa hỗ trợ ghi dự án thật.</p>
        <p className="mt-2 text-sm">Khi bắt đầu, extension chuyển sang tab dự án test. Giữ nguyên tab đó và không thu nhỏ Chrome. Khi các hộp thoại đã dừng mở/đóng, quay lại trang này để đọc kết quả.</p>
        <button type="button" disabled={running} className="mr-3 mt-4 rounded-lg border border-blue-700 px-4 py-3 font-semibold text-blue-700 disabled:opacity-50"
          onClick={() => handleExperiment(true)}>Kiểm tra mở biểu mẫu (không ghi)</button>
        {experiment && <div className="mt-4" role="status">
          <p className="font-semibold">{experiment.complete
            ? experiment.readOnly ? `Đã mở và đọc được ${experiment.tasks.length} biểu mẫu. Chưa ghi dữ liệu.` : 'Đã ghi, đọc lại và hoàn nguyên cả 2 công việc.'
            : 'Thử nghiệm chưa hoàn tất. Kiểm tra trạng thái trước khi chạy lại.'}</p>
          <p className="mt-2 text-sm">Bước cuối: {experiment.phase}. Số lần kích hoạt lưu: {experiment.saveAttempts}.</p>
          {experiment.saveAttempts === 0 && <p className="text-sm">Runner chưa kích hoạt lưu công việc nào.</p>}
          {experiment.error && <p className="text-red-700">{experiment.error}</p>}
          {!!experiment.diagnostics?.length && <div className="mt-3 rounded-lg bg-slate-50 p-3 text-sm">
            <p className="font-semibold">Chẩn đoán mở biểu mẫu (bản 14)</p>
            <p>Chỉ ghi nhận giao diện, không xác nhận yêu cầu lưu. Không tự mở dropdown.</p>
            {experiment.diagnostics.map((line, index) => <p className="mt-2" key={index}>{line}</p>)}
          </div>}
          {experiment.tasks.map((row) => <p className="mt-2 text-sm" key={row.code}>{experiment.readOnly
            ? `${row.code}: đã đọc No. ${row.before}, không kích hoạt lưu.`
            : `${row.code}: ${row.before} → ${row.temporary} → ${row.before}. Ghi: ${row.applied ? 'đã xác nhận' : 'chưa xác nhận'}; hoàn nguyên: ${row.restored ? 'đã xác nhận' : 'chưa xác nhận'}. ${row.error || ''}`}</p>)}
        </div>}
      </section>}
      {plan && <section className="rounded-2xl border bg-white p-6">
        <h2 className="text-lg font-semibold">{plan.changes.length} thay đổi, {plan.unchanged} không đổi, {plan.untouched} công việc giữ nguyên ngoài phạm vi dán</h2>
        <p className="mt-2 text-sm text-slate-600">Danh sách nguồn lúc {plan.capturedAt}. Đây là ảnh chụp dữ liệu, chưa xác nhận trạng thái hiện tại để ghi.</p>
        {plan.errors.length > 0 && <ul role="alert" className="my-4 list-inside list-disc text-amber-800">{plan.errors.map((e) => <li key={e}>{e}</li>)}</ul>}
        <div className="mt-4 overflow-auto"><table className="w-full text-left text-sm">
          <thead><tr>{['Mã công việc', 'Tên công việc', 'No. hiện tại', 'No. mới', 'Kết quả'].map((label) => <th className="border-b p-3" key={label}>{label}</th>)}</tr></thead>
          <tbody>{plan.rows.map((row) => <tr key={row.internalId}>
            <td className="border-b p-3">{row.code}</td><td className="border-b p-3">{row.name}</td>
            <td className="border-b p-3">{row.before || '(trống)'}</td><td className="border-b p-3">{row.after}</td>
            <td className="border-b p-3">{row.before === row.after ? 'Không đổi' : planResult?.tasks.find((task) => task.code === row.code && task.after === row.after)?.finalized ? 'Đã đọc lại xác nhận' : 'Dự kiến cập nhật'}</td>
          </tr>)}</tbody>
        </table></div>
        {!writableEnvironment || plan.projectId !== '870' ? <p className="mt-5 rounded-lg bg-amber-50 p-4 text-sm text-amber-900">Chưa bật ghi cho dự án và môi trường này.</p> : <div className="mt-5 rounded-lg bg-amber-50 p-4 text-sm text-amber-900">
          <p>Ghi theo danh sách (bản 14): cho phép TEST01.01, TEST01.2, TEST01.3, TEST01.INSERT và TEST01.4 trong dự án test. Kiểm tra hiện trạng và kết quả bằng request ngầm, mở trực tiếp biểu mẫu lưu. Dùng No. gốc, không dùng số thứ tự tự sinh thay ô trống. Có nhật ký phục hồi cục bộ. Giữ nguyên mã và quan hệ cha/con. Không tự hoàn nguyên.</p>
          <p className="mt-2">Sau lưu, đối chiếu No. và các trường dữ liệu API trả về, không khẳng định API bao phủ toàn bộ biểu mẫu. Phản hồi thiếu hoặc khác dự kiến sẽ dừng lượt ghi để kiểm tra.</p>
          <p className="mt-2">Tab nguồn sẽ tự mở lên. Giữ tab đó hiển thị, không thu nhỏ Chrome hoặc sửa công việc trong lúc chạy.</p>
          <button type="button" disabled={!allowedPlan || running || writeBlocked || !journalReady} onClick={handleWritePlan}
            className="mt-3 rounded-lg bg-blue-700 px-5 py-3 font-semibold text-white disabled:opacity-50">Ghi No. theo bảng xem trước (dự án test)</button>
          {!allowedPlan && <p className="mt-2">Cần bản xem trước không lỗi, có thay đổi và chỉ thuộc năm mã công việc test cho phép.</p>}
        </div>}
      </section>}
      {writeBlocked && <p role="alert" className="rounded-lg bg-amber-50 p-4 text-amber-900">Đã khóa ghi tiếp trong trang này. Kiểm tra No. trên 1Office, thu thập lại danh sách dự án và mở bản đối chiếu mới trước khi lập lượt ghi khác. Không dùng lại bản xem trước cũ.</p>}
      {planResult && <section role="status" className="rounded-2xl border bg-white p-6">
        <h2 className="font-semibold">{planResult.complete ? 'Đã ghi và đọc lại xác nhận toàn bộ No. mới.' : 'Chưa hoàn tất ghi theo danh sách. Kiểm tra trạng thái, không chạy lại.'}</h2>
        <p className="mt-2">{planResult.phase}. Số lần kích hoạt lưu: {planResult.saveAttempts}.</p>
        {planResult.metrics && <p className="mt-2 text-sm">{planResult.metrics.formLoads} lượt tải biểu mẫu; {planResult.metrics.readRequests ?? 0} request đọc ngầm; {(planResult.metrics.durationMs / 1000).toFixed(1)} giây. Chỉ là thời gian lượt chạy này, chưa phải tốc độ cho dự án lớn.</p>}
        {planResult.error && <p className="mt-2 text-red-700">{planResult.error}</p>}
        {planResult.tasks.map((row) => <p className="mt-2" key={row.code}>{row.code}: {row.before} → {row.after}. {row.finalized ? 'Đã đọc lại xác nhận No. cuối cùng.' : row.applied ? `Đã xác nhận No. tạm ${row.temporary}, chưa xác nhận No. cuối cùng.` : 'Chưa xác nhận thay đổi, kiểm tra thủ công nếu đã gửi lưu.'}</p>)}
      </section>}
    </div>
    {!embedded && <ProjectImportLauncher floating inventory={inventory} sourceTabId={Number(new URLSearchParams(window.location.search).get('sourceTabId'))} disabled={running} />}
  </Container>;
}
