import React, { useEffect, useRef, useState } from 'react';
import { collectTestBaseline, downloadTestBaseline, type TestBaseline } from '../oneoffice/testBaseline';
import { findTestPreservationSample, type TestPreservationProbe } from '../oneoffice/testPreservation';

export function TestBaselinePanel({ sourceTabId, onBusyChange, disabled = false }: { sourceTabId: number; onBusyChange: (busy: boolean) => void; disabled?: boolean }) {
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState('');
  const [baseline, setBaseline] = useState<TestBaseline | null>(null);
  const [probe, setProbe] = useState<TestPreservationProbe | null>(null);
  const [taskId, setTaskId] = useState('');
  const [searching, setSearching] = useState(false);
  const stop = useRef(false);
  const cancelled = useRef(false);
  const running = useRef(false);
  useEffect(() => { cancelled.current = false; return () => { cancelled.current = true; }; }, []);
  const run = async () => {
    if (running.current || disabled) return;
    running.current = true; setBusy(true); onBusyChange(true); setBaseline(null); setProbe(null); setMessage('');
    try {
      const result = await collectTestBaseline(sourceTabId, (value) => { if (!cancelled.current) setMessage(value); }, () => cancelled.current);
      if (!cancelled.current) { setBaseline(result); setMessage('Đã lập bản hiện trạng và chọn 5 mẫu. Chưa ghi dữ liệu.'); }
    } catch (error) {
      if (!cancelled.current) setMessage(error instanceof Error ? error.message : 'Không thể lập bản hiện trạng.');
    } finally { running.current = false; if (!cancelled.current) { setBusy(false); onBusyChange(false); } }
  };
  const search = async (specific: boolean) => {
    if (!baseline || running.current || disabled) return;
    running.current = true; stop.current = false; setBusy(true); setSearching(true); onBusyChange(true);
    try {
      const result = await findTestPreservationSample(sourceTabId, baseline, probe, specific ? taskId.trim() : null,
        (value) => { if (!cancelled.current) setMessage(value); }, () => cancelled.current || stop.current);
      if (!cancelled.current) { setProbe(result); setMessage(result.sample ? 'Có mẫu thảo luận trong kết quả tìm. Chưa phê duyệt ghi.' : 'Chưa tìm thấy mẫu có thảo luận trong lượt này. Không kết luận toàn dự án không có thảo luận.'); }
    } catch (error) { if (!cancelled.current) setMessage(error instanceof Error ? error.message : 'Không thể tìm mẫu.'); }
    finally { running.current = false; if (!cancelled.current) { setBusy(false); setSearching(false); onBusyChange(false); } }
  };
  const exportBaseline = baseline ? { ...baseline, preservation: probe } : null;
  return <section className="space-y-4 rounded-2xl border border-blue-200 bg-white p-5" aria-label="Chuẩn bị kiểm thử site test">
    <h2 className="text-lg font-bold">Chuẩn bị kiểm thử trên site test</h2>
    <p className="text-sm text-slate-600">Bản 4: chỉ dự án 203 tại testoffice.meygroup.vn:9443. Lập hiện trạng 5 mẫu, sau đó tìm mẫu có thảo luận để chuẩn bị đối chiếu bảo toàn. Chỉ đọc, không mở biểu mẫu sửa và không lưu lên 1Office.</p>
    <button type="button" onClick={() => void run()} disabled={disabled || busy || !Number.isSafeInteger(sourceTabId) || sourceTabId <= 0} className="rounded-xl bg-blue-700 px-5 py-3 text-sm font-semibold text-white disabled:opacity-50">{busy ? 'Đang đọc hiện trạng...' : 'Lập bản hiện trạng và chọn 5 mẫu (chỉ đọc)'}</button>
    {message && <p role="status" className="text-sm">{message}</p>}
    {baseline && <>
      <p className="text-sm font-semibold">{baseline.inventory.tasks.length} công việc · {baseline.samples.length} mẫu · 0 lần lưu</p>
      <p className="break-all text-xs text-slate-500">Nguồn: {baseline.origin} · Bắt đầu: {baseline.startedAt} · Kết thúc: {baseline.completedAt}</p>
      <p className="rounded-lg bg-amber-50 p-3 text-sm text-amber-900">Độ phủ bản nguồn: No. gốc đã xác minh {baseline.inventory.tasks.filter((task) => task.sourceNo !== undefined).length}/{baseline.inventory.tasks.length}; đọc được mã {baseline.inventory.tasks.filter((task) => !!task.code).length}/{baseline.inventory.tasks.length}. Đủ danh sách không có nghĩa đã đọc đủ các trường để cập nhật.</p>
      {baseline.inventory.collection.warnings.length > 0 && <ul className="list-inside list-disc text-sm text-amber-900">{baseline.inventory.collection.warnings.map((warning) => <li key={warning}>{warning}</li>)}</ul>}
      <div className="overflow-x-auto"><table className="w-full text-left text-sm"><thead><tr>{['ID / mã', 'Tên / lý do chọn', 'No. gốc', 'ID cha', 'Thảo luận'].map((label) => <th key={label} className="border-b p-2">{label}</th>)}</tr></thead><tbody>{baseline.samples.map((sample) => <tr key={sample.task.internalId}>
        <td className="border-b p-2">{sample.task.internalId}<br />{sample.task.code || '(chưa đọc được mã)'}</td>
        <td className="border-b p-2">{sample.task.name}<p className="text-xs text-slate-500">{sample.reason}</p></td>
        <td className="border-b p-2">{sample.task.sourceNo === undefined ? '(chưa xác minh)' : sample.task.sourceNo || '(trống)'}</td>
        <td className="border-b p-2">{sample.task.parentInternalId ?? '(không có)'}</td>
        <td className="border-b p-2">{sample.error ?? `${sample.discussion?.comments.length ?? 0} bình luận/phản hồi đã đọc`}</td>
      </tr>)}</tbody></table></div>
      <div className="space-y-2 rounded-xl border border-blue-100 p-4"><h3 className="font-semibold">Đối chiếu trực tiếp 5 mẫu (tách biệt với bản nguồn API)</h3>
        {baseline.samples.map((sample) => <p key={sample.task.internalId} className="text-sm">ID {sample.task.internalId}: mã {sample.details?.code.state === 'value' ? sample.details.code.displayValue : '(chưa xác minh)'}; No. trong phản hồi {sample.details?.no.state === 'empty' ? '(trả về trống)' : sample.details?.no.displayValue || '(chưa đọc được)'}. {sample.detailError || (!sample.details ? 'Chưa đọc chi tiết ở lượt này.' : '')}</p>)}
        <p className="text-xs text-amber-900">Kết quả chi tiết được giữ riêng, không thay thế dữ liệu nguồn ở bảng trên và chưa phê duyệt ghi. “Trả về trống” khác với “chưa đọc được”; ký hiệu -- không được coi là giá trị trống để ghi.</p>
      </div>
      {!probe?.sample && !baseline.samples.some((sample) => sample.discussion?.collection.complete && sample.discussion.comments.length) && <p className="rounded-lg bg-amber-50 p-3 text-sm text-amber-900">Chưa có mẫu xác minh được thảo luận có sẵn. Chưa đủ điều kiện kiểm thử bảo toàn thảo luận.</p>}
      <section className="space-y-3 rounded-xl border border-blue-200 bg-blue-50 p-4" aria-label="Tìm mẫu kiểm tra bảo toàn">
        <h3 className="font-semibold">Bước tiếp: tìm mẫu có thảo luận</h3>
        <p className="text-sm">Mỗi lần duyệt tối đa 20 công việc, dừng khi tìm thấy mẫu. Không tự quét hết dự án, không tạo thảo luận thử và không ghi thay đổi.</p>
        <button type="button" disabled={busy || disabled || probe?.remaining === 0} onClick={() => void search(false)} className="rounded-xl bg-blue-700 px-4 py-2 text-sm font-semibold text-white disabled:opacity-50">{probe ? 'Tìm tiếp tối đa 20 công việc (chỉ đọc)' : 'Tìm mẫu có thảo luận (tối đa 20 công việc)'}</button>
        <div className="flex flex-wrap items-end gap-2"><label className="text-sm">ID công việc test đã biết có dữ liệu<input value={taskId} onChange={(event) => setTaskId(event.target.value)} disabled={busy || disabled} inputMode="numeric" className="mt-1 block rounded-lg border p-2" /></label>
          <button type="button" disabled={busy || disabled || !taskId.trim()} onClick={() => void search(true)} className="rounded-lg border border-blue-700 px-3 py-2 text-sm text-blue-700 disabled:opacity-50">Đọc mẫu theo ID (chỉ đọc)</button></div>
        {searching && <button type="button" onClick={() => { stop.current = true; setMessage('Đang dừng sau yêu cầu đọc hiện tại...'); }} className="text-sm text-red-700">Dừng tìm mẫu</button>}
        {probe && <>
          <p className="text-sm">Đã thử đọc {probe.attempts.length} công việc · {probe.attempts.filter((item) => item.state === 'unverified').length} chưa xác minh · {probe.remaining} chưa duyệt trong bước tìm mẫu.</p>
          {probe.attempts.some((item) => item.state === 'unverified') && <details className="text-sm"><summary>ID chưa xác minh (có thể đọc lại theo ID)</summary><p className="break-all">{probe.attempts.filter((item) => item.state === 'unverified').map((item) => item.taskId).join(', ')}</p></details>}
          {probe.sample && <div className="space-y-1 rounded-lg bg-white p-3 text-sm">
            <p className="font-semibold">Mẫu ID {probe.sample.task.internalId}: {probe.sample.task.name}</p>
            <p>Thời điểm đọc thảo luận: {probe.sample.discussion?.capturedAt || '(chưa xác định)'}. Thông tin chi tiết: {probe.sample.details?.capturedAt || '(chưa xác định)'}.</p>
            <p>Thảo luận: {probe.sample.discussion?.comments.length} bình luận/phản hồi đã đọc.</p>
            <p>Định danh chi tiết: {probe.sample.details?.identityVerified ? 'đã xác minh' : 'chưa xác minh'}.</p>
            <p>Tiến độ: {probe.sample.details?.identityVerified && probe.sample.details.progress?.state === 'value' ? `${probe.sample.details.progress.percent}% (phản hồi chi tiết)` : 'chưa đọc được, không suy ra từ trạng thái công việc'}.</p>
            <p>Báo cáo thực tế: chưa thu thập. “Kết quả cần đạt được” không thay thế kết quả báo cáo.</p>
            <p className="text-amber-900">Đây chỉ là mẫu đối chiếu ban đầu, chưa chứng minh bảo toàn thảo luận, tiến độ, báo cáo hoặc toàn bộ tệp đính kèm và chưa đủ điều kiện ghi.</p>
            {probe.sample.detailError && <p className="text-red-700">{probe.sample.detailError}</p>}
          </div>}
          <p className="text-xs text-slate-600">Lượt tìm chỉ giữ trong trang. Tải bản hiện trạng JSON bên dưới để lấy kèm mẫu và danh sách ID đã thử đọc.</p>
        </>}
      </section>
      <ul className="list-inside list-disc text-sm text-amber-900">{baseline.limitations.map((item) => <li key={item}>{item}</li>)}</ul>
      <details><summary className="cursor-pointer text-sm">Xem bản hiện trạng JSON (dữ liệu nội bộ)</summary><pre className="mt-2 max-h-80 overflow-auto whitespace-pre-wrap text-xs">{JSON.stringify(exportBaseline, null, 2)}</pre></details>
      <button type="button" onClick={() => { if (exportBaseline) downloadTestBaseline(exportBaseline); }} className="rounded-xl border border-blue-700 px-4 py-2 text-sm font-semibold text-blue-700">Tải bản hiện trạng JSON về máy</button>
    </>}
  </section>;
}
