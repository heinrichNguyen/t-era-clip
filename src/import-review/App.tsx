import React, { useEffect, useState } from 'react';
import { ProjectImportReview } from '../no-review/ProjectImportReview';
import { readProjectImportContext } from '../oneoffice/projectImportWorkspace';
import { useTemporaryData } from '../oneoffice/temporaryData';
import type { OneOfficeProjectInventory } from '../oneoffice/types';
import { ClearTemporaryDataButton } from '../shared/TemporaryDataControls';
import { ProjectImportIcon } from '../shared/ProjectImportLauncher';
import { TadtLogo } from '../shared/TadtLogo';
import { NoReviewApp } from '../no-review/App';
import { TestBaselinePanel } from './TestBaselinePanel';
import { TEST_BASELINE_ORIGIN, TEST_BASELINE_PROJECT } from '../oneoffice/testBaseline';

export function ProjectImportApp() {
  const [inventory, setInventory] = useState<OneOfficeProjectInventory | null>(null);
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(true);
  const [activeTool, setActiveTool] = useState<'codes' | 'numbering'>('codes');
  const [numberingOpened, setNumberingOpened] = useState(false);
  const [numberingRunning, setNumberingRunning] = useState(false);
  const [validationOpened, setValidationOpened] = useState(false);
  const temporary = useTemporaryData(() => { setInventory(null); setError(''); setNumberingRunning(false); });
  useEffect(() => {
    let disposed = false;
    const revision = temporary.revisionRef.current;
    void readProjectImportContext(window.location.search).then((value) => {
      if (!disposed && revision === temporary.revisionRef.current) setInventory(value);
    }).catch((e) => { if (!disposed && revision === temporary.revisionRef.current) setError(e instanceof Error ? e.message : 'Chưa đọc được dữ liệu dự án.'); })
      .finally(() => { if (!disposed) setLoading(false); });
    return () => { disposed = true; };
  }, []);
  return <main className="min-h-screen bg-slate-50 text-slate-900">
    <div className="border-b border-slate-200 bg-white"><div className="mx-auto flex max-w-7xl items-center justify-between gap-4 px-5 py-4 sm:px-8">
      <div className="flex items-center gap-3"><TadtLogo className="h-10 w-20 object-contain" /><span className="border-l border-slate-200 pl-3 text-sm font-semibold text-blue-900">T-eraClip <span className="block text-xs font-normal text-slate-500">Công cụ dữ liệu dự án</span></span></div>
      <span className="rounded-full border border-blue-100 bg-blue-50 px-3 py-1 text-xs font-semibold text-blue-800">{activeTool === 'codes' ? 'Bản xem trước' : 'Công cụ No.'}</span>
    </div></div>
    <div className="mx-auto max-w-7xl space-y-6 px-5 py-6 sm:px-8">
      <header className="overflow-hidden rounded-3xl bg-gradient-to-br from-blue-900 via-indigo-800 to-violet-700 p-6 text-white shadow-xl shadow-blue-900/15 sm:p-8">
        <div className="flex flex-wrap items-start justify-between gap-6">
          <div className="min-w-0 flex-1"><p className="text-xs font-semibold uppercase tracking-widest text-blue-200">Hiện trạng / File duyệt / ID công việc</p><h1 className="mt-3 text-3xl font-bold tracking-tight sm:text-4xl">Chuẩn hóa mã dự án</h1><p className="mt-3 text-sm leading-6 text-blue-100">Đối chiếu mã và No. theo file được duyệt, giữ ID công việc hiện hữu, không xóa rồi tạo lại.</p><p className="mt-2 text-xs text-blue-200">{activeTool === 'codes' ? 'Luồng chính: ghép ID, xem mã/No. trước và sau, lập phương án. Chưa bật ghi mã và No. lên 1Office.' : 'Công cụ No. riêng trước đây, không cập nhật mã. Quyền ghi test cũ không áp dụng cho bảng đối chiếu mã + No.'}</p></div>
          <div className="hidden shrink-0 rounded-2xl border border-white/20 bg-white/10 p-4 lg:block"><ProjectImportIcon className="h-12 w-12 text-blue-100" /><p className="mt-3 text-xs text-blue-100">Mã + No. theo ID<br />Đối chiếu trước khi cập nhật</p></div>
        </div>
        {activeTool === 'codes' && <ol className="mt-7 grid gap-3 sm:grid-cols-3">{[
          ['01', 'Chọn file được duyệt', 'Mã và No. trong file là giá trị đích'],
          ['02', 'Xác nhận ID công việc', 'Ghép việc cũ hoặc việc mới đã tạo/import'],
          ['03', 'Đối chiếu mã + No.', 'Kiểm tra trước/sau, trùng và hoán đổi'],
        ].map(([number, title, help]) => <li key={number} className="flex items-start gap-3 rounded-2xl border border-white/15 bg-white/10 p-4"><span className="text-xl font-bold text-blue-200">{number}</span><span><span className="block text-sm font-semibold">{title}</span><span className="mt-1 block text-xs leading-5 text-blue-100">{help}</span></span></li>)}</ol>}
      </header>
      <div role="tablist" aria-label="Công cụ chuẩn hóa dự án" className="grid gap-3 rounded-2xl border bg-white p-2 sm:grid-cols-2">
        <button type="button" role="tab" id="codes-tab" aria-controls="codes-panel" aria-selected={activeTool === 'codes'} disabled={numberingRunning} onClick={() => setActiveTool('codes')} className={`rounded-xl px-5 py-4 text-left transition disabled:opacity-50 ${activeTool === 'codes' ? 'bg-blue-700 text-white shadow-md' : 'text-slate-700 hover:bg-slate-50'}`}><span className="block text-sm font-bold">Đối chiếu mã + No. theo file duyệt</span><span className={`mt-1 block text-xs ${activeTool === 'codes' ? 'text-blue-100' : 'text-slate-500'}`}>Luồng chính: ID ổn định, mã và No. có thể thay đổi</span></button>
        <button type="button" role="tab" id="numbering-tab" aria-controls="numbering-panel" aria-selected={activeTool === 'numbering'} disabled={numberingRunning} onClick={() => { setNumberingOpened(true); setActiveTool('numbering'); }} className={`rounded-xl px-5 py-4 text-left transition disabled:opacity-50 ${activeTool === 'numbering' ? 'bg-blue-700 text-white shadow-md' : 'text-slate-700 hover:bg-slate-50'}`}><span className="block text-sm font-bold">Công cụ No. riêng (luồng cũ)</span><span className={`mt-1 block text-xs ${activeTool === 'numbering' ? 'text-blue-100' : 'text-slate-500'}`}>Chỉ đổi No. trong phạm vi test cũ, không dùng để đồng bộ mã + No.</span></button>
      </div>
      <section className="flex flex-wrap items-center justify-between gap-3 rounded-2xl border bg-white px-5 py-3" aria-label="Phạm vi đối chiếu">
        {inventory && <p className="w-full break-all text-xs text-slate-600">Nguồn: {new URL(inventory.project.url).origin}</p>}
        <div><p className="text-sm font-semibold">{inventory ? `Đối chiếu với: ${inventory.project.name}` : 'Kiểm tra file trên máy'}</p><p className="mt-1 text-xs text-slate-600">{inventory ? 'Dùng danh sách đã thu thập, không phải dữ liệu đọc trực tiếp lúc này. Kiểm tra đúng dự án trước khi đối chiếu.' : 'Không cần mở dự án để thử lập mã lần đầu. Muốn đối chiếu việc cũ, mở công cụ từ danh sách dự án đã thu thập.'}</p></div>
        <ClearTemporaryDataButton onClick={() => void temporary.clearTemporaryData()} disabled={temporary.clearing || loading || numberingRunning} error={temporary.clearError} />
      </section>
      {error && <p role="alert" className="rounded-xl border border-amber-200 bg-amber-50 p-4 text-sm text-amber-900">{error}</p>}
      {loading ? <p role="status">Đang chuẩn bị công cụ...</p> : <>
        <div role="tabpanel" id="codes-panel" aria-labelledby="codes-tab" hidden={activeTool !== 'codes'}><ProjectImportReview key={temporary.revisionRef.current} inventory={inventory} primaryReconciliation sourceTabId={Number(new URLSearchParams(window.location.search).get('sourceTabId'))} onBusyChange={setNumberingRunning} disabled={temporary.clearing || numberingRunning} /></div>
        <div role="tabpanel" id="numbering-panel" aria-labelledby="numbering-tab" hidden={activeTool !== 'numbering'}>{numberingOpened && <NoReviewApp embedded inventoryContext={inventory} onRunningChange={setNumberingRunning} />}</div>
        {inventory?.project.internalId === TEST_BASELINE_PROJECT && new URL(inventory.project.url).origin === TEST_BASELINE_ORIGIN && <details className="rounded-2xl border bg-white p-4" onToggle={(event) => setValidationOpened(event.currentTarget.open)}><summary onClick={(event) => { if (numberingRunning) event.preventDefault(); }} className="cursor-pointer font-semibold">Kiểm thử kỹ thuật và bảo toàn dữ liệu (phụ, không bắt buộc để lập bảng đối chiếu)</summary>{validationOpened && <TestBaselinePanel key={temporary.revisionRef.current} sourceTabId={Number(new URLSearchParams(window.location.search).get('sourceTabId'))} onBusyChange={setNumberingRunning} disabled={numberingRunning} />}</details>}
      </>}
    </div>
  </main>;
}
