import React, { useState } from 'react';
import type { ImportReview } from '../oneoffice/projectImportReview';
import { proposeImportNormalization, type NormalizationOptions } from '../oneoffice/projectImportNormalization';
import type { OneOfficeProjectInventory } from '../oneoffice/types';
import { ProjectImportDiff } from './ProjectImportDiff';
import { ProjectErpCatalogReview } from './ProjectErpCatalogReview';
import type { ProjectErpCatalog } from '../oneoffice/projectErpCatalog';

const stateLabels = { excluded: 'Không lập mã ở bước này', blocked: 'Cần kiểm tra', existing: 'Giữ mã trên 1Office', proposed: 'Đã có đề xuất mã' };
const outlineExample = (level: number) => Array.from({ length: level }, () => '1').join('.');

export function ProjectImportNormalization({ report, inventory, disabled }: { report: ImportReview; inventory: OneOfficeProjectInventory | null; disabled: boolean }) {
  const [options, setOptions] = useState<NormalizationOptions>({ mode: 'update', prefix: '', firstTaskLevel: 3, packageLevel: 4, mappings: '', newRows: '' });
  const [show, setShow] = useState(false);
  const [page, setPage] = useState(0);
  const [catalog, setCatalog] = useState<ProjectErpCatalog | null>(null);
  const [catalogLoading, setCatalogLoading] = useState(false);
  const change = (patch: Partial<NormalizationOptions>) => { setOptions((previous) => ({ ...previous, ...patch, ...('packageLevel' in patch || 'firstTaskLevel' in patch || 'mode' in patch ? { erpPackages: {} } : {}) })); setShow(false); setPage(0); };
  const result = show ? proposeImportNormalization(report, inventory, options, catalog) : null;
  const lastPage = Math.max(0, Math.ceil((result?.rows.length || 0) / 50) - 1);
  const counts = result ? {
    existing: result.rows.filter((r) => r.kind === 'existing').length,
    proposed: result.rows.filter((r) => r.kind === 'proposed').length,
    blocked: result.rows.filter((r) => r.kind === 'blocked').length,
    excluded: result.rows.filter((r) => r.kind === 'excluded').length,
  } : null;
  return <div className="mt-6 rounded-xl border border-blue-200 bg-blue-50 p-4">
    <div className="flex flex-wrap items-center justify-between gap-2">
      <h3 className="text-lg font-semibold">Bước 2: Xem trước mã công việc</h3>
      <span className="rounded-full bg-white px-3 py-1 text-xs font-semibold text-blue-800">Chỉ xem trước, chưa cập nhật 1Office</span>
    </div>
    <p className="mt-2 text-sm text-slate-700">Chọn cách lập mã và cấp công việc để xem bảng mã dự kiến. Mã trong file gốc và dữ liệu trên 1Office chưa bị thay đổi.</p>
    <fieldset disabled={disabled} className="mt-4 grid gap-4 md:grid-cols-2">
      <div>
        <label htmlFor="normalization-mode" className="text-sm font-semibold">Bạn muốn lập mã cho trường hợp nào?</label>
        <select id="normalization-mode" aria-describedby="normalization-mode-help" className="mt-1 w-full rounded border bg-white p-2" value={options.mode} onChange={(e) => change({ mode: e.target.value as NormalizationOptions['mode'] })}>
          <option value="update">Danh sách đã có trên 1Office: giữ mã cũ</option><option value="initial">Danh sách mới: lập mã lần đầu</option>
        </select>
        <p id="normalization-mode-help" className="mt-1 text-xs text-slate-600">{options.mode === 'update' ? 'Đối chiếu mã với danh sách 1Office đã thu thập. Không tự đánh lại mã công việc cũ.' : 'Chỉ chọn khi danh sách chưa được lập mã trên 1Office. Mã cấp con được đề xuất theo thứ tự trong file.'}</p>
      </div>
      <div>
        <label htmlFor="normalization-prefix" className="text-sm font-semibold">Phần đầu dùng chung cho mã dự án</label>
        <input id="normalization-prefix" aria-describedby="normalization-prefix-help" className="mt-1 w-full rounded border p-2" placeholder="Ví dụ: CNHY hoặc HP.HH1" maxLength={160} value={options.prefix} onChange={(e) => change({ prefix: e.target.value })} />
        <p id="normalization-prefix-help" className="mt-1 text-xs text-slate-600">Chỉ nhập mã dự án. Mã mới sẽ ghép đủ: Mã dự án.Hạng mục.Gói thầu.Cha.Con.Cháu.Chắt… Phần đầu này không dùng để xác định dự án trên 1Office.</p>
      </div>
      <div>
        <label htmlFor="normalization-item-level" className="text-sm font-semibold">Hạng mục nằm ở cấp nào trong file?</label>
        <select id="normalization-item-level" aria-describedby="normalization-level-help" className="mt-1 w-full rounded border bg-white p-2" value={options.firstTaskLevel} onChange={(e) => change({ firstTaskLevel: Number(e.target.value) })}>{[1,2,3,4,5,6,7,8].map((n) => <option key={n} value={n}>Cấp {n} (ví dụ No. {outlineExample(n)})</option>)}</select>
      </div>
      <div>
        <label htmlFor="normalization-package-level" className="text-sm font-semibold">Gói thầu hoặc công việc chính nằm ở cấp nào?</label>
        <select id="normalization-package-level" aria-describedby="normalization-level-help" className="mt-1 w-full rounded border bg-white p-2" value={options.packageLevel} onChange={(e) => change({ packageLevel: Number(e.target.value) })}>{[2,3,4,5,6,7,8,9].map((n) => <option key={n} value={n}>Cấp {n} (ví dụ No. {outlineExample(n)})</option>)}</select>
      </div>
      <div id="normalization-level-help" className="rounded-lg bg-white p-3 text-sm md:col-span-2">
        <p className="font-semibold">Cách đọc cấp công việc</p>
        <p className="mt-1 text-slate-600">Đếm các phần ngăn bằng dấu chấm: No. 1 là cấp 1, 1.2 là cấp 2, 1.2.1 là cấp 3. Đây là gợi ý từ file, không phải thay đổi cây công việc trên 1Office.</p>
        <p className="mt-2">Đang chọn: hạng mục ở cấp {options.firstTaskLevel}, gói thầu ở cấp {options.packageLevel}. Cấp {options.packageLevel + 1} được gợi ý là Segment (nhóm việc dưới gói), các cấp sâu hơn là công việc con.</p>
        <p className="mt-2">Ví dụ minh họa cách ghép, không phải mã đã duyệt: dự án DA, hạng mục HM, gói GT.001 → DA.HM.GT.001 → DA.HM.GT.001.01 → DA.HM.GT.001.01.01.</p>
        <p className="mt-1 text-slate-600">Cấp lấy từ No. trong file, không đếm dấu chấm trong mã công việc. Mã danh mục như GT.001 được giữ nguyên. {catalog ? 'Đã nạp danh mục ERP để kiểm tra mã dự án/gói thầu; chưa xác minh mã hạng mục tiến độ và segment.' : 'Chưa tự kiểm tra mã với danh mục ERP. Nạp danh mục ở phần bên dưới để đối chiếu.'}</p>
        <p className="mt-1 text-slate-600">{options.firstTaskLevel === 1 ? 'Bắt đầu lập mã từ cấp 1, không bỏ qua cấp phía trên.' : `Các dòng cấp ${options.firstTaskLevel === 2 ? '1' : `1 đến ${options.firstTaskLevel - 1}`} nằm phía trên hạng mục sẽ không được lập mã ở bước này.`}</p>
      </div>
    </fieldset>
    <details className="mt-4 rounded-lg border bg-white p-3">
      <summary className="cursor-pointer text-sm font-semibold">Thiết lập bổ sung: ghép mã thủ công hoặc khai báo việc mới</summary>
      <p className="mt-2 text-xs text-slate-600">Có thể để trống nếu chưa cần. Dùng số dòng hiển thị ở mé trái Excel, không dùng No. của công việc.</p>
      <fieldset disabled={disabled} className="mt-3 grid gap-4 md:grid-cols-2">
        <div>
          <label htmlFor="normalization-mappings" className="text-sm font-semibold">Mã đầy đủ cho từng dòng Excel (tùy chọn)</label>
          <textarea id="normalization-mappings" aria-describedby="normalization-mappings-help" className="mt-1 min-h-24 w-full rounded border p-2 font-mono" maxLength={1000000} placeholder={'27\tDA.HM.GT.001.01'} value={options.mappings} onChange={(e) => change({ mappings: e.target.value })} />
          <p id="normalization-mappings-help" className="mt-1 text-xs text-slate-600">Dán hai cột: Dòng Excel và Mã đầy đủ, không kèm tiêu đề. Khi lập mã mới, phải có đủ phần mã cha, không chỉ nhập mã gói thầu. Khi giữ mã cũ, mã phải khớp duy nhất một công việc trên 1Office. Mã ví dụ chỉ minh họa cấu trúc.</p>
        </div>
        {options.mode === 'update' && <div>
          <label htmlFor="normalization-new-rows" className="text-sm font-semibold">Những dòng là công việc mới, chưa có trên 1Office</label>
          <textarea id="normalization-new-rows" aria-describedby="normalization-new-rows-help" className="mt-1 min-h-24 w-full rounded border p-2 font-mono" maxLength={100000} placeholder="Ví dụ: 27, 31" value={options.newRows} onChange={(e) => change({ newRows: e.target.value })} />
          <p id="normalization-new-rows-help" className="mt-1 text-xs text-slate-600">Nhập số dòng Excel, cách nhau bằng dấu phẩy. Chỉ khai báo sau khi xác nhận đó không phải việc cũ đổi mã. Mã đề xuất không được trùng danh sách mã hiện có.</p>
        </div>}
      </fieldset>
    </details>
    {options.mode === 'update' && <p className="mt-3 rounded-lg bg-amber-50 p-3 text-sm text-amber-900">Dự án dùng để đối chiếu: <strong>{inventory?.project.name || '(chưa thu thập danh sách)'}</strong>. Kiểm tra đúng dự án trước khi ghép mã. Cần danh sách đầy đủ để xác nhận mã cũ và đề xuất mã mới.</p>}
    {options.mode === 'initial' && <p className="mt-3 rounded-lg bg-amber-50 p-3 text-sm text-amber-900">Không dùng “Lập mã lần đầu” cho công việc đã có trên 1Office. Nếu chèn việc vào file cũ rồi sinh lại mã theo thứ tự, mã của các việc phía sau có thể bị đổi.</p>}
    <ProjectErpCatalogReview report={report} options={options} catalog={catalog} disabled={disabled}
      onCatalog={(value) => { setCatalog(value); change({ erpPackages: {} }); }} onLoading={setCatalogLoading}
      onSelect={(row, code) => { const selections = { ...options.erpPackages }; if (code) selections[row] = code; else delete selections[row]; change({ erpPackages: selections }); }} />
    <button type="button" disabled={disabled || catalogLoading} className="mt-4 rounded-lg bg-blue-700 px-4 py-2 font-semibold text-white disabled:opacity-50" onClick={() => { setShow(true); setPage(0); }}>Tạo bảng mã xem trước</button>
    {result && counts && <div className="mt-5" role="status">
      <h4 className="font-semibold">Kết quả đề xuất, chưa áp dụng</h4>
      <div className="mt-3 grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
        {([
          ['Giữ mã trên 1Office', counts.existing, 'Đã đối chiếu được công việc hiện có.'],
          ['Có đề xuất mã mới', counts.proposed, 'Mã dự kiến để bạn kiểm tra, chưa tạo công việc.'],
          ['Cần kiểm tra', counts.blocked, 'Xem lý do ở từng dòng trong bảng bên dưới.'],
          ['Không lập mã ở bước này', counts.excluded, 'Các dòng thuộc cấp phía trên hạng mục đã chọn.'],
        ] as const).map(([label, count, help]) => <div key={label} className="rounded-lg border bg-white p-3"><p className="text-sm font-semibold">{label}</p><p className="mt-1 text-2xl font-bold">{count}</p><p className="mt-1 text-xs text-slate-600">{help}</p></div>)}
      </div>
      <p className={`mt-3 text-sm ${counts.blocked || result.errors.length ? 'text-amber-900' : 'text-slate-600'}`}>{counts.blocked || result.errors.length ? 'Chưa hoàn tất: còn dòng hoặc thiết lập cần kiểm tra. Có mã đề xuất không có nghĩa toàn bộ danh sách đã sẵn sàng import.' : 'Đã lập bảng xem trước cho các dòng trong phạm vi. Vẫn cần duyệt mã và dữ liệu trước khi import.'}</p>
      {!!result.errors.length && <ul role="alert" className="mt-2 list-inside list-disc text-red-700">{result.errors.map((error) => <li key={error}>{error}</li>)}</ul>}
      <details className="mt-3 rounded-lg bg-white p-3"><summary className="cursor-pointer text-sm font-semibold">Phạm vi của bước xem trước</summary><ul className="mt-2 list-inside list-disc text-sm text-slate-600">{result.warnings.map((warning) => <li key={warning}>{warning}</li>)}</ul></details>
      <div className="mt-3 overflow-auto"><table className="w-full text-left text-sm"><thead><tr>{['Vị trí trong Excel', 'Công việc và mã trong file', 'Nhóm công việc dự kiến', 'Mã công việc sau chuẩn hóa (xem trước)', 'Mã công việc cha theo file', 'Trạng thái và lý do cần kiểm tra'].map((h) => <th key={h} className="border-b p-2">{h}</th>)}</tr></thead>
        <tbody>{result.rows.slice(page * 50, (page + 1) * 50).map((r) => <tr key={r.row}><td className="border-b p-2">Dòng Excel {r.row}<p className="text-slate-500">No.: {r.no}</p></td><td className="border-b p-2">{r.name}<p className="text-slate-500">Mã trong file: {r.sourceCode || '(trống)'}</p>{r.erpCode && <p className="text-slate-500">Mã gói thầu dùng đối chiếu ERP: {r.erpCode}</p>}</td><td className="border-b p-2">{r.group}</td><td className="border-b p-2">{r.code || '(chưa xác định)'}</td><td className="border-b p-2">{r.parentCode || (r.kind === 'excluded' ? '(không lập mã)' : r.kind === 'blocked' ? '(chưa xác định mã cha)' : '(không có mã cha trong phạm vi)')}</td><td className="border-b p-2">{stateLabels[r.kind]}{r.internalId ? <p className="text-slate-500">ID công việc đối chiếu: {r.internalId}</p> : null}<p className="text-amber-900">{r.issues.join('; ')}</p></td></tr>)}</tbody></table></div>
      <div className="mt-3 flex flex-wrap items-center gap-4"><button type="button" disabled={page === 0} onClick={() => setPage(page - 1)}>Trang trước</button><span>Trang {page + 1}/{lastPage + 1}</span><button type="button" disabled={page >= lastPage} onClick={() => setPage(page + 1)}>Trang sau</button></div>
      <ProjectImportDiff key={`${inventory?.project.internalId || ''}:${inventory?.capturedAt || ''}`} report={report} proposal={result} inventory={inventory} disabled={disabled} />
    </div>}
  </div>;
}
