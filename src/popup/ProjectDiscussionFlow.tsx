import React from 'react';
import type { OneOfficeProjectInventory } from '../oneoffice/types';
import { projectInventoryWarningLabel } from '../shared/projectInventoryWarning';

export function ProjectDiscussionFlow({ inventory, checking, opening, disabled, onCollect, onCheckOnly, onCollectPage }: {
  inventory: OneOfficeProjectInventory | null; checking: boolean; opening: boolean; disabled: boolean;
  onCollect: () => void; onCheckOnly: () => void; onCollectPage: () => void;
}) {
  const warnings = inventory?.collection.warnings || [];
  const count = inventory?.tasks.length || 0;
  const depth = inventory ? Math.max(0, ...inventory.tasks.map((r) => r.depth)) : 0;
  return <div className="mt-4">
    <button type="button" disabled={disabled || checking || opening} onClick={onCollect}
      className="w-full rounded-xl bg-white px-4 py-3 text-sm font-bold text-blue-800 shadow-sm transition hover:bg-blue-50 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-white disabled:cursor-not-allowed disabled:opacity-60">
      {checking ? 'Đang kiểm tra danh sách...' : opening ? 'Đang mở trang thu thập...' : inventory && count && !inventory.collection.complete ? `Tiếp tục với ${count} công việc đã lấy` : 'Thu thập thảo luận toàn dự án'}
    </button>
    <p className="mt-2 text-[11px] leading-5 text-blue-100">Kiểm tra danh sách → Thu thập thảo luận → Xem trước và xuất Excel</p>
    {inventory && <div className="mt-3 rounded-xl border border-white/20 bg-white/10 p-3 text-xs">
      <p className="font-semibold">{inventory.collection.loadedTaskCount} công việc · tối đa {depth} cấp</p>
      <p className={`mt-1 ${inventory.collection.complete ? 'text-blue-100' : 'text-amber-200'}`}>{inventory.collection.complete ? 'Đã kiểm tra danh sách.' : 'Danh sách chưa đầy đủ. Kiểm tra cảnh báo trước khi tiếp tục.'}</p>
      {!!warnings.length && <details className="mt-2" open={!inventory.collection.complete}><summary className="cursor-pointer font-semibold text-amber-200">{warnings.length} cảnh báo danh sách</summary><ul className="mt-1 list-inside list-disc space-y-1 leading-5 text-amber-100">{warnings.map((warning, i) => <li key={i}>{projectInventoryWarningLabel(warning)}</li>)}</ul></details>}
      {!count && <p className="mt-2 text-amber-200">Chưa có công việc để thu thập. Hãy kiểm tra lại danh sách.</p>}
    </div>}
    <details className="mt-3 border-t border-white/20 pt-3 text-xs">
      <summary className="cursor-pointer font-medium text-blue-100">Tùy chọn khác</summary>
      <div className="mt-2 grid gap-2">
        <button type="button" disabled={disabled || checking || opening} onClick={onCheckOnly} className="rounded-lg border border-white/30 px-3 py-2 text-left text-white transition hover:bg-white/10 disabled:opacity-50">{inventory ? 'Kiểm tra lại danh sách công việc' : 'Chỉ kiểm tra danh sách công việc'}</button>
        <p className="text-[10px] leading-4 text-blue-200">Dùng danh sách này cho công cụ Chuẩn hóa mã dự án ở cuối menu.</p>
        <button type="button" disabled={disabled || checking || opening} onClick={onCollectPage} className="rounded-lg border border-white/30 px-3 py-2 text-left text-white transition hover:bg-white/10 disabled:opacity-50">Chỉ lấy thảo luận trang hiện tại</button>
      </div>
    </details>
  </div>;
}
