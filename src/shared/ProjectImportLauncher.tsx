import React, { useState } from 'react';
import { openProjectImportWorkspace } from '../oneoffice/projectImportWorkspace';
import type { OneOfficeProjectInventory } from '../oneoffice/types';

export function ProjectImportIcon({ className = 'h-6 w-6' }: { className?: string }) {
  return <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6" aria-hidden="true"><rect x="4" y="3" width="16" height="18" rx="3" /><path d="M4 9h16M10 9v12M14 13h3M14 17h3M8 6h8" /></svg>;
}

export function ProjectImportLauncher({ inventory, sourceTabId, floating = false, disabled = false }: {
  inventory?: OneOfficeProjectInventory | null; sourceTabId?: number; floating?: boolean; disabled?: boolean;
}) {
  const [opening, setOpening] = useState(false);
  const [error, setError] = useState('');
  return <div className={floating ? 'fixed bottom-6 right-6 z-40 max-w-[calc(100vw-3rem)]' : 'shrink-0 border-t border-slate-200 bg-white p-3'}>
    <button type="button" disabled={disabled || opening} aria-label="Mở Chuẩn hóa mã dự án" onClick={() => {
      if (opening || disabled) return;
      setOpening(true); setError('');
      void openProjectImportWorkspace(inventory, sourceTabId).catch((e) => setError(e instanceof Error ? e.message : 'Không thể mở công cụ.')).finally(() => setOpening(false));
    }} className="flex w-full items-center gap-3 rounded-2xl bg-gradient-to-r from-blue-700 to-indigo-700 px-5 py-3 text-left text-white shadow-lg shadow-blue-900/25 transition hover:from-blue-800 hover:to-indigo-800 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-blue-600 disabled:cursor-not-allowed disabled:opacity-50">
      <span className="rounded-xl bg-white/15 p-2"><ProjectImportIcon /></span>
      <span><span className="block text-sm font-bold">{opening ? 'Đang mở công cụ...' : 'Chuẩn hóa mã dự án'}</span><span className="mt-0.5 block text-xs text-blue-100">File tiến độ · Mã ERP · So sánh</span></span>
      <svg aria-hidden="true" className="ml-auto h-5 w-5 shrink-0" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M7 17 17 7M7 7h10v10" /></svg>
    </button>
    {error && <p role="alert" className="mt-2 max-w-sm rounded-lg border border-red-200 bg-white p-2 text-sm text-red-700">{error}</p>}
  </div>;
}
