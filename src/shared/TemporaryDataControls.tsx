import React from 'react';
import { DISCUSSIONS_ONLY } from './buildProfile';

export function TemporaryDataCleared() {
  return <main className="mx-auto max-w-3xl p-8 text-slate-700">
    <h1 className="text-xl font-semibold text-blue-900">Đã xóa dữ liệu tạm</h1>
    <p className="mt-3">{DISCUSSIONS_ONLY ? 'Mở T-eraClip trên tab 1Office để thu thập lại. File đã tải và dữ liệu trên 1Office không bị xóa.' : 'Mở T-eraClip trên tab 1Office để thu thập lại. Capture, file đã tải và dữ liệu trên 1Office không bị xóa.'}</p>
  </main>;
}

export function ClearTemporaryDataButton({ onClick, disabled, error }: {
  onClick: () => void; disabled: boolean; error: string;
}) {
  return <div className="my-3">
    <button type="button" disabled={disabled} onClick={onClick}
      className="rounded-lg border border-slate-300 bg-white px-4 py-2 text-sm font-semibold text-slate-700 disabled:opacity-50">Xóa dữ liệu tạm</button>
    {error && <p role="alert" className="mt-2 text-sm text-red-700">{error}</p>}
  </div>;
}
