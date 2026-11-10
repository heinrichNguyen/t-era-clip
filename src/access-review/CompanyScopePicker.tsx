import React from 'react';
import { companyScopes, companyScopeLabel } from '../oneoffice/companyScopes';

export function CompanyScopePicker({ value, disabled, onChange }: {
  value: string; disabled: boolean; onChange: (value: string) => void;
}) {
  const selected = companyScopes.find((scope) => scope.code === value || scope.members.some((member) => member.code === value));
  return <fieldset disabled={disabled} className="rounded-xl border border-slate-200 bg-slate-50 p-4">
    <legend className="px-2 font-semibold text-blue-900">Phạm vi công ty</legend>
    <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-3">
      {companyScopes.map((scope) => {
        const active = selected?.code === scope.code;
        return <button key={scope.code} type="button" aria-label={`Chọn công ty ${scope.title}`} aria-pressed={active}
          onClick={() => onChange(scope.code)} className={`rounded-xl border p-4 text-left transition focus-visible:outline focus-visible:outline-2 focus-visible:outline-blue-600 disabled:opacity-50 ${active ? 'border-blue-600 bg-blue-50 ring-1 ring-blue-600' : 'border-slate-200 bg-white hover:border-blue-300'}`}>
          <span className="flex items-center justify-between gap-2 font-bold text-blue-900">{scope.title}<span aria-hidden="true" className={`h-4 w-4 rounded-full border ${active ? 'border-blue-600 bg-blue-600 ring-2 ring-blue-100' : 'border-slate-300'}`} /></span>
          <span className="mt-1 block text-sm text-slate-600">{scope.description}</span>
          {scope.members.length > 0 && <span className="mt-2 block text-xs text-blue-700">Toàn nhóm hoặc chọn 1 trong {scope.members.length} công ty thành viên</span>}
        </button>;
      })}
    </div>
    {selected && selected.members.length > 0 && <div className="mt-4 rounded-lg border border-blue-200 bg-white p-4">
      <label className="block text-sm font-semibold text-blue-900">Phạm vi trong {selected.title}
        <select aria-label={`Phạm vi trong ${selected.title}`} value={value} onChange={(event) => onChange(event.target.value)}
          className="mt-2 block w-full rounded-lg border border-slate-300 bg-white p-3">
          <option value={selected.code}>Toàn nhóm {selected.title}, gồm các công ty thành viên</option>
          {selected.members.map((member) => <option key={member.code} value={member.code}>{member.code} · {member.title}</option>)}
        </select>
      </label>
      <p className="mt-2 text-xs text-slate-500">Toàn nhóm gồm nhân sự cấp tổng công ty và các thành viên. Chọn riêng chỉ lấy nhân sự thuộc công ty đó.</p>
    </div>}
    <p className="mt-4 text-sm text-blue-800">Đang chọn: <strong>{companyScopeLabel(value)}</strong>. Chỉ tài khoản active.</p>
  </fieldset>;
}
