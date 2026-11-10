// @vitest-environment jsdom
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { expect, it, vi } from 'vitest';
import { AccessReviewApp } from '../../src/access-review/App';
import { loadAccessPageFromTab } from '../../src/oneoffice/accessHistory';
import { downloadAccessWorkbook } from '../../src/oneoffice/export/accessXlsx';
import { loadAccessPersonnel } from '../../src/oneoffice/accessPersonnel';

it('clears report and date snapshot so changing the cutoff cannot restore erased data', async () => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
  const remove = vi.fn((_keys: string[], done: () => void) => done());
  vi.stubGlobal('chrome', { runtime: {}, storage: { session: { remove, set: (_value: unknown, done: () => void) => done() } } });
  const confirm = vi.spyOn(window, 'confirm').mockReturnValue(true);
  window.history.replaceState({}, '', '/access-review.html?sourceTabId=42');
  vi.mocked(loadAccessPageFromTab).mockResolvedValue({ start: 1, end: 1, total: 1, records: [
    { id: '1', account: 'A', name: 'Private report name', timeRaw: '07:00:00 01/01/2020', timeIso: '2020-01-01T00:00:00Z' },
  ] });
  vi.mocked(loadAccessPersonnel).mockResolvedValue({ collectedAt: '', activeAccounts: ['A'], users: [], warnings: [] });
  const container = document.createElement('div'); document.body.append(container);
  const root = createRoot(container);
  const button = (label: string) => [...container.querySelectorAll('button')].find((node) => node.textContent === label);
  try {
    await act(async () => root.render(React.createElement(AccessReviewApp)));
    await act(async () => button('Thu thập nhật ký')!.click());
    expect(container.textContent).toContain('Private report name');
    await act(async () => button('Xóa dữ liệu tạm')!.click());
    expect(remove).toHaveBeenCalledWith(['oneOfficeDiscussionReview', 'oneOfficeProjectBatchSession', 'oneOfficeNoReview'], expect.any(Function));
    expect(container.textContent).toContain('Đã xóa dữ liệu tạm');
    expect(container.textContent).not.toContain('Private report name');
    expect(button('Xuất Excel')).toBeUndefined();
    expect(container.querySelector('input[type="date"]')).toBeNull();
  } finally {
    await act(async () => root.unmount()); container.remove(); confirm.mockRestore(); vi.unstubAllGlobals();
    window.history.replaceState({}, '', '/');
  }
});

vi.mock('../../src/oneoffice/accessHistory', async (original) => ({
  ...await original<typeof import('../../src/oneoffice/accessHistory')>(),
  loadAccessPageFromTab: vi.fn(),
}));
vi.mock('../../src/oneoffice/export/accessXlsx', () => ({ downloadAccessWorkbook: vi.fn().mockResolvedValue(undefined) }));
vi.mock('../../src/oneoffice/accessPersonnel', () => ({ loadAccessPersonnel: vi.fn().mockResolvedValue({ collectedAt: '', users: [], warnings: [], activeAccounts: [] }) }));

it('only offers export after complete collection and removes stale results on retry failure', async () => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
  window.history.replaceState({}, '', '/access-review.html?sourceTabId=42');
  const confirm = vi.spyOn(window, 'confirm').mockReturnValue(true);
  vi.mocked(loadAccessPageFromTab).mockResolvedValue({ start: 0, end: 0, total: 0, records: [] });
  const container = document.createElement('div'); document.body.append(container);
  const root = createRoot(container);
  const button = (label: string) => [...container.querySelectorAll('button')].find((node) => node.textContent === label || node.getAttribute('aria-label') === label);
  try {
    await act(async () => root.render(React.createElement(AccessReviewApp)));
    expect(button('Xuất Excel')).toBeUndefined();
    await act(async () => button('Thu thập nhật ký')!.click());
    expect(button('Xuất Excel')).toBeDefined();
    expect(loadAccessPageFromTab).toHaveBeenCalledWith(42, 1, undefined, 'https://office.meygroup.vn');
    await act(async () => button('Xuất Excel')!.click());
    expect(downloadAccessWorkbook).toHaveBeenCalledTimes(1);
    vi.mocked(loadAccessPageFromTab).mockRejectedValue(new Error('Nguồn thay đổi'));
    await act(async () => button('Thu thập lại từ 1Office')!.click());
    expect(button('Xuất Excel')).toBeUndefined();
    expect(container.textContent).toContain('Nguồn thay đổi');
  } finally {
    await act(async () => root.unmount()); container.remove(); confirm.mockRestore(); vi.unstubAllGlobals();
    window.history.replaceState({}, '', '/');
  }
});

it('shows current personnel but blocks export when active status can no longer be verified', async () => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
  window.history.replaceState({}, '', '/access-review.html?sourceTabId=42');
  vi.mocked(loadAccessPageFromTab).mockResolvedValue({ start: 1, end: 1, total: 1, records: [
    { id: '1', account: 'A', name: 'Historical name', timeRaw: '07:00:00 01/01/2020', timeIso: '2020-01-01T00:00:00Z' },
  ] });
  vi.mocked(loadAccessPersonnel).mockResolvedValueOnce({ collectedAt: '2026-10-06T00:00:00Z', activeAccounts: ['A'],
    users: [{ account: 'A', position: 'Current position', department: 'Current department' }], warnings: [] });
  const container = document.createElement('div'); document.body.append(container);
  const root = createRoot(container);
  const button = (label: string) => [...container.querySelectorAll('button')].find((node) => node.textContent === label || node.getAttribute('aria-label') === label);
  try {
    await act(async () => root.render(React.createElement(AccessReviewApp)));
    await act(async () => button('Thu thập nhật ký')!.click());
    expect(container.textContent).toContain('Current position');
    expect(container.textContent).toContain('Current department');
    expect(loadAccessPersonnel).toHaveBeenLastCalledWith(42, ['A'], expect.any(AbortSignal), expect.any(Function), 'https://office.meygroup.vn');
    const historyCalls = vi.mocked(loadAccessPageFromTab).mock.calls.length;
    vi.mocked(loadAccessPersonnel).mockResolvedValueOnce({ collectedAt: '2026-10-06T01:00:00Z', activeAccounts: ['A'],
      users: [{ account: 'A', position: 'Updated position', department: 'Updated department' }], warnings: [] });
    await act(async () => button('Cập nhật vị trí/phòng ban')!.click());
    expect(container.textContent).toContain('Updated position');
    expect(loadAccessPageFromTab).toHaveBeenCalledTimes(historyCalls);
    vi.mocked(loadAccessPersonnel).mockRejectedValueOnce(new Error('Thiếu quyền đọc người dùng'));
    await act(async () => button('Thu thập lại từ 1Office')!.click());
    expect(container.textContent).toContain('Thiếu quyền đọc người dùng');
    expect(container.textContent).not.toContain('Historical name');
    expect(container.textContent).not.toContain('Current position');
    expect(container.textContent).not.toContain('Updated position');
    expect(button('Xuất Excel')).toBeUndefined();
  } finally {
    await act(async () => root.unmount()); container.remove(); vi.unstubAllGlobals();
    window.history.replaceState({}, '', '/');
  }
});

it('selects TDR before reading history and invalidates the snapshot when filters change', async () => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
  window.history.replaceState({}, '', '/access-review.html?sourceTabId=42');
  vi.mocked(loadAccessPersonnel).mockResolvedValue({ collectedAt: '', warnings: [], activeAccounts: ['A', 'B'], users: [
    { account: 'A', position: 'Giám đốc', department: 'R - Hội đồng quản trị', companyCode: 'TDR', company: 'Meyland' },
    { account: 'B', position: '', department: '', companyCode: 'OTHER' },
  ] });
  vi.mocked(loadAccessPageFromTab).mockClear().mockResolvedValue({ start: 1, end: 1, total: 1, records: [
    { id: '1', account: 'A', name: 'A', timeRaw: '07:00:00 01/01/2020', timeIso: '2020-01-01T00:00:00Z' },
  ] });
  const container = document.createElement('div'); document.body.append(container);
  const root = createRoot(container);
  const button = (label: string) => [...container.querySelectorAll('button')].find((node) => node.textContent === label || node.getAttribute('aria-label') === label);
  try {
    await act(async () => root.render(React.createElement(AccessReviewApp)));
    await act(async () => button('Chọn công ty TDR / Meyland')!.click());
    await act(async () => button('Thu thập nhật ký')!.click());
    expect(loadAccessPersonnel).toHaveBeenLastCalledWith(42, null, expect.any(AbortSignal), expect.any(Function), 'https://office.meygroup.vn');
    expect(vi.mocked(loadAccessPageFromTab).mock.calls.every((args) => args[2] === 'A')).toBe(true);
    expect(container.textContent).toContain('1 tài khoản hiện còn active');
    expect(container.textContent).toContain('Meyland');
    expect(button('Xuất Excel')).toBeDefined();
    await act(async () => button('Chọn công ty TDR / Meyland')!.click());
    expect(button('Xuất Excel')).toBeUndefined();
    vi.mocked(loadAccessPersonnel).mockResolvedValueOnce({ collectedAt: '', warnings: [], users: [], activeAccounts: [] });
    vi.mocked(loadAccessPageFromTab).mockClear();
    await act(async () => button('Thu thập nhật ký')!.click());
    expect(loadAccessPageFromTab).not.toHaveBeenCalled();
    expect(container.textContent).toContain('Chưa xác minh được cơ cấu');
  } finally {
    await act(async () => root.unmount()); container.remove(); vi.unstubAllGlobals();
    window.history.replaceState({}, '', '/');
  }
});

it('excludes historical records for locked users in the unfiltered report and removes newly locked users on refresh', async () => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
  window.history.replaceState({}, '', '/access-review.html?sourceTabId=42');
  vi.mocked(loadAccessPersonnel).mockResolvedValue({ collectedAt: '2026-10-06T00:00:00Z', warnings: [], activeAccounts: ['ACTIVE'], users: [
    { account: 'ACTIVE', position: '', department: '' },
  ] });
  vi.mocked(loadAccessPageFromTab).mockClear().mockResolvedValue({ start: 1, end: 2, total: 2, records: [
    { id: '2', account: 'LOCKED', name: 'Locked name', timeRaw: '07:00:00 01/01/2020', timeIso: '2020-01-01T00:00:00Z' },
    { id: '1', account: 'ACTIVE', name: 'Active name', timeRaw: '07:00:00 01/01/2020', timeIso: '2020-01-01T00:00:00Z' },
  ] });
  const container = document.createElement('div'); document.body.append(container);
  const root = createRoot(container);
  const button = (label: string) => [...container.querySelectorAll('button')].find((node) => node.textContent === label || node.getAttribute('aria-label') === label);
  const confirm = vi.spyOn(window, 'confirm').mockReturnValue(true);
  try {
    await act(async () => root.render(React.createElement(AccessReviewApp)));
    await act(async () => button('Thu thập nhật ký')!.click());
    expect(container.textContent).toContain('Active name');
    expect(container.textContent).not.toContain('Locked name');
    await act(async () => button('Xuất Excel')!.click());
    const exported = vi.mocked(downloadAccessWorkbook).mock.calls.at(-1)![0];
    expect(exported.recordCount).toBe(1);
    expect(exported.users.map((row) => row.account)).toEqual(['ACTIVE']);
    vi.mocked(loadAccessPersonnel).mockResolvedValueOnce({ collectedAt: '2026-10-06T01:00:00Z', warnings: [], activeAccounts: [], users: [] });
    await act(async () => button('Cập nhật vị trí/phòng ban')!.click());
    expect(container.textContent).not.toContain('Active name');
    expect(container.textContent).toContain('0 tài khoản có bản ghi');
  } finally {
    await act(async () => root.unmount()); container.remove(); confirm.mockRestore(); vi.unstubAllGlobals();
    window.history.replaceState({}, '', '/');
  }
});
