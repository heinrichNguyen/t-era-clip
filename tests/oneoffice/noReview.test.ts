import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { afterEach, expect, it, vi } from 'vitest';
import { NoReviewApp } from '../../src/no-review/App';

afterEach(() => { vi.unstubAllGlobals(); document.body.innerHTML = ''; });

function localJournalStorage(initial: Record<string, unknown> = {}) {
  const data: Record<string, unknown> = { ...initial };
  return {
    get: (key: string, done: (value: unknown) => void) => done({ [key]: data[key] }),
    set: (value: Record<string, unknown>, done: () => void) => { Object.assign(data, value); done(); },
  };
}

it('enables previewed changes for all five allowlisted task codes without sending writes', async () => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
  const tasks = [
    ['100113', 'TEST01.01', '1.1'], ['100116', 'TEST01.2', '1.2'],
    ['100119', 'TEST01.3', '1.7'], ['100122', 'TEST01.INSERT', '1.6'], ['100143', 'TEST01.4', ''],
  ].map(([internalId, code, sourceNo]) => ({ internalId, code, sourceNo, no: sourceNo || '1.1.5', name: code }));
  const executeScript = vi.fn();
  vi.stubGlobal('chrome', { runtime: {}, scripting: { executeScript }, storage: {
    local: localJournalStorage(), session: { get: (_: unknown, done: (value: unknown) => void) => done({ oneOfficeNoReview: {
      project: { type: 'project', internalId: '870', name: 'test', url: 'https://office.meygroup.vn/apps/work-project-project/view?ID=870' }, collection: { complete: true }, tasks,
    } }) },
  } });
  const container = document.createElement('div'); document.body.append(container);
  const root = createRoot(container);
  try {
    await act(async () => { root.render(React.createElement(NoReviewApp)); });
    const input = container.querySelector('textarea')!;
    await act(async () => {
      Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!.call(input, tasks.map((task, i) => `${task.code}\t2.${i + 1}`).join('\n'));
      input.dispatchEvent(new Event('input', { bubbles: true }));
    });
    await act(async () => { container.querySelector('button')!.click(); });
    expect(container.textContent).toContain('5 thay đổi');
    expect(container.querySelector('tbody')!.textContent).toContain('TEST01.4TEST01.4(trống)2.5');
    const write = [...container.querySelectorAll('button')].find((button) => button.textContent?.startsWith('Ghi No. theo'))!;
    expect(write.disabled).toBe(false); expect(executeScript).not.toHaveBeenCalled();
  } finally { await act(async () => { root.unmount(); }); }
});

it('loads pending recovery intent after reopening and does not offer a new write or delete it', async () => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
  const executeScript = vi.fn();
  vi.stubGlobal('chrome', {
    runtime: {}, scripting: { executeScript }, storage: {
      local: localJournalStorage({ oneOfficeNoWriteJournalV1: {
        version: 1, id: 'interrupted-run', origin: 'https://office.meygroup.vn', projectId: '870',
        createdAt: '2026-10-16', updatedAt: '2026-10-16', status: 'pending', rows: [
          { internalId: '100119', code: 'TEST01.3', before: '1.3', after: '1.4', temporary: '99001.1', observed: null, state: 'unknown' },
        ],
      } }),
      session: { get: (_: unknown, done: (value: unknown) => void) => done({}) },
    },
  });
  const container = document.createElement('div'); document.body.append(container);
  const root = createRoot(container);
  try {
    await act(async () => { root.render(React.createElement(NoReviewApp)); });
    expect(container.textContent).toContain('Chưa đối soát, khóa lượt ghi mới');
    expect(container.textContent).toContain('TEST01.3: 1.3 → 1.4');
    expect(container.textContent).toContain('Kiểm tra trạng thái trên 1Office (không ghi)');
    expect(container.textContent).not.toContain('Xóa nhật ký đã đối soát');
    expect(executeScript).not.toHaveBeenCalled();
  } finally { await act(async () => { root.unmount(); }); }
});

it('previews an ID-matched No. change without calling any write or scripting API', async () => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
  const fetch = vi.fn();
  const executeScript = vi.fn();
  vi.stubGlobal('fetch', fetch);
  vi.stubGlobal('chrome', {
    runtime: {}, scripting: { executeScript },
    tabs: {
      get: vi.fn().mockResolvedValue({ url: 'https://office.meygroup.vn/apps/work-project-project/view?ID=870' }),
      update: vi.fn().mockResolvedValue({ active: true }),
    },
    storage: { local: localJournalStorage(), session: { get: (_: unknown, callback: (data: unknown) => void) => callback({
      oneOfficeNoReview: {
        project: { type: 'project', internalId: '870', name: 'test', url: 'https://office.meygroup.vn/apps/work-project-project/view?ID=870' }, capturedAt: '2026-10-14',
        collection: { complete: true },
        tasks: [{ internalId: '100119', code: 'TEST01.3', name: 'Công việc cha 3', no: '1.3' }],
      },
    }) } },
  });
  const container = document.createElement('div');
  document.body.append(container);
  const root = createRoot(container);
  try {
    await act(async () => { root.render(React.createElement(NoReviewApp)); });
    const input = container.querySelector('textarea')!;
    await act(async () => {
      Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!.call(input, 'TEST01.3\t1.4');
      input.dispatchEvent(new Event('input', { bubbles: true }));
    });
    await act(async () => { container.querySelector('button')!.click(); });
    expect(container.textContent).toContain('1 thay đổi');
    expect(container.querySelector('tbody')?.textContent).toContain('TEST01.3Công việc cha 31.31.4');
    expect(container.textContent).toContain('Ghi No. theo bảng xem trước (dự án test)');
    expect(fetch).not.toHaveBeenCalled();
    expect(executeScript).not.toHaveBeenCalled();
  } finally {
    await act(async () => { root.unmount(); });
  }
});

it('requires explicit confirmation for the preview and locks further writes after submission', async () => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
  vi.stubGlobal('navigator', { locks: { request: (_: unknown, __: unknown, action: (lock: unknown) => unknown) => action({}) } });
  const previousUrl = window.location.href;
  window.history.replaceState(null, '', '?sourceTabId=1');
  const confirm = vi.spyOn(window, 'confirm').mockReturnValue(false);
  const executeScript = vi.fn().mockResolvedValue([{ result: {
    mode: 'plan', readOnly: false, complete: true, phase: 'Đã xác nhận', saveAttempts: 2,
    error: null, tasks: [{ code: 'TEST01.3', before: '1.3', after: '1.5', temporary: '99001.1', applied: true, finalized: true, restored: false, error: null }],
  } }]);
  vi.stubGlobal('chrome', {
    runtime: {}, scripting: { executeScript },
    tabs: { get: vi.fn().mockResolvedValue({ url: 'https://office.meygroup.vn/apps/work-project-project/view?ID=870' }), update: vi.fn().mockResolvedValue({ active: true }) },
    storage: { local: localJournalStorage(), session: { get: (_: unknown, callback: (data: unknown) => void) => callback({ oneOfficeNoReview: {
      project: { type: 'project', internalId: '870', name: 'test', url: 'https://office.meygroup.vn/apps/work-project-project/view?ID=870' }, capturedAt: '2026-10-16', collection: { complete: true },
      tasks: [{ internalId: '100119', code: 'TEST01.3', name: 'Test', no: '1.3', sourceNo: '1.3' }],
    } }) } },
  });
  const container = document.createElement('div'); document.body.append(container);
  const root = createRoot(container);
  try {
    await act(async () => { root.render(React.createElement(NoReviewApp)); });
    const input = container.querySelector('textarea')!;
    await act(async () => {
      Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!.call(input, 'TEST01.3\t1.5');
      input.dispatchEvent(new Event('input', { bubbles: true }));
    });
    await act(async () => { container.querySelector('button')!.click(); });
    const write = Array.from(container.querySelectorAll('button')).find((button) => button.textContent?.startsWith('Ghi No. theo'))!;
    await act(async () => { write.click(); });
    expect(confirm).toHaveBeenCalledWith(expect.stringContaining('KHÔNG tự hoàn nguyên'));
    expect(executeScript).not.toHaveBeenCalled();
    confirm.mockReturnValue(true);
    await act(async () => { write.click(); });
    expect(executeScript).toHaveBeenCalledExactlyOnceWith(expect.objectContaining({ args: [false, {
      silentReads: true,
      directEdit: true,
      changes: [{ internalId: '100119', code: 'TEST01.3', before: '1.3', after: '1.5' }], sourceTasks: [{ internalId: '100119', no: '1.3' }],
    }] }));
    expect(write.disabled).toBe(true);
    expect(container.textContent).toContain('Đã ghi và đọc lại xác nhận toàn bộ No. mới');
    expect(container.textContent).toContain('Đã khóa ghi tiếp');
  } finally {
    await act(async () => { root.unmount(); });
    confirm.mockRestore(); window.history.replaceState(null, '', previousUrl);
  }
});

it('shows menu diagnostics from the read-only experiment without asking to write', async () => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
  vi.stubGlobal('navigator', { locks: { request: (_: unknown, __: unknown, action: (lock: unknown) => unknown) => action({}) } });
  const previousUrl = window.location.href;
  window.history.replaceState(null, '', '?sourceTabId=1');
  const confirm = vi.spyOn(window, 'confirm');
  const executeScript = vi.fn().mockResolvedValue([{ result: {
    readOnly: true, complete: false, phase: 'Mở menu Sửa', saveAttempts: 0,
    error: 'Hết thời hạn', tasks: [], diagnostics: ['TEST01.3: chưa thấy liên kết hiển thị.'],
  } }]);
  vi.stubGlobal('chrome', {
    runtime: {}, scripting: { executeScript },
    tabs: {
      get: vi.fn().mockResolvedValue({ url: 'https://office.meygroup.vn/apps/work-project-project/view?ID=870' }),
      update: vi.fn().mockResolvedValue({ active: true }),
    },
    storage: { local: localJournalStorage(), session: { get: (_: unknown, callback: (data: unknown) => void) => callback({
      oneOfficeNoReview: { project: { type: 'project', internalId: '870', name: 'test', url: 'https://office.meygroup.vn/apps/work-project-project/view?ID=870' }, tasks: [] },
    }) } },
  });
  const container = document.createElement('div'); document.body.append(container);
  const root = createRoot(container);
  try {
    await act(async () => { root.render(React.createElement(NoReviewApp)); });
    await act(async () => {
      Array.from(container.querySelectorAll('button')).find((b) => b.textContent === 'Kiểm tra mở biểu mẫu (không ghi)')!.click();
    });
    expect(executeScript).toHaveBeenCalledWith(expect.objectContaining({ target: { tabId: 1 }, args: [true, null, false, true] }));
    expect(confirm).not.toHaveBeenCalled();
    expect(container.textContent).toContain('Chẩn đoán mở biểu mẫu (bản 14)');
    expect(container.textContent).toContain('TEST01.3: chưa thấy liên kết hiển thị.');
    expect(container.textContent).toContain('Số lần kích hoạt lưu: 0');
  } finally {
    await act(async () => { root.unmount(); });
    confirm.mockRestore();
    window.history.replaceState(null, '', previousUrl);
  }
});
