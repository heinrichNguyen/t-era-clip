import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { afterEach, expect, it, vi } from 'vitest';
import { openProjectImportWorkspace, readProjectImportContext } from '../../src/oneoffice/projectImportWorkspace';
import { ProjectImportLauncher } from '../../src/shared/ProjectImportLauncher';
import { ProjectImportApp } from '../../src/import-review/App';
import type { OneOfficeProjectInventory } from '../../src/oneoffice/types';

const inventory = { project: { type: 'project', internalId: '870', name: 'Test project', url: 'https://office.meygroup.vn/apps/work-project-project/view?ID=870' }, capturedAt: '2026-10-21T01:00:00Z', tasks: [], collection: { complete: true } } as unknown as OneOfficeProjectInventory;
afterEach(() => { vi.unstubAllGlobals(); document.body.innerHTML = ''; window.history.replaceState(null, '', '/'); });
function chrome() {
  const stored: Record<string, unknown> = {};
  const get = vi.fn((_: unknown, done: (value: unknown) => void) => done(stored));
  const set = vi.fn((value: unknown, done: () => void) => { Object.assign(stored, value); done(); });
  const create = vi.fn().mockResolvedValue({ id: 10 }); const executeScript = vi.fn();
  const localGet = vi.fn();
  vi.stubGlobal('chrome', { runtime: { getURL: (path: string) => `chrome-extension://example/${path}` }, tabs: { create }, scripting: { executeScript }, storage: { local: { get: localGet }, session: { get, set } } });
  return { stored, get, set, create, executeScript, localGet };
}

it('opens the standalone tool without collecting data or reading a stale shared inventory', async () => {
  const api = chrome(); api.stored.oneOfficeNoReview = inventory;
  await openProjectImportWorkspace();
  expect(api.create).toHaveBeenCalledExactlyOnceWith({ url: 'chrome-extension://example/import-review.html' });
  expect(await readProjectImportContext('')).toBeNull();
  expect(api.get).not.toHaveBeenCalled(); expect(api.set).not.toHaveBeenCalled(); expect(api.executeScript).not.toHaveBeenCalled();
});

it('opens the separate page with explicit snapshot context and rejects changed or removed snapshots', async () => {
  const api = chrome();
  await openProjectImportWorkspace(inventory, 5);
  const url = new URL(api.create.mock.calls[0][0].url);
  expect(url.pathname).toBe('/import-review.html'); expect(url.searchParams.get('sourceTabId')).toBe('5');
  expect(await readProjectImportContext(url.search)).toBe(inventory);
  api.stored.oneOfficeNoReview = { ...inventory, capturedAt: 'new snapshot' };
  await expect(readProjectImportContext(url.search)).rejects.toThrow('đã thay đổi');
  delete api.stored.oneOfficeNoReview;
  await expect(readProjectImportContext(url.search)).rejects.toThrow('đã thay đổi');
  expect(api.executeScript).not.toHaveBeenCalled(); expect(api.localGet).not.toHaveBeenCalled();
});

it('offers an accessible floating launcher, respects disabled state and reports opening errors', async () => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
  const api = chrome(); api.create.mockRejectedValue(new Error('Cannot open tab'));
  const container = document.createElement('div'); document.body.append(container); const root = createRoot(container);
  try {
    await act(async () => { root.render(React.createElement(ProjectImportLauncher, { floating: true, disabled: true })); });
    expect(container.firstElementChild?.className).toContain('fixed');
    const button = container.querySelector('button')!;
    expect(button.getAttribute('aria-label')).toBe('Mở Chuẩn hóa mã dự án');
    await act(async () => { button.click(); }); expect(api.create).not.toHaveBeenCalled();
    await act(async () => { root.render(React.createElement(ProjectImportLauncher, { floating: true })); });
    await act(async () => { button.click(); });
    expect(container.querySelector('[role=alert]')?.textContent).toContain('Cannot open tab');
    expect(button.disabled).toBe(false);
  } finally { await act(async () => { root.unmount(); }); }
});

it('renders the independent workspace without write/test controls, journal access or network calls', async () => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
  const api = chrome(); const fetch = vi.fn(); vi.stubGlobal('fetch', fetch);
  const container = document.createElement('div'); document.body.append(container); const root = createRoot(container);
  try {
    await act(async () => { root.render(React.createElement(ProjectImportApp)); });
    expect(container.querySelector('h1')?.textContent).toBe('Chuẩn hóa mã dự án');
    expect(container.querySelectorAll('ol li')).toHaveLength(3);
    expect(container.textContent).toContain('Đối chiếu mã + No. theo file duyệt');
    expect(container.textContent).toContain('Chưa bật ghi mã và No.');
    expect(container.querySelector('#project-import-file')).not.toBeNull();
    expect(container.textContent).not.toContain('Kiểm tra biểu mẫu trên dự án test');
    expect(container.textContent).not.toContain('Ghi No. theo');
    expect(container.textContent).toContain('Kiểm tra file trên máy');
    expect(api.localGet).not.toHaveBeenCalled(); expect(api.executeScript).not.toHaveBeenCalled(); expect(fetch).not.toHaveBeenCalled();
  } finally { await act(async () => { root.unmount(); }); }
});

it('groups numbering inside the workspace, keeps file controls mounted across tabs, and does not load an unrelated snapshot', async () => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
  const api = chrome();
  api.stored.oneOfficeNoReview = inventory;
  api.localGet.mockImplementation((_: unknown, done: (value: unknown) => void) => done({}));
  const container = document.createElement('div'); document.body.append(container); const root = createRoot(container);
  try {
    await act(async () => { root.render(React.createElement(ProjectImportApp)); });
    const sourceInput = container.querySelector('#project-import-file');
    expect(container.querySelector('#codes-tab')?.getAttribute('aria-selected')).toBe('true');
    await act(async () => { container.querySelector<HTMLButtonElement>('#numbering-tab')!.click(); });
    expect(container.querySelector<HTMLElement>('#codes-panel')!.hidden).toBe(true);
    expect(container.querySelector<HTMLElement>('#numbering-panel')!.hidden).toBe(false);
    expect(container.querySelector('#numbering-panel')?.textContent).toContain('Đối chiếu No. công việc hàng loạt');
    expect(container.querySelector('#numbering-panel textarea')).toBeNull();
    expect(container.querySelector('#numbering-panel input[type=file]')).toBeNull();
    expect(api.get).not.toHaveBeenCalled();
    await act(async () => { container.querySelector<HTMLButtonElement>('#codes-tab')!.click(); });
    expect(container.querySelector('#project-import-file')).toBe(sourceInput);
    expect(container.querySelector<HTMLElement>('#numbering-panel')!.hidden).toBe(true);
    expect(api.executeScript).not.toHaveBeenCalled();
  } finally { await act(async () => { root.unmount(); }); }
});

it('uses only the validated project context for numbering and still requires confirmation before a test write', async () => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
  const api = chrome();
  api.localGet.mockImplementation((_: unknown, done: (value: unknown) => void) => done({}));
  const snapshot = { ...inventory, tasks: [{ internalId: '100119', code: 'TEST01.3', name: 'Test task', no: '1.3', sourceNo: '1.3' }] };
  api.stored.oneOfficeNoReview = snapshot;
  window.history.replaceState(null, '', `?projectId=870&capturedAt=${encodeURIComponent(inventory.capturedAt)}&sourceTabId=5`);
  const confirm = vi.spyOn(window, 'confirm').mockReturnValue(false);
  const container = document.createElement('div'); document.body.append(container); const root = createRoot(container);
  try {
    await act(async () => { root.render(React.createElement(ProjectImportApp)); });
    api.stored.oneOfficeNoReview = { ...inventory, project: { ...inventory.project, internalId: '999', name: 'Unrelated' } };
    await act(async () => { container.querySelector<HTMLButtonElement>('#numbering-tab')!.click(); });
    expect(api.get).toHaveBeenCalledTimes(1);
    expect(container.querySelector('#numbering-panel')?.textContent).toContain('Test project');
    expect(container.querySelector('#numbering-panel')?.textContent).not.toContain('Unrelated');
    const input = container.querySelector<HTMLTextAreaElement>('#numbering')!;
    await act(async () => {
      Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!.call(input, 'TEST01.3\t1.4');
      input.dispatchEvent(new Event('input', { bubbles: true }));
    });
    await act(async () => { [...container.querySelectorAll<HTMLButtonElement>('#numbering-panel button')].find((b) => b.textContent === 'Xem trước thay đổi')!.click(); });
    const write = [...container.querySelectorAll<HTMLButtonElement>('#numbering-panel button')].find((b) => b.textContent?.startsWith('Ghi No. theo'))!;
    expect(write.disabled).toBe(false);
    await act(async () => { write.click(); });
    expect(confirm).toHaveBeenCalledWith(expect.stringContaining('KHÔNG tự hoàn nguyên'));
    expect(api.executeScript).not.toHaveBeenCalled();
  } finally { confirm.mockRestore(); await act(async () => { root.unmount(); }); }
});
