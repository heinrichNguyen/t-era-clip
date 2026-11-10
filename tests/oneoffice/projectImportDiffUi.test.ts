import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { afterEach, expect, it, vi } from 'vitest';
import { ProjectImportDiff } from '../../src/no-review/ProjectImportDiff';
import { reviewProjectImport } from '../../src/oneoffice/projectImportReview';
import { proposeImportNormalization } from '../../src/oneoffice/projectImportNormalization';
import type { OneOfficeProjectInventory } from '../../src/oneoffice/types';

afterEach(() => { vi.unstubAllGlobals(); document.body.innerHTML = ''; });

it('requires explicit field/project selection, renders the diff safely and invalidates it on mapping changes without network calls', async () => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
  const fetch = vi.fn(); const executeScript = vi.fn(); vi.stubGlobal('fetch', fetch); vi.stubGlobal('chrome', { scripting: { executeScript } });
  const inventory = { project: { name: 'Test', internalId: '870' }, capturedAt: '2026-10-20', collection: { complete: true }, tasks: [{ internalId: '10', code: 'P.A', name: 'Old', parentInternalId: null, sourceNo: '' }] } as OneOfficeProjectInventory;
  const report = reviewProjectImport({ sheet: 'S', data: [['No.', 'Mã gói thầu/Mã CV', 'Nội dung công việc', 'Bắt đầu', 'Bắt đầu thực tế'], ['1', 'P.A', '<img src=x onerror=alert(1)>', '2026-10-01', '2026-10-02']] }, inventory);
  const proposal = proposeImportNormalization(report, inventory, { mode: 'update', prefix: 'P', firstTaskLevel: 1, packageLevel: 2, mappings: '', newRows: '' });
  const container = document.createElement('div'); document.body.append(container); const root = createRoot(container);
  const checkbox = (text: string) => [...container.querySelectorAll('label')].find((label) => label.textContent?.includes(text))!.querySelector<HTMLInputElement>('input')!;
  const button = () => container.querySelector('button')!;
  try {
    await act(async () => { root.render(React.createElement(ProjectImportDiff, { report, proposal, inventory, disabled: false })); });
    expect([...container.querySelectorAll<HTMLInputElement>('input[type=checkbox]')].every((input) => !input.checked)).toBe(true);
    await act(async () => { button().click(); });
    expect(container.textContent).toContain('Chọn ít nhất một trường');
    await act(async () => { checkbox('Tên công việc').click(); checkbox('Tôi xác nhận file này').click(); });
    await act(async () => { button().click(); });
    expect(container.textContent).toContain('Thay đổi dự kiến');
    expect(container.textContent).toContain('Old');
    expect(container.textContent).toContain('<img src=x onerror=alert(1)>'); expect(container.querySelector('img')).toBeNull();
    expect(container.textContent).toContain('Chưa có nút áp dụng');
    await act(async () => { checkbox('Bắt đầu kế hoạch').click(); });
    expect(container.querySelector('table')).toBeNull();
    await act(async () => { button().click(); });
    expect(container.textContent).toContain('Cần xác nhận các cột ngày');
    await act(async () => { checkbox('Tôi xác nhận các cột ngày').click(); });
    const source = container.querySelector<HTMLSelectElement>('#diff-column-plannedStart')!;
    await act(async () => { source.value = 'Bắt đầu thực tế'; source.dispatchEvent(new Event('change', { bubbles: true })); });
    expect(checkbox('Tôi xác nhận các cột ngày').checked).toBe(false);
    await act(async () => { button().click(); });
    expect(container.textContent).toContain('Không dùng cột “Bắt đầu thực tế”');
    expect(fetch).not.toHaveBeenCalled(); expect(executeScript).not.toHaveBeenCalled();
  } finally { await act(async () => { root.unmount(); }); }
});
