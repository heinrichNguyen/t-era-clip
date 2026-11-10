import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { afterEach, expect, it, vi } from 'vitest';
import { ProjectImportNormalization } from '../../src/no-review/ProjectImportNormalization';
import { reviewProjectImport } from '../../src/oneoffice/projectImportReview';

afterEach(() => { vi.unstubAllGlobals(); document.body.innerHTML = ''; });

it('defaults to preserving codes and requires explicitly selecting first-time coding; changes invalidate the preview', async () => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
  const fetch = vi.fn(); vi.stubGlobal('fetch', fetch);
  const executeScript = vi.fn(); vi.stubGlobal('chrome', { scripting: { executeScript } });
  const report = reviewProjectImport({ sheet: 'S', data: [['No.', 'Mã gói thầu/Mã CV', 'Nội dung công việc'], ['1.1.1', 'ITEM', 'I']] }, null);
  const container = document.createElement('div'); document.body.append(container);
  const root = createRoot(container);
  try {
    await act(async () => { root.render(React.createElement(ProjectImportNormalization, { report, inventory: null, disabled: false })); });
    const mode = container.querySelector<HTMLSelectElement>('#normalization-mode')!;
    expect(mode.value).toBe('update');
    const prefix = container.querySelector<HTMLInputElement>('#normalization-prefix')!;
    await act(async () => {
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!.call(prefix, 'P');
      prefix.dispatchEvent(new Event('input', { bubbles: true }));
      container.querySelector('button')!.click();
    });
    expect(container.textContent).toContain('Cần danh sách 1Office đầy đủ');
    await act(async () => { mode.value = 'initial'; mode.dispatchEvent(new Event('change', { bubbles: true })); });
    expect(container.querySelector('table')).toBeNull();
    await act(async () => { container.querySelector('button')!.click(); });
    expect(container.textContent).toContain('P.ITEM');
    expect(container.textContent).toContain('Có đề xuất mã mới1');
    expect(container.textContent).toContain('Kết quả đề xuất, chưa áp dụng');
    expect(container.textContent).toContain('Dòng Excel 2');
    await act(async () => {
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!.call(prefix, 'OTHER');
      prefix.dispatchEvent(new Event('input', { bubbles: true }));
    });
    expect(container.querySelector('table')).toBeNull();
    expect(fetch).not.toHaveBeenCalled(); expect(executeScript).not.toHaveBeenCalled();
  } finally { await act(async () => { root.unmount(); }); }
});

it('explains levels and manual fields, separates pending rows from completed proposals, and labels controls accessibly', async () => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
  const report = reviewProjectImport({ sheet: 'S', data: [
    ['No.', 'Mã gói thầu/Mã CV', 'Nội dung công việc'],
    ['1', 'P', 'Project'], ['1.1.1', '', 'Missing item code'],
  ] }, null);
  const container = document.createElement('div'); document.body.append(container);
  const root = createRoot(container);
  try {
    await act(async () => { root.render(React.createElement(ProjectImportNormalization, { report, inventory: null, disabled: false })); });
    expect(container.textContent).toContain('Bước 2: Xem trước mã công việc');
    expect(container.textContent).toContain('No. 1 là cấp 1, 1.2 là cấp 2, 1.2.1 là cấp 3');
    expect(container.textContent).toContain('Các dòng cấp 1 đến 2');
    expect(container.textContent).toContain('Mã dự án.Hạng mục.Gói thầu.Cha.Con.Cháu.Chắt');
    expect(container.textContent).toContain('không đếm dấu chấm trong mã công việc');
    expect(container.textContent).toContain('Chưa tự kiểm tra mã với danh mục ERP');
    expect(container.querySelector('details')!.open).toBe(false);
    expect(container.textContent).toContain('không dùng No. của công việc');
    for (const input of container.querySelectorAll('select, input, textarea')) {
      expect(container.querySelector(`label[for="${input.id}"]`)).not.toBeNull();
      expect(document.getElementById(input.getAttribute('aria-describedby')!)).not.toBeNull();
    }
    const mode = container.querySelector<HTMLSelectElement>('#normalization-mode')!;
    const prefix = container.querySelector<HTMLInputElement>('#normalization-prefix')!;
    await act(async () => {
      mode.value = 'initial'; mode.dispatchEvent(new Event('change', { bubbles: true }));
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!.call(prefix, 'P');
      prefix.dispatchEvent(new Event('input', { bubbles: true }));
    });
    await act(async () => { container.querySelector('button')!.click(); });
    expect(container.textContent).toContain('Chưa hoàn tất: còn dòng hoặc thiết lập cần kiểm tra');
    expect(container.textContent).toContain('Cần kiểm tra1');
    expect(container.textContent).toContain('Không lập mã ở bước này1');
    expect(container.textContent).toContain('Trạng thái và lý do cần kiểm tra');
    const level = container.querySelector<HTMLSelectElement>('#normalization-item-level')!;
    await act(async () => { level.value = '2'; level.dispatchEvent(new Event('change', { bubbles: true })); });
    expect(container.textContent).toContain('Các dòng cấp 1 nằm phía trên hạng mục');
    expect(container.querySelector('table')).toBeNull();
  } finally { await act(async () => { root.unmount(); }); }
});
