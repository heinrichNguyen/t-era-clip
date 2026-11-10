import React, { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { PopupApp } from '../../src/popup/App';
import type { TeraClipItem } from '../../src/shared/types';

const item: TeraClipItem = {
  id: 'capture-1', type: 'note', title: 'Ghi nhận cuộc họp',
  originalText: 'Nội dung gốc', sourceUrl: 'https://example.test', sourceTitle: 'Nguồn',
  createdAt: '2026-09-30T01:00:00Z', updatedAt: '2026-09-30T01:00:00Z',
  status: 'inbox', priority: 'medium', owner: '', dueDate: '', category: '', tags: [], notes: '',
};

let root: Root;
let container: HTMLDivElement;
let data: Record<string, unknown>;
let failRemoval: boolean;
const runtime: { lastError?: { message: string } } = {};

const click = async (label: string): Promise<void> => {
  const button = Array.from(container.querySelectorAll('button'))
    .find((candidate) => candidate.textContent?.trim().startsWith(label));
  expect(button).toBeDefined();
  await act(async () => { button!.click(); });
};

beforeEach(async () => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
  data = {
    teraClipItems: [item],
    teraClipSettings: { defaultReportLanguage: 'vi' },
    discussionSession: { comments: ['preserved'] },
  };
  failRemoval = false;
  delete runtime.lastError;
  vi.stubGlobal('chrome', {
    runtime,
    storage: { local: {
      get: (_keys: string[], callback: (value: unknown) => void) => callback(data),
      remove: (keys: string[], callback: () => void) => {
        if (failRemoval) {
          runtime.lastError = { message: 'Storage unavailable' };
        } else {
          keys.forEach((key) => { delete data[key]; });
        }
        callback();
        delete runtime.lastError;
      },
    } },
  });
  container = document.createElement('div');
  document.body.append(container);
  root = createRoot(container);
  await act(async () => { root.render(React.createElement(PopupApp)); });
  await click('Capture');
});

afterEach(async () => {
  await act(async () => { root.unmount(); });
  container.remove();
  vi.unstubAllGlobals();
});

describe('Capture reset', () => {
  it('requires confirmation and removes only Capture data, including filtered-out items', async () => {
    const statusFilter = container.querySelector('select')!;
    await act(async () => {
      statusFilter.value = 'done';
      statusFilter.dispatchEvent(new Event('change', { bubbles: true }));
    });
    await click('Tạo báo cáo');
    await click('Reset danh sách');
    expect(data.teraClipItems).toEqual([item]);
    await click('Xóa toàn bộ');
    expect(data.teraClipItems).toBeUndefined();
    expect(data.teraClipSettings).toEqual({ defaultReportLanguage: 'vi' });
    expect(data.discussionSession).toEqual({ comments: ['preserved'] });
    expect(statusFilter.value).toBe('all');
    expect(container.textContent).toContain('Đã xóa toàn bộ danh sách Capture.');
    expect(container.textContent).toContain('Chưa có nội dung nào được lưu.');
    expect(Array.from(container.querySelectorAll('textarea')).every((area) => area.value === '')).toBe(true);
  });

  it('keeps all data when confirmation is cancelled', async () => {
    await click('Reset danh sách');
    await click('Hủy');
    expect(data.teraClipItems).toEqual([item]);
    expect(container.querySelector('input')?.value).toBe(item.title);
  });

  it('preserves the list and reports a storage failure', async () => {
    failRemoval = true;
    await click('Reset danh sách');
    await click('Xóa toàn bộ');
    expect(data.teraClipItems).toEqual([item]);
    expect(container.querySelector('[role="alert"]')?.textContent).toContain('Không thể reset danh sách');
    expect(container.querySelector('input')?.value).toBe(item.title);
  });
});
