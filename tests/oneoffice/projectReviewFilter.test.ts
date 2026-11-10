import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { expect, it, vi } from 'vitest';
import { ProjectReviewApp } from '../../src/project-review/App';
import type { OneOfficeProjectBatchResult } from '../../src/oneoffice/types';

vi.mock('../../src/oneoffice/projectBatchSession', () => ({
  getOneOfficeProjectBatchSession: async () => ({ sourceTab: { id: 1 }, inventory: { project: { name: 'Test', url: 'https://office.meygroup.vn/apps/work-project-project/view?ID=870' }, tasks: [], collection: { complete: true, warnings: ['24 duplicate task No. value(s) were found.'] } } }),
}));
vi.mock('../../src/oneoffice/projectBatchCollector', () => ({
  collectOneOfficeProjectTaskDiscussions: vi.fn(),
}));
vi.mock('../../src/oneoffice/export/projectXlsx', () => ({
  createOneOfficeProjectXlsxFilename: () => 'test.xlsx',
  downloadOneOfficeProjectXlsx: vi.fn().mockResolvedValue('test.xlsx'),
}));
import { collectOneOfficeProjectTaskDiscussions } from '../../src/oneoffice/projectBatchCollector';
import { downloadOneOfficeProjectXlsx } from '../../src/oneoffice/export/projectXlsx';

it('updates filtered count, exports the filter and resets without recollecting', async () => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
  vi.stubGlobal('chrome', { runtime: { getURL: (path: string) => path }, storage: { session: {
    remove: vi.fn((_keys, done) => done()), set: vi.fn((_value, done) => done()),
  } } });
  const confirm = vi.spyOn(window, 'confirm').mockReturnValue(true);
  const batch = {
    tasks: [{ task: { internalId: '1', name: 'Task', no: '1' }, state: 'complete', discussion: {
      collection: { loadedRootCommentCount: 1, loadedReplyCount: 1 },
      comments: [
        { id: '1', parentId: null, authorName: 'A', createdAtIso: '2026-09-30T00:00:00+07:00' },
        { id: '2', parentId: '1', authorName: 'B', createdAtIso: '2026-10-01T00:00:00+07:00' },
      ],
    } }], warnings: [],
  } as unknown as OneOfficeProjectBatchResult;
  vi.mocked(collectOneOfficeProjectTaskDiscussions).mockResolvedValue(batch);
  const container = document.createElement('div');
  document.body.append(container);
  const root = createRoot(container);
  const click = async (text: string) => {
    const button = [...container.querySelectorAll('button')].find((item) => item.textContent === text)!;
    expect(button).toBeDefined();
    await act(async () => button.click());
  };
  try {
    await act(async () => root.render(React.createElement(ProjectReviewApp)));
    expect(container.textContent).toContain('2 bình luận/phản hồi thuộc 1 công việc');
    expect(container.textContent).toContain('Cảnh báo danh sách công việc (1)');
    expect(container.textContent).toContain('Có 24 giá trị No. bị trùng.');
    const kind = container.querySelectorAll('select')[1];
    await act(async () => { kind.value = 'replies'; kind.dispatchEvent(new Event('change', { bubbles: true })); });
    expect(container.textContent).toContain('1 bình luận/phản hồi thuộc 1 công việc');
    await click('Xuất Excel');
    await click('Tải file Excel');
    expect(downloadOneOfficeProjectXlsx).toHaveBeenCalledWith(batch, expect.objectContaining({ kind: 'replies' }));
    await click('Xóa bộ lọc');
    expect(container.textContent).toContain('2 bình luận/phản hồi thuộc 1 công việc');
    expect(collectOneOfficeProjectTaskDiscussions).toHaveBeenCalledTimes(1);
    await click('Xóa dữ liệu tạm');
    expect(container.textContent).toContain('Đã xóa dữ liệu tạm');
    expect(container.textContent).not.toContain('Xuất Excel');
    expect(container.textContent).not.toContain('Bộ lọc thảo luận');
  } finally {
    await act(async () => root.unmount());
    container.remove();
    confirm.mockRestore();
    vi.unstubAllGlobals();
  }
});
