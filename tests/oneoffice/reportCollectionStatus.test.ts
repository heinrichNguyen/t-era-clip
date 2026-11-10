import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { expect, it, vi } from 'vitest';
import { ProjectReviewApp } from '../../src/project-review/App';

const task = {internalId: '', collectionKey: 'report-row:0', name: 'Task', no: '1'};
vi.mock('../../src/oneoffice/projectBatchSession', () => ({
  getOneOfficeProjectBatchSession: async () => ({sourceTab: {id: 1}, inventory: {
    project: {type: 'report', name: 'Report', url: 'https://office.meygroup.vn/apps/report-report/view?ID=74'},
    tasks: [], collection: {complete: true, warnings: []},
  }}),
}));
vi.mock('../../src/oneoffice/projectBatchCollector', () => ({
  collectOneOfficeProjectTaskDiscussions: async () => ({tasks: [{task, state: 'unresolved', discussion: null,
    error: 'Chưa xác minh ID'}], warnings: []}),
}));
vi.mock('../../src/oneoffice/export/projectXlsx', () => ({
  createOneOfficeProjectXlsxFilename: () => 'test.xlsx', downloadOneOfficeProjectXlsx: vi.fn(),
}));

it('does not call an all-unresolved report complete or promise task-page navigation', async () => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
  vi.stubGlobal('chrome', {runtime: {getURL: (path: string) => path}});
  const container = document.createElement('div');
  document.body.append(container);
  const root = createRoot(container);
  try {
    await act(async () => root.render(React.createElement(ProjectReviewApp)));
    const header = container.querySelector('header')!;
    expect(header.textContent).toContain('chưa lấy được thảo luận của công việc nào');
    expect(header.textContent).not.toContain('Đã hoàn tất thu thập');
    expect(header.textContent).toContain('không mở lần lượt trang công việc');
    expect(header.textContent).not.toContain('phương án dự phòng');
    expect(header.textContent).toContain('1Chưa lấy được thảo luận');
  } finally {
    await act(async () => root.unmount());
    container.remove();
    vi.unstubAllGlobals();
  }
});
