import { beforeEach, describe, expect, it } from 'vitest';
import {
  buildOneOfficeFilteredTaskInventory,
  collectOneOfficeFilteredTaskSelectionFromPage,
} from '../../src/oneoffice/filteredTaskInventory';
import type { OneOfficeProjectInventory } from '../../src/oneoffice/types';

const filter = JSON.stringify({
  ors: [[
    { f: 'status', o: 'in', p: ['DOING'] },
    { f: 'project_id', o: 'in', p: ['258'] },
  ]],
});

const pageHtml = (page: number, totalPages: number, totalTasks: number, id: number): string => `
  <div class="table-pages">
    <div class="table-page-title">Hiển thị ${id} - ${id} / ${totalTasks} bản ghi</div>
    <div>Trang: 0${page} / 0${totalPages}</div>
  </div>
  <table><tbody><tr class="table-row-body">
    <td col="code">CODE.${id}</td>
    <td col="title"><a class="table-main-title" href="/apps/work-normal-normal/view?ID=${id}">Task ${id}</a></td>
    <td col="project_title">Filtered Project</td>
    <td col="status">Đang thực hiện</td>
  </tr></tbody></table>
`;

beforeEach(() => {
  window.history.replaceState(
    {},
    '',
    `/apps/work-task-task/list?@sfilter=${encodeURIComponent(filter)}&p=1`,
  );
  document.body.innerHTML = pageHtml(1, 3, 3, 101);
});

describe('collectOneOfficeFilteredTaskSelectionFromPage', () => {
  it('loads every page in the current project filter', async () => {
    const visitedPages: number[] = [1];
    const renderPage = (page: number): void => {
      document.body.innerHTML = `${pageHtml(page, 3, 3, 100 + page)}
        <a class="table-header-btn icon-caret-left${page === 1 ? ' disabled' : ''}"></a>
        <a class="table-header-btn icon-caret-right${page === 3 ? ' disabled' : ''}"></a>`;
      window.history.replaceState({}, '', `/apps/work-task-task/list?@sfilter=${encodeURIComponent(filter)}&p=${page}`);
      document.querySelector('.icon-caret-left')?.addEventListener('click', (event) => {
        event.preventDefault();
        if (page > 1) {
          visitedPages.push(page - 1);
          renderPage(page - 1);
        }
      });
      document.querySelector('.icon-caret-right')?.addEventListener('click', (event) => {
        event.preventDefault();
        if (page < 3) {
          visitedPages.push(page + 1);
          renderPage(page + 1);
        }
      });
    };
    renderPage(1);

    const selection = await collectOneOfficeFilteredTaskSelectionFromPage({
      allowedOrigin: window.location.origin,
    });

    expect(visitedPages).toEqual([1, 2, 3, 2, 1]);
    expect(selection).toMatchObject({
      projectInternalId: '258',
      projectName: 'Filtered Project',
      expectedTaskCount: 3,
      complete: true,
      warnings: [],
    });
    expect(selection.tasks.map((task) => task.internalId)).toEqual(['101', '102', '103']);
  });

  it('waits until delayed table rows finish rendering before collecting a page', async () => {
    const rows = (ids: number[]): string => ids.map((id) => `
      <tr class="table-row-body">
        <td col="code">CODE.${id}</td>
        <td col="title"><a class="table-main-title" href="/apps/work-normal-normal/view?ID=${id}">Task ${id}</a></td>
        <td col="project_title">Filtered Project</td>
      </tr>
    `).join('');
    const render = (page: number, ids: number[]): void => {
      const start = page === 1 ? 1 : 3;
      const end = page === 1 ? 2 : 4;
      document.body.innerHTML = `
        <div class="table-pages">
          <div class="table-page-title">Hiển thị ${start} - ${end} / 4 bản ghi</div>
          <div>Trang: 0${page} / 02</div>
        </div>
        <table><tbody>${rows(ids)}</tbody></table>
        <a class="table-header-btn icon-caret-left${page === 1 ? ' disabled' : ''}"></a>
        <a class="table-header-btn icon-caret-right${page === 2 ? ' disabled' : ''}"></a>`;
      document.querySelector('.icon-caret-left')?.addEventListener('click', (event) => {
        event.preventDefault();
        if (page === 2) {
          render(1, [201, 202]);
        }
      });
      document.querySelector('.icon-caret-right')?.addEventListener('click', (event) => {
        event.preventDefault();
        if (page === 1) {
          render(2, [201, 202]);
          globalThis.setTimeout(() => render(2, [203, 202]), 100);
          globalThis.setTimeout(() => render(2, [203, 204]), 250);
        }
      });
    };
    render(1, [201, 202]);

    const selection = await collectOneOfficeFilteredTaskSelectionFromPage({
      allowedOrigin: window.location.origin,
    });

    expect(selection.complete).toBe(true);
    expect(selection.warnings).toEqual([]);
    expect(selection.tasks.map((task) => task.internalId)).toEqual(['201', '202', '203', '204']);
  });

  it('rejects a task list that is not limited to one project', async () => {
    window.history.replaceState({}, '', '/apps/work-task-task/list?tab=all');

    await expect(
      collectOneOfficeFilteredTaskSelectionFromPage({
        allowedOrigin: window.location.origin,
      }),
    ).rejects.toThrow('exactly one project');
  });
});

describe('buildOneOfficeFilteredTaskInventory', () => {
  it('keeps filter order while reusing detailed project metadata', () => {
    const projectInventory: OneOfficeProjectInventory = {
      schemaVersion: 1,
      source: '1office-project-api',
      capturedAt: '2026-09-29T00:00:00.000Z',
      project: {
        type: 'project',
        internalId: '258',
        code: null,
        name: 'Project',
        url: `${window.location.origin}/apps/work-project-project/view?ID=258`,
      },
      collection: { complete: true, expectedTaskCount: 2, loadedTaskCount: 2, warnings: [] },
      tasks: [
        {
          sourceOrder: 0, internalId: '101', no: '1.1', code: 'A', name: 'First',
          url: `${window.location.origin}/apps/work-task-task/view?ID=101`, depth: 2,
          parentNo: '1', parentInternalId: '100',
        },
        {
          sourceOrder: 1, internalId: '102', no: '1.2', code: 'B', name: 'Second',
          url: `${window.location.origin}/apps/work-task-task/view?ID=102`, depth: 2,
          parentNo: '1', parentInternalId: '100',
        },
      ],
    };
    const selection = {
      sourceUrl: `${window.location.origin}/apps/work-task-task/list?project_id=258`,
      projectInternalId: '258',
      projectName: 'Filtered Project',
      expectedTaskCount: 2,
      complete: true,
      warnings: [],
      tasks: [projectInventory.tasks[1], projectInventory.tasks[0]],
    };

    const inventory = buildOneOfficeFilteredTaskInventory(selection, projectInventory);

    expect(inventory.project.name).toBe('Filtered Project');
    expect(inventory.project.url).toBe(selection.sourceUrl);
    expect(inventory.tasks.map((task) => [task.internalId, task.no, task.sourceOrder])).toEqual([
      ['102', '1.2', 0],
      ['101', '1.1', 1],
    ]);
  });
});
