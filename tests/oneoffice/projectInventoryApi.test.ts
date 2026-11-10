import { beforeEach, describe, expect, it, vi } from 'vitest';
import { collectOneOfficeProjectInventoryFromApi } from '../../src/oneoffice/projectInventoryApi';

const sourceUrl = `${window.location.origin}/apps/work-project-project/view?ID=248&name=visual-work&tab=work`;

const jsonResponse = (data: unknown): Response =>
  ({
    ok: true,
    status: 200,
    json: async () => data,
  }) as Response;

beforeEach(() => {
  document.body.innerHTML = `
    <h1 id="header-title">Large Project</h1>
    <a href="/apps/work-project-project/view?ID=248">Công việc (414)</a>
  `;
  vi.restoreAllMocks();
});

describe('collectOneOfficeProjectInventoryFromApi', () => {
  it('reports missing field coverage even when all task IDs and enrichment records were returned', async () => {
    document.querySelector('a')!.textContent = 'Công việc (2)';
    vi.stubGlobal('fetch', vi.fn(async (_url: string, init?: RequestInit) => init?.method === 'POST'
      ? jsonResponse({ enriched: { 100: { fields: { status: 'Chờ thực hiện' } }, 101: { fields: { cf19: '' } } } })
      : jsonResponse({ tasks: [{ ID: 100, title: 'Parent' }, { ID: 101, parent_id: 100, title: 'Child' }], phases: [] })));
    const inventory = await collectOneOfficeProjectInventoryFromApi({ allowedOrigin: window.location.origin, sourceUrl });
    expect(inventory.collection.complete).toBe(true);
    expect(inventory.tasks[0].sourceNo).toBeUndefined();
    expect(inventory.tasks[1].sourceNo).toBe('');
    expect(inventory.collection.warnings).toEqual(expect.arrayContaining([
      expect.stringContaining('No. gốc của 1/2'), expect.stringContaining('mã công việc của 2/2'),
    ]));
    expect(inventory.collection.warnings.join(' ')).toContain('Không kết luận mã trên 1Office đang trống');
  });
  it('loads 414 virtualized tasks and enriches them in batches of at most 50', async () => {
    const tasks = Array.from({ length: 414 }, (_, index) => {
      const id = index + 1;
      return {
        ID: id,
        parent_id: id >= 2 && id <= 6 ? id - 1 : null,
        phase_id: 407,
        title: `Task ${String(id).padStart(3, '0')}`,
        code: `TASK.${id}`,
      };
    });
    const requestedBatches: string[][] = [];
    const fetchMock = vi.fn(async (_url: string, init?: RequestInit) => {
      if (init?.method !== 'POST') {
        return jsonResponse({
          tasks,
          phases: [{ phase_db_id: 407 }],
          groupType: 'phase',
        });
      }

      const body = new URLSearchParams(String(init.body));
      const ids = JSON.parse(body.get('ids') ?? '[]') as string[];
      requestedBatches.push(ids);
      return jsonResponse({
        enriched: Object.fromEntries(ids.map((id) => [id, { fields: { cf19: '' } }])),
      });
    });
    vi.stubGlobal('fetch', fetchMock);

    const inventory = await collectOneOfficeProjectInventoryFromApi({
      allowedOrigin: window.location.origin,
      sourceUrl,
    });

    expect(inventory.source).toBe('1office-project-api');
    expect(inventory.collection).toMatchObject({
      complete: true,
      expectedTaskCount: 414,
      loadedTaskCount: 414,
      warnings: [],
    });
    expect(inventory.tasks).toHaveLength(414);
    expect(requestedBatches).toHaveLength(9);
    expect(requestedBatches.every((batch) => batch.length <= 50)).toBe(true);
    expect(requestedBatches.flat()).toHaveLength(414);
    expect(inventory.tasks.slice(0, 6).map((task) => task.no)).toEqual([
      '1.1',
      '1.1.1',
      '1.1.1.1',
      '1.1.1.1.1',
      '1.1.1.1.1.1',
      '1.1.1.1.1.1.1',
    ]);
    expect(inventory.tasks[5]).toMatchObject({
      internalId: '6',
      parentInternalId: '5',
      parentNo: '1.1.1.1.1.1',
    });
  });

  it('uses the enriched No. value and falls back to the generated outline when blank', async () => {
    document.querySelector('a')!.textContent = 'Công việc (2)';
    const fetchMock = vi.fn(async (_url: string, init?: RequestInit) => {
      if (init?.method !== 'POST') {
        return jsonResponse({
          tasks: [
            { ID: 100, parent_id: null, phase_id: 407, title: 'Parent', code: 'P' },
            { ID: 101, parent_id: 100, phase_id: 407, title: 'Child', code: 'C' },
          ],
          phases: [{ phase_db_id: 407 }],
        });
      }

      return jsonResponse({
        enriched: {
          100: { fields: { cf19: '1.10' } },
          101: { fields: { cf19: '' } },
        },
      });
    });
    vi.stubGlobal('fetch', fetchMock);

    const inventory = await collectOneOfficeProjectInventoryFromApi({
      allowedOrigin: window.location.origin,
      sourceUrl,
    });

    expect(inventory.collection.complete).toBe(true);
    expect(inventory.tasks[0]).toMatchObject({ no: '1.10', sourceNo: '1.10', parentNo: null });
    expect(inventory.tasks[1]).toMatchObject({
      no: '1.1.1',
      sourceNo: '',
      parentNo: '1.10',
      parentInternalId: '100',
    });
  });

  it('collects confirmed task export fields from structure and enrichment responses', async () => {
    document.querySelector('a')!.textContent = 'Công việc (1)';
    const fetchMock = vi.fn(async (_url: string, init?: RequestInit) => {
      if (_url.includes('/app/autofill/users/items')) {
        return jsonResponse({
          items: [
            { ID: 1, name: 'Người A' },
            { ID: 2, name: 'Người B' },
            { ID: 3, name: 'Người giao A' },
            { ID: 4, name: 'Người giao B' },
          ],
        });
      }
      if (init?.method !== 'POST') {
        return jsonResponse({
          tasks: [{
            ID: 100,
            parent_id: null,
            phase_id: 407,
            title: 'Task with details',
            code: 'TASK.100',
            start_plan: '2026-10-01 00:00:00',
            end_plan: '2026-10-10 00:00:00',
            start: '2026-10-02 00:00:00',
            end: '2026-10-11 00:00:00',
            status: 'COMPLETED',
            out_date: 2,
            fields: {
              start_plan: '01/10/2026',
              end_plan: '10/10/2026',
              start: '02/10/2026',
              end: '11/10/2026',
              status: 'Hoàn thành',
              out_date: 2,
            },
          }],
          phases: [{ phase_db_id: 407 }],
        });
      }

      return jsonResponse({
        enriched: {
          100: {
            assign_names_html:
              `<elem name='users' data-props='{"user":{"ID":1,"name":"Người A"},"num":2,"ids":["1","2"]}'></elem>`,
            fields: {
              cf19: '1.1',
              owner_ids:
                `<elem name='datadisplay.users' data-props='{"mode":"list","content":"3,4","value":"3,4"}'></elem>`,
              department_ids: 'Ban A; Ban B',
              status:
                `<elem name='chip' data-props='{"label":"Hoàn thành"}'></elem>`,
              cf171: '29/09/2026',
              cf174: '09/10/2026',
              out_date: '2',
              out_date_completed: '1',
              tag_ids:
                `<elem name='tags' data-props='{"items":[{"label":"Nhãn A"},{"label":"Nhãn B"}]}'></elem>`,
            },
          },
        },
      });
    });
    vi.stubGlobal('fetch', fetchMock);

    const inventory = await collectOneOfficeProjectInventoryFromApi({
      allowedOrigin: window.location.origin,
      sourceUrl,
    });

    expect(inventory.tasks[0]).toMatchObject({
      assigneeNames: ['Người A', 'Người B'],
      assignerNames: ['Người giao A', 'Người giao B'],
      departmentNames: ['Ban A', 'Ban B'],
      taskStatus: 'Hoàn thành',
      baselineStart: '29/09/2026',
      baselineEnd: '09/10/2026',
      plannedStart: '01/10/2026',
      plannedEnd: '10/10/2026',
      actualStart: '02/10/2026',
      actualEnd: '11/10/2026',
      overdueDays: '2',
      completedOverdueDays: '1',
      labels: ['Nhãn A', 'Nhãn B'],
    });
  });

  it('keeps the task inventory usable when optional detailed fields are missing', async () => {
    document.querySelector('a')!.textContent = 'Công việc (1)';
    vi.stubGlobal(
      'fetch',
      vi.fn(async (_url: string, init?: RequestInit) =>
        init?.method === 'POST'
          ? jsonResponse({ enriched: {} })
          : jsonResponse({
              tasks: [{ ID: 100, parent_id: null, phase_id: 407, title: 'Parent' }],
              phases: [{ phase_db_id: 407 }],
            }),
      ),
    );

    const inventory = await collectOneOfficeProjectInventoryFromApi({
      allowedOrigin: window.location.origin,
      sourceUrl,
    });

    expect(inventory.collection.complete).toBe(true);
    expect(inventory.collection.warnings).toContain(
      'Loaded detailed task fields for 0 of 1 tasks.',
    );
  });
});
