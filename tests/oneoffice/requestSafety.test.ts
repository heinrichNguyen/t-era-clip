import { afterEach, describe, expect, it, vi } from 'vitest';
import { fetchOneOfficeCommentsFromPage } from '../../src/oneoffice/projectCommentApi';
import { probeOneOfficeQuickViewFromPage } from '../../src/oneoffice/projectFastPathProbe';
import { collectOneOfficeProjectInventoryFromApi } from '../../src/oneoffice/projectInventoryApi';
import { readAccessHistoryPage } from '../../src/oneoffice/accessHistory';
import { readAccessPersonnelSource } from '../../src/oneoffice/accessPersonnel';

afterEach(() => { vi.useRealTimers(); vi.unstubAllGlobals(); vi.resetModules(); });

describe('30 second request deadline', () => {
  it.each(['comments', 'quickview', 'inventory', 'history', 'personnel'])('aborts %s and reports a retryable error', async (kind) => {
    vi.useFakeTimers();
    vi.stubGlobal('location', { origin: 'https://office.meygroup.vn', pathname: '/apps/admin-reconciliation-histories' });
    const fetch = vi.fn((_url: string, options: RequestInit) => new Promise((_resolve, reject) => {
      options.signal!.addEventListener('abort', () => reject(new Error('aborted')), { once: true });
    }));
    vi.stubGlobal('fetch', fetch);
    const options = { allowedOrigin: window.location.origin, taskId: '1', offset: 0 };
    let completed = false;
    const request = kind === 'comments' ? fetchOneOfficeCommentsFromPage(options)
      : kind === 'quickview' ? probeOneOfficeQuickViewFromPage(options)
      : kind === 'history' ? readAccessHistoryPage(1, true)
      : kind === 'personnel' ? readAccessPersonnelSource({ kind: 'directory', tab: 'active', page: 1 })
      : collectOneOfficeProjectInventoryFromApi({ allowedOrigin: window.location.origin,
          sourceUrl: `${window.location.origin}/apps/work-project-project/view?ID=1` });
    const outcome = request.then((value) => { completed = true; return JSON.stringify(value); },
      (error) => { completed = true; return error.message as string; });
    await vi.advanceTimersByTimeAsync(29999);
    expect(completed).toBe(false);
    expect(fetch).toHaveBeenCalledTimes(1);
    await vi.advanceTimersByTimeAsync(1);
    expect(await outcome).toContain('30 giây');
    expect(await outcome).toContain('thử lại');
    expect(fetch.mock.calls[0][1].signal!.aborted).toBe(true);
    expect(fetch.mock.calls[0][1].redirect).toBe('error');
    expect(vi.getTimerCount()).toBe(0);
  });

  it.each([401, 403, 429])('reports HTTP %s without returning comment data', async (status) => {
    vi.stubGlobal('fetch', vi.fn(async () => new Response('{}', { status })));
    const options = { allowedOrigin: window.location.origin, taskId: '1', offset: 0 };
    expect((await fetchOneOfficeCommentsFromPage(options))?.requestError).toContain(`HTTP ${status}`);
    expect(await probeOneOfficeQuickViewFromPage(options)).toMatchObject({ requestFailed: true, available: false });
  });

  it('marks a timed-out task failed without starting quickview or DOM fallback', async () => {
    const executeScript = vi.fn(async () => [{ result: { requestError: 'Yêu cầu 1Office: máy chủ không phản hồi trong 30 giây. Hãy thử lại.' } }]);
    vi.stubGlobal('chrome', { scripting: { executeScript }, tabs: { get: vi.fn(async () => ({ url: 'https://office.meygroup.vn/apps/work-project-project/view?ID=1' })) } });
    const { collectOneOfficeProjectTaskDiscussions } = await import('../../src/oneoffice/projectBatchCollector');
    const task = { internalId: '1', name: 'Task', no: '1', code: 'T1', depth: 1, sourceOrder: 0,
      parentNo: null, parentInternalId: null, url: 'https://office.meygroup.vn/apps/work-task-task/view?ID=1' };
    const inventory = { schemaVersion: 1 as const, source: '1office-project-dom' as const, capturedAt: '',
      project: { type: 'project' as const, internalId: '1', code: 'P', name: 'P', url: 'https://office.meygroup.vn/apps/work-project-project/view?ID=1' },
      collection: { complete: true, expectedTaskCount: 1, loadedTaskCount: 1, warnings: [] }, tasks: [task] };
    const result = await collectOneOfficeProjectTaskDiscussions(42, inventory, vi.fn(), () => false);
    expect(result.tasks[0]).toMatchObject({ state: 'failed', discussion: null });
    expect(result.tasks[0].error).toContain('30 giây');
    expect(executeScript).toHaveBeenCalledTimes(1);
  });
});
