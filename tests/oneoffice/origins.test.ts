import { afterEach, expect, it, vi } from 'vitest';
import { getOneOfficeTabOrigin, oneOfficeOrigin, ONEOFFICE_ORIGINS } from '../../src/oneoffice/origins';
import { isAccessHistoryUrl, readAccessHistoryPage, openAccessReview } from '../../src/oneoffice/accessHistory';
import { readAccessPersonnelSource } from '../../src/oneoffice/accessPersonnel';
import { parseOneOfficeCommentApiResponse } from '../../src/oneoffice/projectCommentApi';
import type { OneOfficeProjectInventory, OneOfficeProjectTaskInventoryItem } from '../../src/oneoffice/types';
import { readProjectImportContext } from '../../src/oneoffice/projectImportWorkspace';

const production = ONEOFFICE_ORIGINS[0];
const testOrigin = ONEOFFICE_ORIGINS[2];
const task = (origin = testOrigin): OneOfficeProjectTaskInventoryItem => ({
  internalId: '7459', code: 'TEST', name: 'Test task', no: '1', depth: 1, sourceOrder: 1,
  parentNo: null, parentInternalId: null, url: `${origin}/apps/work-task-task/view?ID=7459`,
});
const inventory = (origin = testOrigin): OneOfficeProjectInventory => ({
  schemaVersion: 1, source: '1office-project-dom', capturedAt: '2026-10-21',
  project: { type: 'project', internalId: '203', code: 'P', name: 'Test', url: `${origin}/apps/work-project-project/view?ID=203` },
  tasks: [task(origin)], collection: { complete: true, warnings: [], expectedTaskCount: 1, loadedTaskCount: 1 },
});
afterEach(() => { vi.unstubAllGlobals(); vi.resetModules(); });

it.each(ONEOFFICE_ORIGINS)('recognizes project, filtered list and history on %s', async (origin) => {
  const { isSupportedOneOfficeUrl, isOneOfficeProjectPageUrl, isOneOfficeFilteredTaskListUrl } = await import('../../src/oneoffice/collectionClient');
  expect(oneOfficeOrigin(`${origin}/`)).toBe(origin);
  expect(isSupportedOneOfficeUrl(origin)).toBe(true);
  expect(isOneOfficeProjectPageUrl(`${origin}/apps/work-project-project/view?ID=203`)).toBe(true);
  expect(isOneOfficeFilteredTaskListUrl(`${origin}/apps/work-task-task/list?project_id=203`)).toBe(true);
  expect(isAccessHistoryUrl(`${origin}/apps/admin-reconciliation-histories`)).toBe(true);
});

it.each(['http://testoffice.meygroup.vn', 'https://testoffice.meygroup.vn:9444',
  'https://office.meygroup.vn:9443', 'https://testoffice.meygroup.vn.attacker.test',
  'https://testoffice.meygroup.vn@attacker.test', 'https://user:secret@testoffice.meygroup.vn:9443',
  'javascript:alert(1)', 'not a URL'])('rejects non-approved URL %s', (url) => {
  expect(() => oneOfficeOrigin(url)).toThrow();
  expect(isAccessHistoryUrl(url)).toBe(false);
});

it('rejects a source tab that switched to another supported environment', async () => {
  vi.stubGlobal('chrome', { tabs: { get: vi.fn(async () => ({ url: `${production}/` })) } });
  await expect(getOneOfficeTabOrigin(42, testOrigin)).rejects.toThrow('chuyển môi trường');
});

it('pins inventory and comment collection to the test source, including port', async () => {
  const data = inventory();
  const executeScript = vi.fn(async ({ func }: { func: Function }) => [{ result:
    func.name === 'collectOneOfficeProjectInventoryFromApi' ? data : { total: 0, items: [] } }]);
  vi.stubGlobal('chrome', { tabs: { get: vi.fn(async () => ({ url: data.project.url })) }, scripting: { executeScript } });
  const { collectOneOfficeProjectInventory } = await import('../../src/oneoffice/collectionClient');
  expect(await collectOneOfficeProjectInventory(42)).toEqual(data);
  expect(executeScript.mock.calls[0]).toEqual([expect.objectContaining({ args: [{ allowedOrigin: testOrigin }] })]);
  const { collectOneOfficeProjectTaskDiscussions } = await import('../../src/oneoffice/projectBatchCollector');
  const result = await collectOneOfficeProjectTaskDiscussions(42, data, vi.fn(), () => false);
  expect(result.tasks[0].state).toBe('complete');
  expect(result.tasks[0].discussion?.entity.url).toBe(task().url);
  expect(executeScript.mock.calls[1]).toEqual([expect.objectContaining({ args: [expect.objectContaining({ allowedOrigin: testOrigin, taskId: '7459' })] })]);
});

it('refuses mixed-origin tasks before sending any request', async () => {
  const executeScript = vi.fn();
  vi.stubGlobal('chrome', { scripting: { executeScript } });
  const { collectOneOfficeProjectTaskDiscussions } = await import('../../src/oneoffice/projectBatchCollector');
  await expect(collectOneOfficeProjectTaskDiscussions(42, { ...inventory(), tasks: [task(), task(production)] }, vi.fn(), () => false)).rejects.toThrow('nhiều môi trường');
  expect(executeScript).not.toHaveBeenCalled();
});

it('keeps exported comment links on their source environment', () => {
  const result = parseOneOfficeCommentApiResponse({ total: 0, items: [] }, task());
  expect(result?.entity.url).toBe(task().url);
});

it('pins the access review URL to the originating site', async () => {
  const create = vi.fn();
  vi.stubGlobal('chrome', { tabs: { get: vi.fn(async () => ({ url: `${testOrigin}/apps/admin-reconciliation-histories` })), create },
    runtime: { getURL: (path: string) => `chrome-extension://example/${path}` } });
  await openAccessReview(42);
  expect(new URL(create.mock.calls[0][0].url).searchParams.get('sourceOrigin')).toBe(testOrigin);
});

it('blocks history and personnel if the source changes environment, even to another allowed site', async () => {
  vi.stubGlobal('location', { origin: production, pathname: '/apps/admin-reconciliation-histories' });
  const fetch = vi.fn(); vi.stubGlobal('fetch', fetch);
  expect((await readAccessHistoryPage(1, true, null, testOrigin)).ok).toBe(false);
  expect((await readAccessPersonnelSource({ kind: 'organization' }, testOrigin)).ok).toBe(false);
  expect(fetch).not.toHaveBeenCalled();
});

it('uses test-origin GET requests for history and personnel without following redirects', async () => {
  vi.stubGlobal('location', { origin: testOrigin, pathname: '/apps/admin-reconciliation-histories' });
  const fetch = vi.fn(async () => ({ ok: false, status: 403 })); vi.stubGlobal('fetch', fetch);
  await readAccessHistoryPage(1, true, null, testOrigin);
  await readAccessPersonnelSource({ kind: 'organization' }, testOrigin);
  expect(fetch).toHaveBeenCalledTimes(2);
  for (const [url, options] of fetch.mock.calls as unknown as Array<[URL, RequestInit]>) {
    expect(new URL(url).origin).toBe(testOrigin);
    expect(options.redirect).toBe('error');
    expect(options.method ?? 'GET').toBe('GET');
  }
});

it('rejects a different environment even when snapshot project ID and capture time are identical', async () => {
  const data = inventory(production);
  vi.stubGlobal('chrome', { runtime: {}, storage: { session: { get: (_: unknown, done: Function) => done({ oneOfficeNoReview: data }) } } });
  const query = new URLSearchParams({ projectId: '203', capturedAt: data.capturedAt, sourceOrigin: testOrigin });
  await expect(readProjectImportContext(query.toString())).rejects.toThrow('đã thay đổi');
});

it('does not enable the production-only form runner on the new test site', async () => {
  const executeScript = vi.fn();
  vi.stubGlobal('chrome', { scripting: { executeScript }, tabs: { get: vi.fn(async () => ({ url: `${testOrigin}/apps/work-project-project/view?ID=870` })), update: vi.fn() } });
  const { runNoWriteExperiment } = await import('../../src/oneoffice/noWriteExperiment');
  await expect(runNoWriteExperiment(42, true)).rejects.toThrow('ID 870');
  expect(executeScript).not.toHaveBeenCalled();
});

it('rejects a test-site inventory before creating any production write journal', async () => {
  const executeScript = vi.fn(); const set = vi.fn();
  vi.stubGlobal('chrome', { scripting: { executeScript }, storage: { local: { set } } });
  const data = inventory();
  data.project.internalId = '870';
  data.project.url = `${testOrigin}/apps/work-project-project/view?ID=870`;
  data.tasks = [{ ...task(), internalId: '100119', code: 'TEST01.3', no: '1.3', sourceNo: '1.3' }];
  const { runNoUpdatePlan } = await import('../../src/oneoffice/noWriteExperiment');
  await expect(runNoUpdatePlan(42, data, 'TEST01.3\t1.4')).rejects.toThrow('Chưa bật ghi');
  expect(executeScript).not.toHaveBeenCalled(); expect(set).not.toHaveBeenCalled();
});
