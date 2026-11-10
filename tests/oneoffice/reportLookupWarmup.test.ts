import {afterEach, expect, it, vi} from 'vitest';
import type {OneOfficeProjectTaskInventoryItem} from '../../src/oneoffice/types';

const task: OneOfficeProjectTaskInventoryItem = {sourceOrder: 0, collectionKey: 'report-row:0', internalId: '',
  no: '', code: 'CODE', name: 'Task', url: '', depth: 0, parentNo: null, parentInternalId: null};
const origin = 'https://office.meygroup.vn';
const sourceUrl = `${origin}/apps/report-report/view?ID=74`;
afterEach(() => {vi.unstubAllGlobals(); vi.resetModules();});

it('shares the first discovery and supplies verified widget hints to waiting rows, scoped to the origin', async () => {
  let release!: () => void;
  const gate = new Promise<void>(resolve => {release = resolve;});
  const executeScript = vi.fn().mockImplementation(async ({args: [options]}) => {
    if (executeScript.mock.calls.length === 1) await gate;
    return [{result: {...options.task, internalId: '123', url: `${new URL(options.sourceUrl).origin}/apps/work-normal-normal/view?ID=123`,
      identityWidget: {tabId: 2184, widgetId: 47091}}}];
  });
  vi.stubGlobal('chrome', {tabs: {get: async (tab: number) => ({url: tab === 5 ? sourceUrl : 'https://testoffice.meygroup.vn/apps/report-report/view?ID=74'})},
    scripting: {executeScript}});
  const {resolveReportTask} = await import('../../src/oneoffice/reportTaskLookup');
  const first = resolveReportTask(5, sourceUrl, task);
  const second = resolveReportTask(5, sourceUrl, task);
  await vi.waitFor(() => expect(executeScript).toHaveBeenCalledTimes(1));
  release();
  const results = await Promise.all([first, second]);
  expect(executeScript.mock.calls[1][0].args[0].widgetHints).toEqual({'work-normal-normal': {tabId: 2184, widgetId: 47091}});
  expect(results[0]).not.toHaveProperty('identityWidget');
  await resolveReportTask(6, 'https://testoffice.meygroup.vn/apps/report-report/view?ID=74', task);
  expect(executeScript.mock.calls[2][0].args[0].widgetHints).toEqual({});
});

it('does not serialize every remaining lookup when the first discovery fails', async () => {
  const executeScript = vi.fn().mockResolvedValueOnce([{result: {...task, identityError: 'Missing task'}}]);
  vi.stubGlobal('chrome', {tabs: {get: async () => ({url: sourceUrl})}, scripting: {executeScript}});
  const {resolveReportTask} = await import('../../src/oneoffice/reportTaskLookup');
  await resolveReportTask(5, sourceUrl, task);
  let release!: () => void;
  const gate = new Promise<void>(resolve => {release = resolve;});
  executeScript.mockImplementation(async () => {await gate; return [{result: {...task, identityError: 'Missing task'}}];});
  const running = [resolveReportTask(5, sourceUrl, task), resolveReportTask(5, sourceUrl, task)];
  await vi.waitFor(() => expect(executeScript).toHaveBeenCalledTimes(3));
  release();
  await Promise.all(running);
});
