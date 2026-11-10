import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { lookupReportTaskFromPage } from '../../src/oneoffice/reportTaskLookup';
import type { OneOfficeProjectTaskInventoryItem } from '../../src/oneoffice/types';

const task: OneOfficeProjectTaskInventoryItem = { sourceOrder: 0, collectionKey: 'report-row:0', internalId: '', url: '',
  code: 'CODE.18526', name: 'Công việc A', no: '1', depth: 0, parentInternalId: null, parentNo: null };
const path = '/apps/report-report/view?ID=74&from=01/10/2026&to=05/11/2026&department_ids[]=424';
beforeEach(() => window.history.replaceState({}, '', path));
afterEach(() => vi.unstubAllGlobals());
const options = () => ({sourceUrl: location.href, task});
const detail = '<div class="detail-field"><div class="detail-field-label">Mã công việc</div><div class="detail-field-content">CODE.18526</div></div>';
const found = '<div class="table-page-title">Hiển thị 1 - 1 / 1 bản ghi</div><table><tr class="table-row-body"><td><a class="table-main-title" href="/apps/work-normal-normal/view?ID=89876">Công việc A</a></td></tr></table>';
const response = (text: string) => ({ok: true, text: async () => text});
const elem = (name: string, props: unknown, extra = '') => `<elem name="${name}" data-props="${JSON.stringify(props).replace(/&/g, '&amp;').replace(/"/g, '&quot;')}" ${extra}></elem>`;
const declaredLink = (id: number, name: string, main = true) => `<elem name="a" data-atts="${JSON.stringify({href: `/apps/work-normal-normal/view?ID=${id}`, class: main ? 'table-main-title d-block' : 'table-sub-title'}).replace(/"/g, '&quot;')}" data-children="${JSON.stringify([name]).replace(/"/g, '&quot;')}"></elem>`;
const liveShapeSearch = () => elem('table', {object: 'work-task-task', page: {total: 1, params: {s: task.code}}, vars: {s: task.code},
  rows: [{ID: 89876, vals: {title: task.name}, fields: {title: `<span>${declaredLink(89876, task.name)}${declaredLink(26481, 'CV cha: Parent', false)}</span>`},
    url: 'apps/work-normal-normal/view?ID=89876'}]});
const boardShell = () => elem('board', {object: 'work-normal-normal', mode: 'view', page: 'detail', module: 'work',
  postId: 988, post: {ID: 89876}, curentTabId: 2184, tabs: [{ID: 2184, object: 'work-normal-normal', module: 'work'}]});
const widgetBody = (id = 89876, code = task.code) => JSON.stringify({content: elem('visual.widget', {
  name: 'Work-Task-TaskNormalInformation', postId: id, board_id: 47091, url: '/apps/work-normal-normal/view/widget', isAjax: false,
  body: `<elem name="div"><template children>${elem('detail.field', {label: 'Mã công việc', content: code})}</template></elem>`,
})});
const templateLink = (id: number, name: string) => `<elem name="a" data-atts="${JSON.stringify({href: `/apps/work-normal-normal/view?ID=${id}`, class: 'table-main-title d-block'}).replace(/"/g, '&quot;')}"><template children><span>${name}</span></template></elem>`;
const searchRows = (names: string[], total = names.length, query = task.code) => elem('table', {object: 'work-task-task', vars: {s: query}, page: {total},
  rows: names.map((name, index) => ({ID: 89876 + index, vals: {title: name}, fields: {title: templateLink(89876 + index, name)},
    url: `apps/work-normal-normal/view?ID=${89876 + index}`}))});

describe('background report lookup', () => {
  it('only sends same-origin GETs, retains the report filters and uses the real returned ID', async () => {
    const fetchMock = vi.fn().mockResolvedValueOnce(response(found)).mockResolvedValueOnce(response(detail));
    vi.stubGlobal('fetch', fetchMock);
    const before = location.href;
    const result = await lookupReportTaskFromPage(options());
    expect(result.internalId).toBe('89876'); expect(result.identityError).toBeUndefined();
    expect(location.href).toBe(before);
    expect(fetchMock).toHaveBeenCalledTimes(2);
    for (const [href, init] of fetchMock.mock.calls) {
      expect(new URL(href).origin).toBe(location.origin); expect(init.method).toBe('GET');
      expect(init.redirect).toBe('error'); expect(init.credentials).toBe('include');
    }
  });
  it('parses serialized table data and detail properties without executing event scripts', async () => {
    const props = {object: 'work-task-task', page: {total: 1, params: {s: task.code}},
      rows: [{ID: 89876, vals: {title: '<a href="/apps/work-normal-normal/view?ID=89876">Công việc A</a>'}}]};
    const serialized = `<elem name="table" data-props='${JSON.stringify(props)}'></elem>`;
    const fields = `<elem name="detail.field" data-props='${JSON.stringify({label: 'Mã công việc', content: task.code})}'></elem>`;
    vi.stubGlobal('fetch', vi.fn().mockResolvedValueOnce(response(JSON.stringify({content: serialized})))
      .mockResolvedValueOnce(response(JSON.stringify({content: fields}))));
    expect((await lookupReportTaskFromPage(options())).internalId).toBe('89876');
  });
  it('retains a row on unsupported responses without opening a task page', async () => {
    const fetchMock = vi.fn().mockResolvedValue(response('<html><body>Unknown response</body></html>'));
    vi.stubGlobal('fetch', fetchMock);
    const before = location.href; const result = await lookupReportTaskFromPage(options());
    expect(result.identityError).toContain('Không mở trang công việc');
    expect(result.internalId).toBe(''); expect(location.href).toBe(before);
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });
  it('reads table and detail elements inside nested template children from server markup', async () => {
    const props = {object: 'work-task-task', page: {total: 1, params: {s: task.code}},
      rows: [{ID: 89876, vals: {title: '<a href="/apps/work-normal-normal/view?ID=89876">Công việc A</a>'}}]};
    const attr = (value: unknown) => JSON.stringify(value).replace(/&/g, '&amp;').replace(/"/g, '&quot;');
    // Structure observed in the live list response: app.layout > template > splitview > template > table.
    const list = `<elem name='app.layout'><template children><elem name='splitview'><template children><elem name='table' data-props="${attr(props)}"></elem></template></elem></template></elem>`;
    const nestedDetail = `<template children><template children>${detail}</template></template>`;
    const fetchMock = vi.fn().mockResolvedValueOnce(response(list)).mockResolvedValueOnce(response(nestedDetail));
    vi.stubGlobal('fetch', fetchMock);
    const result = await lookupReportTaskFromPage(options());
    expect(result.internalId).toBe('89876');
    expect(result.identityError).toBeUndefined();
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });
  it('detects a login form even inside template content', async () => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(response('<template><input type="password"></template>')));
    expect((await lookupReportTaskFromPage(options())).identityError).toContain('Phiên đăng nhập đã hết');
  });
  it('uses fields.title and data-children, not the plain vals.title or parent link', async () => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValueOnce(response(liveShapeSearch())).mockResolvedValueOnce(response(detail)));
    const result = await lookupReportTaskFromPage(options());
    expect(result.internalId).toBe('89876'); expect(result.identityError).toBeUndefined();
  });
  it('reads a long title rendered as template children rather than data-children', async () => {
    const fetchMock = vi.fn().mockResolvedValueOnce(response(searchRows([task.name]))).mockResolvedValueOnce(response(detail));
    vi.stubGlobal('fetch', fetchMock);
    expect((await lookupReportTaskFromPage(options())).internalId).toBe('89876');
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });
  it('selects one exact normalized name among a complete set of prefix matches, then verifies its code', async () => {
    const fetchMock = vi.fn().mockResolvedValueOnce(response(searchRows([' Công việc  A ', 'Công việc A - Con', 'Công việc cha'])))
      .mockResolvedValueOnce(response(detail));
    vi.stubGlobal('fetch', fetchMock);
    const result = await lookupReportTaskFromPage(options());
    expect(result.internalId).toBe('89876'); expect(result.identityError).toBeUndefined();
  });
  it('still rejects a detail-code mismatch after selecting among multiple results', async () => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValueOnce(response(searchRows([task.name, 'Con'])))
      .mockResolvedValueOnce(response(detail.replace('CODE.18526', 'OTHER'))));
    const result = await lookupReportTaskFromPage(options());
    expect(result.internalId).toBe(''); expect(result.identityError).toContain('Mã chi tiết');
  });
  it('does not select the first of two tasks with the same name', async () => {
    const fetchMock = vi.fn().mockResolvedValueOnce(response(searchRows([task.name, task.name])));
    vi.stubGlobal('fetch', fetchMock);
    const result = await lookupReportTaskFromPage(options());
    expect(result.internalId).toBe(''); expect(result.identityError).toContain('2 liên kết khớp tên');
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });
  it('does not assume uniqueness from a partial search page', async () => {
    const fetchMock = vi.fn().mockResolvedValueOnce(response(searchRows([task.name], 51)));
    vi.stubGlobal('fetch', fetchMock);
    expect((await lookupReportTaskFromPage(options())).identityError).toContain('phân trang');
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });
  it('tries only the approved leading-zero representation after zero results and keeps the report code with a warning', async () => {
    const fetchMock = vi.fn().mockResolvedValueOnce(response(searchRows([], 0, '0.19363')))
      .mockResolvedValueOnce(response(searchRows([task.name, 'Con'], 2, '.19363')))
      .mockResolvedValueOnce(response(detail.replace('CODE.18526', '.19363')));
    vi.stubGlobal('fetch', fetchMock);
    const result = await lookupReportTaskFromPage({...options(), task: {...task, code: '0.19363'}});
    expect(result.internalId).toBe('89876'); expect(result.code).toBe('0.19363');
    expect(result.identityWarning).toContain('"0.19363" khác mã chi tiết ".19363"');
    expect(result.identityWarning).toContain('không sửa 1Office');
    expect(fetchMock.mock.calls.slice(0, 2).map(([href]) => new URL(href).searchParams.get('s'))).toEqual(['0.19363', '.19363']);
  });
  it.each(['.21750', '.217500', '0.2175000'])('accepts the approved fractional-code representation %s without changing the report code', async actual => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValueOnce(response(searchRows([], 0, '0.2175')))
      .mockResolvedValueOnce(response(searchRows([task.name], 1, '.2175')))
      .mockResolvedValueOnce(response(detail.replace('CODE.18526', actual))));
    const result = await lookupReportTaskFromPage({...options(), task: {...task, code: '0.2175'}});
    expect(result.internalId).toBe('89876'); expect(result.code).toBe('0.2175');
    expect(result.identityWarning).toContain(`"${actual}"`);
    expect(result.identityWarning).toContain('giữ mã báo cáo trong Excel, không sửa 1Office');
  });
  it.each(['.21751', '.021750', '21750', 'TASK.21750', '.2175.0'])('does not broaden the approved rule to %s', async actual => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValueOnce(response(searchRows([], 0, '0.2175')))
      .mockResolvedValueOnce(response(searchRows([task.name], 1, '.2175')))
      .mockResolvedValueOnce(response(detail.replace('CODE.18526', actual))));
    const result = await lookupReportTaskFromPage({...options(), task: {...task, code: '0.2175'}});
    expect(result.internalId).toBe(''); expect(result.identityError).toContain('Không khớp quy tắc');
  });
  it('uses string normalization rather than floating-point rounding for long fractional codes', async () => {
    const code = '0.123456789012345678901';
    const query = code.slice(1);
    vi.stubGlobal('fetch', vi.fn().mockResolvedValueOnce(response(searchRows([], 0, code)))
      .mockResolvedValueOnce(response(searchRows([task.name], 1, query)))
      .mockResolvedValueOnce(response(detail.replace('CODE.18526', '.1234567890123456789020'))));
    expect((await lookupReportTaskFromPage({...options(), task: {...task, code}})).internalId).toBe('');
  });
  it('can find a shortened representation only after both original searches return zero results', async () => {
    const fetchMock = vi.fn().mockResolvedValueOnce(response(searchRows([], 0, '0.21750')))
      .mockResolvedValueOnce(response(searchRows([], 0, '.21750')))
      .mockResolvedValueOnce(response(searchRows([task.name], 1, '.2175')))
      .mockResolvedValueOnce(response(detail.replace('CODE.18526', '.2175')));
    vi.stubGlobal('fetch', fetchMock);
    const result = await lookupReportTaskFromPage({...options(), task: {...task, code: '0.21750'}});
    expect(result.internalId).toBe('89876'); expect(result.code).toBe('0.21750');
    expect(fetchMock.mock.calls.slice(0, 3).map(([href]) => new URL(href).searchParams.get('s'))).toEqual(['0.21750', '.21750', '.2175']);
  });
  it('does not normalize zeros inside a structured task code', async () => {
    const fetchMock = vi.fn().mockResolvedValueOnce(response(searchRows([], 0, 'VPTGD.0.21750')));
    vi.stubGlobal('fetch', fetchMock);
    expect((await lookupReportTaskFromPage({...options(), task: {...task, code: 'VPTGD.0.21750'}})).internalId).toBe('');
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });
  it('does not retry the alias after an ambiguous original result', async () => {
    const fetchMock = vi.fn().mockResolvedValueOnce(response(searchRows([task.name, task.name], 2, '0.19363')));
    vi.stubGlobal('fetch', fetchMock);
    expect((await lookupReportTaskFromPage({...options(), task: {...task, code: '0.19363'}})).internalId).toBe('');
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });
  it('discovers the identity widget from the actual board shell and verifies its task ID and code', async () => {
    const fetchMock = vi.fn().mockResolvedValueOnce(response(liveShapeSearch())).mockResolvedValueOnce(response(boardShell()))
      .mockResolvedValueOnce(response(JSON.stringify({error: false, widgets: {DEFAULT: [
        {ID: 47106, type: 'system', boardId: 2184, module: 'work'},
        {ID: 47091, type: 'system', boardId: 2184, module: 'work'},
      ]}})))
      .mockResolvedValueOnce(response(JSON.stringify({content: '<div class="board-widget-acccess"></div>'})))
      .mockResolvedValueOnce(response(widgetBody()));
    vi.stubGlobal('fetch', fetchMock);
    const before = location.href;
    const result = await lookupReportTaskFromPage(options());
    expect(result).toMatchObject({internalId: '89876', identityWidget: {tabId: 2184, widgetId: 47091}});
    expect(result.identityError).toBeUndefined();
    expect(location.href).toBe(before);
    expect(fetchMock).toHaveBeenCalledTimes(5);
    for (const [href, init] of fetchMock.mock.calls) {
      expect(new URL(href).origin).toBe(location.origin); expect(init.method).toBe('GET');
    }
    expect(new URL(fetchMock.mock.calls[2][0]).searchParams.get('object_id')).toBe('89876');
  });
  it('reuses only a matching tab hint and still verifies the returned widget identity', async () => {
    const fetchMock = vi.fn().mockResolvedValueOnce(response(liveShapeSearch())).mockResolvedValueOnce(response(widgetBody()));
    vi.stubGlobal('fetch', fetchMock);
    const result = await lookupReportTaskFromPage({...options(), widgetHints: {'work-normal-normal': {tabId: 2184, widgetId: 47091}}});
    expect(result.internalId).toBe('89876'); expect(fetchMock).toHaveBeenCalledTimes(2);
    expect(new URL(fetchMock.mock.calls[0][0]).searchParams.get('_json')).toBe('1');
    expect(new URL(fetchMock.mock.calls[1][0]).pathname).toBe('/features/core-app-features-board-objects-widgets/data');
  });
  it.each([[89877, task.code], [89876, 'OTHER.CODE']])('rejects widget identity or code mismatches (%s, %s)', async (id, code) => {
    const fetchMock = vi.fn().mockResolvedValueOnce(response(liveShapeSearch())).mockResolvedValueOnce(response(widgetBody(id, code)));
    vi.stubGlobal('fetch', fetchMock);
    const result = await lookupReportTaskFromPage({...options(), widgetHints: {'work-normal-normal': {tabId: 2184, widgetId: 47091}}});
    expect(result.internalId).toBe(''); expect(result.identityError).toBeTruthy();
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });
  it('rediscovers a missing cached widget instead of trusting an empty response', async () => {
    const fetchMock = vi.fn().mockResolvedValueOnce(response(liveShapeSearch()))
      .mockResolvedValueOnce(response('<div></div>')).mockResolvedValueOnce(response(boardShell()))
      .mockResolvedValueOnce(response(JSON.stringify({error: false, widgets: {DEFAULT: [
        {ID: 47091, type: 'system', boardId: 2184, module: 'work'},
      ]}}))).mockResolvedValueOnce(response(widgetBody()));
    vi.stubGlobal('fetch', fetchMock);
    const result = await lookupReportTaskFromPage({...options(), widgetHints: {'work-normal-normal': {tabId: 2184, widgetId: 999}}});
    expect(result.internalId).toBe('89876');
    expect(result.identityWidget?.widgetId).toBe(47091);
    expect(fetchMock).toHaveBeenCalledTimes(5);
  });
  it('does not fetch after the report filter changes or for an external identity URL', async () => {
    const fetchMock = vi.fn(); vi.stubGlobal('fetch', fetchMock);
    const stale = options(); window.history.replaceState({}, '', '/apps/report-report/view?ID=74&department_ids[]=523');
    expect((await lookupReportTaskFromPage(stale)).identityError).toContain('thay đổi');
    expect((await lookupReportTaskFromPage({...options(), task: {...task, reportLookupUrl: 'https://example.com/apps/work-task-task/view?ID=1'}})).identityError).toContain('không hợp lệ');
    expect(fetchMock).not.toHaveBeenCalled();
  });
});
