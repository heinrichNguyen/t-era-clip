import { afterEach, expect, it, vi } from 'vitest';
import { collectTestTaskDetails, parseTestTaskDetails, parseTestTaskWidgets, readTestTaskWidgetsFromPage } from '../../src/oneoffice/testTaskDetails';

const html = (id = '7698', project = '203') => `<a href="/apps/work-project-project/view?ID=${project}">Test project</a>
  <div class="cards comments" comments-id="work-task-task@${id}"><div class="cards-body"></div></div>
  <div class="detail-field"><div class="detail-field-label">Mã công việc</div><div class="detail-field-content">.18902.18908</div></div>
  <div class="detail-field"><div class="detail-field-label">No.</div><div class="detail-field-content">--</div></div>`;
afterEach(() => { vi.unstubAllGlobals(); vi.useRealTimers(); });
it('extracts the observed code but retains placeholder No. as display evidence, not a raw blank', () => {
  const result = parseTestTaskDetails(html(), '7698');
  expect(result.identityVerified).toBe(true);
  expect(result.code).toEqual({ state: 'value', displayValue: '.18902.18908' });
  expect(result.no).toEqual({ state: 'placeholder', displayValue: '--' });
  expect(result.error).toContain('Chưa xác minh');
  expect(result).not.toHaveProperty('sourceNo');
});
it.each([html('999'), html('7698', '870'), '<div id="app"></div>', '<form><input type="password"></form>',
  html() + '<div class="cards comments" comments-id="work-task-task@999"></div>'])('rejects wrong identity, login and metadata-only shells', (body) => {
  const result = parseTestTaskDetails(body, '7698');
  expect(result.identityVerified).toBe(false);
  expect(result.code.state).toBe('missing');
  expect(result.error).toBeTruthy();
});
it('does not choose one of conflicting field labels or execute embedded markup', () => {
  const result = parseTestTaskDetails(html() + `<script>throw new Error('bad')</script><div class="detail-field"><div class="detail-field-label">No.</div><div class="detail-field-content">1.7</div></div>`, '7698');
  expect(result.no.state).toBe('ambiguous');
});
it('reads a displayed No. without overwriting or declaring the API source No. verified', () => {
  const result = parseTestTaskDetails(html().replace('>--<', '>1.7<'), '7698');
  expect(result.no).toEqual({ state: 'value', displayValue: '1.7' });
  expect(result.error).toBeNull();
  expect(result).not.toHaveProperty('sourceNo');
});
it.each(['https://office.meygroup.vn/apps/work-project-project/view?ID=203',
  'https://testoffice.meygroup.vn:9443/apps/work-project-project/view?ID=999'])('refuses an unexpected source before fetching: %s', async (href) => {
  vi.stubGlobal('window', { location: { href } });
  const fetch = vi.fn(); vi.stubGlobal('fetch', fetch);
  expect((await readTestTaskWidgetsFromPage('7698')).error).toBeTruthy();
  expect(fetch).not.toHaveBeenCalled();
});
it('uses only the two observed GET widgets, without navigation, redirects, or save parameters', async () => {
  vi.stubGlobal('window', { location: { href: 'https://testoffice.meygroup.vn:9443/apps/work-project-project/view?ID=203' } });
  const fetch = vi.fn(async () => new Response(html(), { status: 200 })); vi.stubGlobal('fetch', fetch);
  expect((await readTestTaskWidgetsFromPage('7698')).general).toContain('.18902.18908');
  expect(fetch).toHaveBeenCalledTimes(2);
  expect(fetch.mock.calls.map(([url]) => new URL(url as string).searchParams.get('ID'))).toEqual(['23', '1915']);
  for (const [url, options] of fetch.mock.calls as unknown as [string, RequestInit][]) {
    const parsed = new URL(url);
    expect(parsed.origin).toBe('https://testoffice.meygroup.vn:9443');
    expect(parsed.pathname).toBe('/features/core-app-features-board-objects-widgets/data');
    expect(parsed.searchParams.get('object_id')).toBe('7698');
    expect(parsed.searchParams.get('mode')).toBe('view');
    expect(options).toMatchObject({ method: 'GET', redirect: 'error', signal: expect.any(AbortSignal) });
  }
});

const escape = (text: string) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/'/g, '&#039;').replace(/"/g, '&quot;');
const elem = (name: string, props: Record<string, unknown>) => `<elem name='${name}' data-props='${escape(JSON.stringify(props))}'></elem>`;
const field = (label: string, content?: string) => elem('detail.field', content === undefined ? { label } : { label, content });
const widget = (id: number, body: string, postId = '7698', name = id === 23 ? 'Work-Task-TaskNormalInformation' : 'customField') => JSON.stringify({
  content: elem('visual.widget', { url: '/apps/work-task-task/view/widget', postId: Number(postId), board_id: id, name,
    body: escape(`<elem name='div'><template children><elem name='detail.section'><template children>${body}</template></elem></template></elem>`) }), w: 1, h: 1,
});
const project = (id = '203', origin = '') => `<elem name='a' data-atts='${escape(JSON.stringify({ href: `${origin}/apps/work-project-project/view?ID=${id}` }))}' data-children='["Project"]'></elem>`;
const general = (projectContent = project(), extra = '') => widget(23, field('Dự án', projectContent) + field('Mã công việc', '.18902.18908') + extra);
const custom = (no: string | undefined = '') => widget(1915, field('No.', no));
// Network uses a raw JSON attribute with the body encoded once, not an escaped JSON attribute.
const networkWidget = (id: number, body: string, postId = '7698') => JSON.stringify({ content:
  `<elem name='visual.widget' data-props='${JSON.stringify({ url: '/apps/work-task-task/view/widget', postId: Number(postId),
    name: id === 23 ? 'Work-Task-TaskNormalInformation' : 'customField', board_id: id,
    body: escape(`<elem name='div'><template children><elem name='detail.section'><template children>${body}</template></elem></template></elem>`),
  })}'></elem>`, w: 1, h: 1 });
const networkGeneral = (projectContent = project(), postId = '7698') => networkWidget(23,
  field('Tên công việc', 'Quyết định đầu tư dự án') + field('Dự án', projectContent) + field('Mã công việc', '.18902.18908'), postId);
it('reads the actual single-encoded Network envelope without decoding nested attributes twice', () => {
  const result = parseTestTaskWidgets(networkGeneral(), networkWidget(1915, field('No.', '')), '7698');
  expect(result.identityVerified).toBe(true);
  expect(result.code).toEqual({ state: 'value', displayValue: '.18902.18908' });
  expect(result.no).toEqual({ state: 'empty', displayValue: '' });
  expect(result.error).toBeNull();
});
it.each([
  [networkGeneral(project('870')), networkWidget(1915, field('No.', '')), 'xác minh liên kết dự án'],
  [networkGeneral(project('203', 'https://example.com')), networkWidget(1915, field('No.', '')), 'xác minh liên kết dự án'],
  [networkGeneral(project() + project('870')), networkWidget(1915, field('No.', '')), 'xác minh liên kết dự án'],
  [networkGeneral(project(), '988'), networkWidget(1915, field('No.', '')), 'xác minh định danh khối 23'],
  [networkGeneral(), networkWidget(1915, field('No.', ''), '999'), 'xác minh định danh khối 1915'],
])('preserves identity checks for the single-encoded envelope and reports the failing stage', (g, c, stage) => {
  const result = parseTestTaskWidgets(g, c, '7698');
  expect(result.identityVerified).toBe(false);
  expect(result.code.state).toBe('missing');
  expect(result.failureStage).toBe(stage);
  expect(result.error).toContain(stage);
  expect(JSON.stringify(result)).not.toContain('data-props');
});
it('keeps missing, empty, display-only and ambiguous fields distinct with the real envelope', () => {
  expect(parseTestTaskWidgets(networkGeneral(), networkWidget(1915, field('No.')), '7698').no.state).toBe('missing');
  expect(parseTestTaskWidgets(networkGeneral(), networkWidget(1915, field('No.', '--')), '7698').no.state).toBe('placeholder');
  expect(parseTestTaskWidgets(networkGeneral(), networkWidget(1915, field('No.', '1.2')), '7698').no).toEqual({ state: 'value', displayValue: '1.2' });
  expect(parseTestTaskWidgets(networkGeneral(), networkWidget(1915, field('No.', '') + field('No.', '1.2')), '7698').no.state).toBe('ambiguous');
});
it('handles nested quotes and entities without interpreting them as new field attributes', () => {
  const body = field('Dự án', project()) + field('Mã công việc', '.18902.18908') +
    field('Công việc cha', `<elem name='a' data-atts='${escape(JSON.stringify({ href: '/apps/work-normal-normal/view?ID=7459' }))}' data-children='["Parent &amp; child"]'></elem>`);
  expect(parseTestTaskWidgets(networkWidget(23, body), networkWidget(1915, field('No.', '')), '7698').identityVerified).toBe(true);
});
it.each([0, 35.5, 100])('reads numeric progress %s from the observed progress component, not from task status', (percent) => {
  const g = networkWidget(23, field('Dự án', project()) + field('Mã công việc', '.18902.18908') + field('Tiến độ', elem('progress', { type: 'progresstask', percent, iconClick: 'throw new Error("not executable")' })));
  const result = parseTestTaskWidgets(g, networkWidget(1915, field('No.', '')), '7698');
  expect(result.progress).toEqual({ state: 'value', percent });
  expect(result).not.toHaveProperty('iconClick');
});
it.each([-1, 101, '35', null])('does not infer unsupported progress values: %s', (percent) => {
  const g = networkWidget(23, field('Dự án', project()) + field('Tiến độ', elem('progress', { type: 'progresstask', percent })));
  expect(parseTestTaskWidgets(g, networkWidget(1915, field('No.', '')), '7698').progress).toEqual({ state: 'missing', percent: null });
});
it('reads nested JSON, encoded bodies and inert template children from the observed widgets', () => {
  const result = parseTestTaskWidgets(general(), custom(), '7698');
  expect(result.identityVerified).toBe(true);
  expect(result.code).toEqual({ state: 'value', displayValue: '.18902.18908' });
  expect(result.no).toEqual({ state: 'empty', displayValue: '' });
  expect(result.error).toBeNull();
  expect(result.evidenceSource).toBe('widget-response');
  expect(result).not.toHaveProperty('sourceNo');
});
it.each([
  [general(project('870')), custom()],
  [general(project('203', 'https://example.com')), custom()],
  [general(project('203&ID=870')), custom()],
  [general(project() + project('870')), custom()],
  [widget(23, field('Dự án', project()), '988'), custom()],
  [general(), widget(1915, field('No.', ''), '999')],
  [general(), widget(1915, field('No.', ''), '7698', 'otherWidget')],
  [general(), '{"content":"<form><input type=password></form>"}'],
  ['<div id="app"></div>', custom()],
  [general(), '{"content":"<elem name=visual.widget data-props=invalid></elem>"}'],
])('fails closed on wrong identity, foreign links, login and changed widget schema', (g, c) => {
  const result = parseTestTaskWidgets(g, c, '7698');
  expect(result.identityVerified).toBe(false);
  expect(result.code.state).toBe('missing');
  expect(result.error).toBeTruthy();
});
it('keeps missing, placeholder, duplicate and explicit empty No. distinct', () => {
  expect(parseTestTaskWidgets(general(), widget(1915, field('No.')), '7698').no.state).toBe('missing');
  expect(parseTestTaskWidgets(general(), custom('--'), '7698').no.state).toBe('placeholder');
  expect(parseTestTaskWidgets(general(), widget(1915, field('No.', '') + field('No.', '1.2')), '7698').no.state).toBe('ambiguous');
  expect(parseTestTaskWidgets(general(), custom('1.2'), '7698').no).toEqual({ state: 'value', displayValue: '1.2' });
});
it('never executes site callbacks or treats nested controls as a code value', () => {
  const result = parseTestTaskWidgets(general(project(), '<script>throw new Error("execute")</script>'), custom(), '7698');
  expect(result.identityVerified).toBe(true);
  expect(parseTestTaskWidgets(general(), custom('<elem name="input" value="1.2"></elem>'), '7698').no.state).toBe('missing');
});
it.each([302, 401, 500])('stops on HTTP %s without retry or reading the second widget', async (status) => {
  vi.stubGlobal('window', { location: { href: 'https://testoffice.meygroup.vn:9443/apps/work-project-project/view?ID=203' } });
  const fetch = vi.fn(async () => new Response('', { status })); vi.stubGlobal('fetch', fetch);
  expect((await readTestTaskWidgetsFromPage('7698')).error).toBeTruthy();
  expect(fetch).toHaveBeenCalledTimes(1);
});
it('rejects oversized content without starting another request', async () => {
  vi.stubGlobal('window', { location: { href: 'https://testoffice.meygroup.vn:9443/apps/work-project-project/view?ID=203' } });
  const fetch = vi.fn(async () => new Response('x'.repeat(5_000_001))); vi.stubGlobal('fetch', fetch);
  expect((await readTestTaskWidgetsFromPage('7698')).error).toBeTruthy();
  expect(fetch).toHaveBeenCalledTimes(1);
});
it('stops if the source tab changes between requests', async () => {
  const location = { href: 'https://testoffice.meygroup.vn:9443/apps/work-project-project/view?ID=203' };
  vi.stubGlobal('window', { location });
  const fetch = vi.fn(async () => { location.href += '&changed=1'; return new Response(general()); }); vi.stubGlobal('fetch', fetch);
  expect((await readTestTaskWidgetsFromPage('7698')).error).toBeTruthy();
  expect(fetch).toHaveBeenCalledTimes(1);
});
it('collects only parsed fields and does not retain raw widget content or callbacks', async () => {
  vi.stubGlobal('chrome', { scripting: { executeScript: vi.fn(async () => [{ result: { general: general(), custom: custom() } }]) } });
  const result = await collectTestTaskDetails(123, '7698');
  expect(result.identityVerified).toBe(true);
  expect(result).not.toHaveProperty('general');
  expect(result).not.toHaveProperty('custom');
});
it('aborts after the shared deadline and never retries', async () => {
  vi.useFakeTimers();
  vi.stubGlobal('window', { location: { href: 'https://testoffice.meygroup.vn:9443/apps/work-project-project/view?ID=203' } });
  const fetch = vi.fn((_url: string, init: RequestInit) => new Promise<Response>((_resolve, reject) => {
    init.signal!.addEventListener('abort', () => reject(new Error('Abort')), { once: true });
  })); vi.stubGlobal('fetch', fetch);
  const pending = readTestTaskWidgetsFromPage('7698');
  await vi.advanceTimersByTimeAsync(30000);
  expect((await pending).error).toContain('30 giây');
  expect(fetch).toHaveBeenCalledTimes(1);
  expect(vi.getTimerCount()).toBe(0);
});
it.each(['../7698', '7698&ID=999', 'not-an-id'])('rejects malformed task IDs before fetching', async (id) => {
  vi.stubGlobal('window', { location: { href: 'https://testoffice.meygroup.vn:9443/apps/work-project-project/view?ID=203' } });
  const fetch = vi.fn(); vi.stubGlobal('fetch', fetch);
  expect((await readTestTaskWidgetsFromPage(id)).error).toBeTruthy();
  expect(fetch).not.toHaveBeenCalled();
});
