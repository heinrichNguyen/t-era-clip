import { afterEach, describe, expect, it, vi } from 'vitest';
import { strFromU8, unzipSync } from 'fflate';
import manifest from '../../public/manifest.json';
import fixture from '../fixtures/oneoffice/project-discussion.html?raw';
import { isSupportedOneOfficeUrl } from '../../src/oneoffice/collectionClient';
import { readAccessHistoryPage } from '../../src/oneoffice/accessHistory';
import { readAccessPersonnelSource } from '../../src/oneoffice/accessPersonnel';
import { fetchOneOfficeCommentsFromPage } from '../../src/oneoffice/projectCommentApi';
import { probeOneOfficeQuickViewFromPage } from '../../src/oneoffice/projectFastPathProbe';
import { buildAccessWorkbook } from '../../src/oneoffice/export/accessXlsx';
import { createOneOfficeXlsxBlob } from '../../src/oneoffice/export/xlsx';
import { serializeOneOfficeJson } from '../../src/oneoffice/export/json';
import { parseOneOfficeDiscussion } from '../../src/oneoffice/parser';
import { personnelText } from '../../src/oneoffice/personnelText';
import { ONEOFFICE_PROGRESS_MESSAGE } from '../../src/oneoffice/collector';

afterEach(() => { vi.restoreAllMocks(); vi.unstubAllGlobals(); vi.resetModules(); });

describe('security controls (mocked source, no production requests)', () => {
  it('keeps manifest permissions narrow and has no external messaging entry point', () => {
    expect(manifest.permissions).toEqual(['storage', 'contextMenus', 'activeTab', 'scripting']);
    expect(manifest).not.toHaveProperty('host_permissions');
    expect(manifest).not.toHaveProperty('externally_connectable');
    expect(manifest).not.toHaveProperty('content_scripts');
  });
  it.each(['http://office.meygroup.vn', 'https://office.meygroup.vn.attacker.test',
    'https://office.meygroup.vn@attacker.test', 'javascript:alert(1)', 'https://other.test'])
  ('rejects unsupported source %s', (url) => expect(isSupportedOneOfficeUrl(url)).toBe(false));
  it('refuses history and personnel reads off the approved origin before fetching', async () => {
    vi.stubGlobal('location', { origin: 'https://attacker.test', pathname: '/apps/admin-reconciliation-histories' });
    const fetch = vi.fn(); vi.stubGlobal('fetch', fetch);
    const history = await readAccessHistoryPage(1, true);
    const personnel = await readAccessPersonnelSource({ kind: 'directory', tab: 'active', page: 1 });
    expect(history.ok).toBe(false); expect(personnel.ok).toBe(false);
    expect(fetch).not.toHaveBeenCalled();
  });
  it.each([401, 403, 429])('does not turn HTTP %s into usable history or personnel', async (status) => {
    vi.stubGlobal('location', { origin: 'https://office.meygroup.vn', pathname: '/apps/admin-reconciliation-histories' });
    const fetch = vi.fn(async () => new Response('not authorized', { status }));
    vi.stubGlobal('fetch', fetch);
    expect((await readAccessHistoryPage(1, true)).ok).toBe(false);
    expect((await readAccessPersonnelSource({ kind: 'directory', tab: 'active', page: 1 })).ok).toBe(false);
    for (const [, options] of fetch.mock.calls as unknown as Array<[string, RequestInit]>) {
      expect(options.redirect).toBe('error'); expect(options.signal).toBeInstanceOf(AbortSignal);
      expect(options.method ?? 'GET').toBe('GET');
    }
  });
  it('rejects invalid pagination and lookup IDs without making a request', async () => {
    vi.stubGlobal('location', { origin: 'https://office.meygroup.vn', pathname: '/apps/admin-reconciliation-histories' });
    const fetch = vi.fn(); vi.stubGlobal('fetch', fetch);
    expect((await readAccessHistoryPage(-1, true)).ok).toBe(false);
    expect((await readAccessPersonnelSource({ kind: 'lookup', ids: ['1&ids[]=2'] })).ok).toBe(false);
    expect(fetch).not.toHaveBeenCalled();
  });
  it('exports formula-like discussion strings as strings, without formulas or external relationships', async () => {
    const doc = new DOMParser().parseFromString(fixture, 'text/html');
    const result = parseOneOfficeDiscussion(doc, { capturedAt: '2026-10-09T00:00:00Z', sourceUrl: 'https://office.meygroup.vn/apps/work-project-project/view?ID=500' });
    const payloads = ['=HYPERLINK("https://attacker.test","click")', '+SUM(1,2)', '-1+2', '@SUM(1,2)', '\t=1+1'];
    result.comments = payloads.map((contentText, i) => ({ ...result.comments[0], id: String(i), contentText }));
    const blob = await createOneOfficeXlsxBlob(result);
    const bytes = await new Promise<ArrayBuffer>((resolve, reject) => {
      const reader = new FileReader(); reader.onerror = () => reject(reader.error);
      reader.onload = () => resolve(reader.result as ArrayBuffer); reader.readAsArrayBuffer(blob);
    });
    const files = unzipSync(new Uint8Array(bytes));
    for (const [path, value] of Object.entries(files)) {
      if (path.startsWith('xl/worksheets/')) expect(strFromU8(value)).not.toMatch(/<f(?:\s|>)/);
      if (path.endsWith('.rels')) expect(strFromU8(value)).not.toContain('TargetMode="External"');
    }
    const xml = strFromU8(files['xl/sharedStrings.xml']);
    expect(xml).toContain('HYPERLINK'); expect(xml).toContain('@SUM(1,2)');
    expect(JSON.parse(serializeOneOfficeJson(result)).comments.map((row: { contentText: string }) => row.contentText)).toEqual(payloads);
  });
  it('strips personnel markup and keeps formula-like account names as text cells', () => {
    expect(personnelText('&lt;span&gt;TDR&lt;/span&gt;<script>bad()</script><iframe src="https://attacker.test"></iframe>')).toBe('TDR');
    const workbook = buildAccessWorkbook({ cutoff: '2026-10-09', collectedAt: '', recordCount: 1, oldestTime: '',
      users: [{ id: '1', account: '=1+1', name: '@SUM(1,2)', timeIso: '', timeRaw: '', eventCount: 1 }] });
    for (const index of [1, 2]) expect(workbook[0].data![1][index]).toMatchObject({ type: String, format: '@' });
  });
});

describe('request and progress hardening', () => {
  it('blocks redirects and sets deadlines on comment and quickview fetches', async () => {
    const fetch = vi.fn(async () => new Response('{}', { headers: { 'content-type': 'application/json' } }));
    vi.stubGlobal('fetch', fetch);
    await fetchOneOfficeCommentsFromPage({ allowedOrigin: window.location.origin, taskId: '1', offset: 0 });
    await probeOneOfficeQuickViewFromPage({ allowedOrigin: window.location.origin, taskId: '1' });
    expect(fetch).toHaveBeenCalledTimes(2);
    for (const [, options] of fetch.mock.calls as unknown as Array<[string, RequestInit]>) {
      expect(options.method).toBe('GET');
      expect(options.redirect).toBe('error'); expect(options.signal).toBeInstanceOf(AbortSignal);
    }
  });
  it('checks progress sender, source tab, collection ID and payload before notifying the UI', async () => {
    let handler: (message: unknown, sender: unknown) => void = () => {};
    vi.stubGlobal('chrome', { runtime: { id: 'extension', onMessage: {
      addListener: vi.fn((callback) => { handler = callback; }), removeListener: vi.fn(),
    } } });
    const { subscribeToOneOfficeProgress: subscribe } = await import('../../src/oneoffice/collectionClient');
    const listener = vi.fn(); const unsubscribe = subscribe(listener, { tabId: 42, collectionId: 'current' });
    handler({ type: ONEOFFICE_PROGRESS_MESSAGE, progress: { bogus: 'unvalidated' } }, { id: 'untrusted', tab: { id: 999 } });
    const progress = { loadedRootCommentCount: 3, expectedRootCommentCount: 10, loadMoreClicks: 1 };
    const message = { type: ONEOFFICE_PROGRESS_MESSAGE, collectionId: 'current', progress };
    const sender = { id: 'extension', tab: { id: 42 }, frameId: 0, url: 'https://office.meygroup.vn/apps/work-task-task/view?ID=1' };
    handler(message, { ...sender, id: 'other' });
    handler(message, { ...sender, tab: { id: 999 } });
    handler(message, { ...sender, frameId: 1 });
    handler(message, { ...sender, url: 'https://attacker.test' });
    handler({ ...message, collectionId: 'previous' }, sender);
    handler({ ...message, progress: { ...progress, loadedRootCommentCount: -1 } }, sender);
    handler(null, sender);
    expect(listener).not.toHaveBeenCalled();
    handler(message, sender);
    expect(listener).toHaveBeenCalledExactlyOnceWith(progress);
    unsubscribe();
  });
});

// These still assert observed limitations, not security guarantees.
describe('remaining security limitations', () => {
  it('documents shared review storage without account binding, expiry or consume-on-read', async () => {
    const doc = new DOMParser().parseFromString(fixture, 'text/html');
    const result = parseOneOfficeDiscussion(doc, { capturedAt: '2020-01-01T00:00:00Z', sourceUrl: 'https://office.meygroup.vn/apps/work-project-project/view?ID=500' });
    const stored: Record<string, unknown> = {};
    const remove = vi.fn();
    vi.stubGlobal('chrome', { runtime: {}, storage: { session: {
      set: (value: Record<string, unknown>, done: () => void) => { Object.assign(stored, value); done(); },
      get: (_keys: string[], done: (value: Record<string, unknown>) => void) => done(stored), remove,
    } } });
    const { saveOneOfficeReview, getOneOfficeReview } = await import('../../src/oneoffice/reviewSession');
    await saveOneOfficeReview(result);
    expect(await getOneOfficeReview()).toEqual(result);
    expect(await getOneOfficeReview()).toEqual(result);
    expect(Object.keys(stored)).toEqual(['oneOfficeDiscussionReview']);
    expect(remove).not.toHaveBeenCalled();
    const otherReview = { ...result, entity: { ...result.entity, internalId: '999', name: 'Other review' } };
    await saveOneOfficeReview(otherReview);
    expect((await getOneOfficeReview())?.entity.internalId).toBe('999');
  });
  it('documents unknown fields surviving review validation and JSON serialization', async () => {
    const { isOneOfficeDiscussionExport } = await import('../../src/oneoffice/validation');
    const result = parseOneOfficeDiscussion(new DOMParser().parseFromString(fixture, 'text/html'), {
      capturedAt: '2026-10-09T00:00:00Z', sourceUrl: 'https://office.meygroup.vn/apps/work-project-project/view?ID=500',
    });
    const extra = { ...result, unexpectedPrivateField: 'synthetic-marker' };
    expect(isOneOfficeDiscussionExport(extra)).toBe(true);
    expect(JSON.parse(serializeOneOfficeJson(extra)).unexpectedPrivateField).toBe('synthetic-marker');
  });
});
