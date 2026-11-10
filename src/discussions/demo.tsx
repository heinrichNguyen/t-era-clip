// Development-only screenshot harness. No real 1Office calls, credentials or user data.
if (!import.meta.env.DEV) throw new Error('Demo is unavailable in release builds.');
const origin = 'https://testoffice.meygroup.vn:9443';
const screen = new URLSearchParams(location.search).get('screen') || 'popup';
const partial = new URLSearchParams(location.search).get('partial') === '1';
const capturedAt = '2026-11-03T02:00:00.000Z';
const tasks = [
  { internalId: '101', name: 'Rà soát hồ sơ trước cuộc họp', code: 'DEMO.01', no: '1', parentInternalId: null, parentNo: null, depth: 1 },
  { internalId: '102', name: 'Xác nhận phương án phối hợp', code: 'DEMO.02', no: '1.1', parentInternalId: '101', parentNo: '1', depth: 2 },
  { internalId: '103', name: 'Tổng hợp ý kiến phòng ban', code: 'DEMO.03', no: '1.2', parentInternalId: '101', parentNo: '1', depth: 2 },
].map((task, i) => ({ ...task, sourceOrder: i, url: `${origin}/apps/work-task-task/view?ID=${task.internalId}` }));
const apiItems = (id: string) => id === '103' ? [] : [{ ID: Number(id) * 10, content: id === '101' ? 'Đề nghị rà soát đủ hồ sơ trước mốc họp ngày 03/11.' : 'Đã tổng hợp các nội dung cần phối hợp giữa phòng ban.',
  dateTime: '09:00:00 02/11/2026', user: { ID: 1, name: 'Key user A' }, files: [],
  total_subs: id === '101' ? 1 : 0, subs: id === '101' ? [{ ID: 1011, parent_id: 1010, content: 'Đã nhận yêu cầu. Sẽ gửi bản rà soát trong ngày.', dateTime: '10:30:00 02/11/2026', user: { ID: 2, name: 'Key user B' }, files: [] }] : [] }];
const comments = [{ id: '1010', parentId: null, rootId: '1010', depth: 0, sourceOrder: 0, authorId: '1', authorName: 'Key user A', createdAtRaw: '09:00:00 02/11/2026', createdAtIso: '2026-11-02T02:00:00.000Z', contentText: 'Đề nghị rà soát đủ hồ sơ trước mốc họp ngày 03/11.', mentions: [], attachments: [] },
  { id: '1011', parentId: '1010', rootId: '1010', depth: 1, sourceOrder: 1, authorId: '2', authorName: 'Key user B', createdAtRaw: '10:30:00 02/11/2026', createdAtIso: '2026-11-02T03:30:00.000Z', contentText: 'Đã nhận yêu cầu. Sẽ gửi bản rà soát trong ngày.', mentions: [], attachments: [] }];
const discussion = { schemaVersion: 1, source: '1office-dom', capturedAt, entity: { type: 'task', ...tasks[0] },
  collection: { complete: true, loadedRootCommentCount: 1, loadedReplyCount: 1, expectedRootCommentCount: 1, loadMoreClicks: 0, warnings: [] }, comments };
const inventory = { schemaVersion: 1, source: '1office-project-api', capturedAt,
  project: { type: 'project', internalId: '100', name: 'DỰ ÁN MINH HỌA KEY USER', code: 'DEMO', url: `${origin}/apps/work-project-project/view?ID=100` },
  collection: { complete: true, expectedTaskCount: tasks.length, loadedTaskCount: tasks.length, warnings: [] }, tasks };
const active = { id: 1, title: inventory.project.name, url: inventory.project.url };
const stored: Record<string, unknown> = { oneOfficeDiscussionReview: discussion, oneOfficeProjectBatchSession: { sourceTab: active, inventory } };
const listeners = new Set<(changes: Record<string, unknown>, area: string) => void>();
(globalThis as any).chrome = {
  runtime: { getURL: (path: string) => `/discussion-demo.html?screen=${path === 'review.html' ? 'task-review' : 'project-review'}`,
    onMessage: { addListener: () => {}, removeListener: () => {} } },
  tabs: { query: async () => [active], get: async () => active, create: async ({ url }: { url: string }) => { location.href = url; } },
  storage: { session: {
    get: (keys: string[], done: (value: unknown) => void) => done(Object.fromEntries(keys.map((key) => [key, stored[key]]))),
    set: (values: Record<string, unknown>, done: () => void) => { Object.assign(stored, values); done(); listeners.forEach((fn) => fn(values, 'session')); },
    remove: (keys: string[], done: () => void) => { keys.forEach((key) => delete stored[key]); done(); },
  }, onChanged: { addListener: (fn: any) => listeners.add(fn), removeListener: (fn: any) => listeners.delete(fn) } },
  scripting: { executeScript: async ({ func, args }: { func: Function; args: any[] }) => {
    if (func.name === 'fetchOneOfficeCommentsFromPage') {
      if (partial && args[0].taskId === '103') return [{ result: { requestError: 'Yêu cầu 1Office: HTTP 403 (mô phỏng hướng dẫn).' } }];
      const items = apiItems(args[0].taskId); return [{ result: { items, total: items.length } }];
    }
    if (func.name === 'collectOneOfficeProjectInventoryFromApi') return [{ result: inventory }];
    if (func.name === 'collectOneOfficeDiscussionFromPage') return [{ result: { capturedAt, complete: true, sourceUrl: tasks[0].url, expectedRootCommentCount: 1, loadMoreClicks: 0, warnings: [],
      detailFieldsHtml: `<div class="detail-field"><div class="detail-field-label">Tên công việc</div><div class="detail-field-content">${tasks[0].name}</div></div>`,
      discussionHtml: '<div class="cards comments" comments-id="work-task-task@101"><div class="cards-body"><div class="comment comment1010"><div class="comment-head"><a class="userlink" href="/user?ID=1">Key user A</a></div><div class="comment-post">Đề nghị rà soát đủ hồ sơ trước mốc họp.</div><div class="comment-panel-time"><span title="09:00:00 02/11/2026">Hôm qua</span></div></div></div></div>',
    } }];
    throw new Error(`Unexpected demo request: ${func.name}`);
  } },
};
const notice = document.createElement('div');
notice.textContent = 'ẢNH GIAO DIỆN THẬT VỚI DỮ LIỆU MINH HỌA • Không kết nối 1Office';
notice.style.cssText = 'padding:12px 20px;background:#fef3c7;color:#78350f;font:600 14px Arial'; document.body.prepend(notice);
if (screen === 'task-review') await import('../review/main');
else if (screen === 'project-review') await import('../project-review/main');
else await import('./main');
