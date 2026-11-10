import { mkdirSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createRequire } from 'node:module';

// Developer-only preview. Mock data is never included in the extension package.
const [modulesPath, releasePath] = process.argv.slice(2);
if (!modulesPath || !releasePath) throw new Error('Usage: node scripts/release-preview.mjs <bundled-node-modules> <release-directory>');
const require = createRequire(resolve(modulesPath, 'package.json'));
const { chromium } = require('playwright');
const output = resolve(releasePath, 'screenshots');
mkdirSync(output, { recursive: true });
const browser = await chromium.launch({ channel: 'chrome', headless: true });
try {
  const context = await browser.newContext({ viewport: { width: 1280, height: 800 }, deviceScaleFactor: 1 });
  await context.route('**/*', (route) => {
    if (new URL(route.request().url()).hostname === '127.0.0.1') return route.continue();
    return route.abort();
  });
  await context.addInitScript(() => {
    const project = { type: 'project', internalId: '999999', name: 'Dự án minh họa (dữ liệu giả)', code: 'DEMO', url: 'https://office.meygroup.vn/apps/work-project-project/view?ID=999999' };
    const tasks = [
      ['1', 'Kế hoạch minh họa'], ['1.1', 'Chuẩn bị hồ sơ minh họa'], ['1.1.1', 'Kiểm tra tài liệu minh họa'],
    ].map(([no, name], index) => ({
      internalId: String(990001 + index), no, code: `DEMO.${no}`, name, sourceOrder: index,
      url: `https://office.meygroup.vn/apps/work-task-task/view?ID=${990001 + index}`,
      depth: no.split('.').length, parentNo: index ? no.slice(0, no.lastIndexOf('.')) : null,
      parentInternalId: index ? String(990000 + index) : null,
    }));
    const inventory = { schemaVersion: 1, source: '1office-project-api', capturedAt: new Date().toISOString(), project, tasks,
      collection: { complete: true, expectedTaskCount: tasks.length, loadedTaskCount: tasks.length, warnings: [] } };
    const storage = {
      teraClipItems: [], teraClipSettings: { defaultReportLanguage: 'vi' },
      oneOfficeProjectBatchSession: { sourceTab: { id: 1, title: project.name, url: project.url }, inventory },
    };
    globalThis.chrome = {
      runtime: { getURL: (path) => `${location.origin}/${path}`, onMessage: { addListener() {}, removeListener() {} } },
      tabs: { query: async () => [{ id: 1, title: project.name, url: project.url }], create: async () => ({ id: 2 }) },
      storage: { local: { get: (keys, callback) => callback(storage) }, session: { get: (keys, callback) => callback(storage) } },
      scripting: { executeScript: async ({ args }) => {
        const id = args?.[0]?.taskId;
        if (!id) return [{ result: inventory }];
        return [{ result: {
          total: id === '990002' ? 1 : 0,
          items: id === '990002' ? [{
            ID: 991001, parent_id: 0, dateTime: '10:00:00 02/10/2026',
            content: 'Nội dung minh họa: đề nghị kiểm tra hồ sơ.', user: { ID: 9901, name: 'Người dùng mẫu A' }, files: [], total_subs: 1,
            subs: [{ ID: 991002, parent_id: 991001, dateTime: '10:05:00 02/10/2026', content: 'Phản hồi minh họa: đã tiếp nhận.', user: { ID: 9902, name: 'Người dùng mẫu B' }, files: [] }],
          }] : [],
        } }];
      } },
    };
  });
  const page = await context.newPage();
  const errors = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.goto('http://127.0.0.1:4173/popup.html');
  await page.getByRole('button', { name: 'Kiểm tra danh sách công việc dự án' }).click();
  await page.getByText('3 công việc, tối đa 3 cấp').waitFor();
  await page.addStyleTag({ content: 'body { display: flex; justify-content: center; align-items: center; min-height: 100vh; }' });
  await page.screenshot({ path: resolve(output, '01-menu-du-lieu-gia.png') });
  await page.goto('http://127.0.0.1:4173/project-review.html');
  await page.getByText('Đã hoàn tất thu thập.', { exact: true }).waitFor();
  if (!await page.getByText('1 bình luận, 1 phản hồi', { exact: true }).count()) throw new Error('Unexpected mock collection result.');
  if (await page.getByText('Phân tích AI theo yêu cầu').count()) throw new Error('Removed AI UI is still present.');
  await page.screenshot({ path: resolve(output, '02-ket-qua-du-lieu-gia.png') });
  await page.getByRole('button', { name: 'Xuất Excel', exact: true }).click();
  await page.getByRole('dialog').waitFor();
  const downloadEvent = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Tải file Excel', exact: true }).click();
  const download = await downloadEvent;
  const exportPath = resolve(releasePath, 'TEST-DU-LIEU-GIA.xlsx');
  await download.saveAs(exportPath);
  await page.getByRole('dialog').waitFor({ state: 'hidden' });
  if (errors.length) throw new Error(errors.join('\n'));
  writeFileSync(resolve(releasePath, 'PREVIEW-CHECK.json'), JSON.stringify({
    method: 'Fresh headless Chrome, localhost production preview, mock Chrome APIs and synthetic 1Office data; external requests blocked.',
    actualExtensionInstallTested: false, liveOneOfficeTested: false, windowsTested: false,
    screenshots: ['01-menu-du-lieu-gia.png', '02-ket-qua-du-lieu-gia.png'], exportedMockWorkbook: 'TEST-DU-LIEU-GIA.xlsx', pageErrors: errors,
  }, null, 2));
  console.log('Rendered 2 synthetic-data screenshots and downloaded XLSX through the current UI. No external data requests.');
} finally {
  await browser.close();
}
