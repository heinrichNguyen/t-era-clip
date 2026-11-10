import type { OneOfficeProjectInventory } from './types';
import { openNoReview, readNoReview } from './noReviewSession';
import { ONEOFFICE_ALLOWED_ORIGIN, oneOfficeOrigin } from './origins';

export async function openProjectImportWorkspace(inventory?: OneOfficeProjectInventory | null, sourceTabId?: number) {
  if (inventory && Number.isInteger(sourceTabId) && sourceTabId! > 0) {
    return openNoReview(inventory, sourceTabId!, 'import-review.html');
  }
  const api = (globalThis as any).chrome;
  if (!api?.runtime?.getURL || !api?.tabs?.create) throw new Error('Không thể mở trang chuẩn hóa mã. Hãy tải lại extension và thử lại.');
  await api.tabs.create({ url: api.runtime.getURL('import-review.html') });
}

export async function readProjectImportContext(search: string) {
  const query = new URLSearchParams(search);
  if (!query.has('projectId') && !query.has('capturedAt')) return null;
  const inventory = await readNoReview();
  if (!inventory || !query.get('projectId') || !query.get('capturedAt') || inventory.project.internalId !== query.get('projectId') || inventory.capturedAt !== query.get('capturedAt') ||
      oneOfficeOrigin(inventory.project.url) !== oneOfficeOrigin(query.get('sourceOrigin') ?? ONEOFFICE_ALLOWED_ORIGIN)) {
    throw new Error('Danh sách dự án tạm đã thay đổi hoặc bị xóa. Mở lại từ dự án cần đối chiếu. Bạn vẫn có thể kiểm tra file trên máy.');
  }
  return inventory;
}
