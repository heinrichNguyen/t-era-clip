import type { OneOfficeProjectInventory } from './types';
import { oneOfficeOrigin } from './origins';

export const NO_REVIEW_KEY = 'oneOfficeNoReview';

export async function openNoReview(inventory: OneOfficeProjectInventory, sourceTabId: number, destination: 'no-review.html' | 'import-review.html' = 'no-review.html') {
  const api = (globalThis as any).chrome;
  if (!api?.storage?.session || !api?.tabs?.create || !api?.runtime?.getURL) {
    throw new Error('Không thể mở đối chiếu. Hãy tải lại extension và thử lại.');
  }
  await new Promise<void>((resolve, reject) => {
    api.storage.session.set({ [NO_REVIEW_KEY]: inventory }, () => {
      if (api.runtime.lastError) reject(new Error(api.runtime.lastError.message));
      else resolve();
    });
  });
  const url = new URL(api.runtime.getURL(destination));
  url.searchParams.set('sourceTabId', String(sourceTabId));
  url.searchParams.set('sourceOrigin', oneOfficeOrigin(inventory.project.url));
  if (destination === 'import-review.html') {
    url.searchParams.set('projectId', inventory.project.internalId);
    url.searchParams.set('capturedAt', inventory.capturedAt);
  }
  await api.tabs.create({ url: url.href });
}

export async function readNoReview(): Promise<OneOfficeProjectInventory | null> {
  const api = (globalThis as any).chrome;
  if (!api?.storage?.session) throw new Error('Chỉ mở trang này từ T-eraClip.');
  return new Promise((resolve, reject) => {
    api.storage.session.get([NO_REVIEW_KEY], (stored: Record<string, unknown>) => {
      if (api.runtime.lastError) { reject(new Error(api.runtime.lastError.message)); return; }
      const value = stored[NO_REVIEW_KEY] as OneOfficeProjectInventory | undefined;
      if (value && (value.project?.type !== 'project' || !Array.isArray(value.tasks))) {
        reject(new Error('Danh sách tạm không hợp lệ. Hãy kiểm tra lại danh sách dự án.'));
        return;
      }
      resolve(value ?? null);
    });
  });
}
