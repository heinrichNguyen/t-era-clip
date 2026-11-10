import type { ActiveTab } from './collectionClient';
import type { OneOfficeProjectInventory } from './types';
import { oneOfficeOrigin } from './origins';

const PROJECT_BATCH_SESSION_KEY = 'oneOfficeProjectBatchSession';
const chromeApi = typeof globalThis !== 'undefined' ? (globalThis as any).chrome : undefined;

export interface OneOfficeProjectBatchSession {
  sourceTab: ActiveTab;
  inventory: OneOfficeProjectInventory;
}

const getSessionStorage = () => {
  const storage = chromeApi?.storage?.session;
  if (!storage) {
    throw new Error('Chrome session storage is unavailable. Reload the extension and try again.');
  }
  return storage;
};

const getChromeLastError = (): string | undefined => chromeApi?.runtime?.lastError?.message;

export async function saveOneOfficeProjectBatchSession(
  session: OneOfficeProjectBatchSession,
): Promise<void> {
  if (oneOfficeOrigin(session.sourceTab.url) !== oneOfficeOrigin(session.inventory.project.url)) {
    throw new Error('Tab nguồn và danh sách dự án khác môi trường 1Office. Hãy thu thập lại.');
  }
  const storage = getSessionStorage();

  return new Promise((resolve, reject) => {
    storage.set({ [PROJECT_BATCH_SESSION_KEY]: session }, () => {
      const error = getChromeLastError();
      if (error) {
        reject(new Error(error));
        return;
      }
      resolve();
    });
  });
}

export async function getOneOfficeProjectBatchSession(): Promise<OneOfficeProjectBatchSession | null> {
  const storage = getSessionStorage();

  return new Promise((resolve, reject) => {
    storage.get([PROJECT_BATCH_SESSION_KEY], (stored: Record<string, unknown>) => {
      const error = getChromeLastError();
      if (error) {
        reject(new Error(error));
        return;
      }

      const value = stored[PROJECT_BATCH_SESSION_KEY] as OneOfficeProjectBatchSession | undefined;
      if (!value) {
        resolve(null);
        return;
      }

      if (
        typeof value.sourceTab?.id !== 'number' ||
        !Array.isArray(value.inventory?.tasks) ||
        (value.inventory.project?.type !== 'project' && !(value.inventory.project?.type === 'report'
          && value.inventory.source === '1office-report-dom' && value.inventory.collection?.complete))
      ) {
        reject(new Error('The temporary project collection session is invalid. Start it again.'));
        return;
      }

      resolve(value);
    });
  });
}

export async function openOneOfficeProjectBatchReview(
  sourceTab: ActiveTab,
  inventory: OneOfficeProjectInventory,
): Promise<void> {
  if (!chromeApi?.tabs?.create || !chromeApi?.runtime?.getURL) {
    throw new Error('Chrome tab creation is unavailable.');
  }

  await saveOneOfficeProjectBatchSession({ sourceTab, inventory });
  await chromeApi.tabs.create({ url: chromeApi.runtime.getURL('project-review.html') });
}
