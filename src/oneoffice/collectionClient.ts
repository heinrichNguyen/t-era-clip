import {
  collectOneOfficeDiscussionFromPage,
  ONEOFFICE_PROGRESS_MESSAGE,
  type OneOfficeCollectionProgress,
  type OneOfficePageSnapshot,
} from './collector';
import { parseOneOfficeDiscussion } from './parser';
import { collectOneOfficeProjectInventoryFromPage } from './projectInventory';
import { collectOneOfficeProjectInventoryFromApi } from './projectInventoryApi';
import {
  buildOneOfficeFilteredTaskInventory,
  collectOneOfficeFilteredTaskSelectionFromPage,
} from './filteredTaskInventory';
import type {
  OneOfficeDiscussionExport,
  OneOfficeProjectInventory,
} from './types';

import { getOneOfficeTabOrigin, oneOfficeOrigin } from './origins';
export { ONEOFFICE_ALLOWED_ORIGIN } from './origins';

export interface ActiveTab {
  id: number;
  title: string;
  url: string;
}

type ProgressListener = (progress: OneOfficeCollectionProgress) => void;

const chromeApi = typeof globalThis !== 'undefined' ? (globalThis as any).chrome : undefined;

export function isSupportedOneOfficeUrl(url: string): boolean {
  try {
    oneOfficeOrigin(url);
    return true;
  } catch {
    return false;
  }
}

export function isOneOfficeProjectPageUrl(url: string): boolean {
  try {
    const parsedUrl = new URL(url);
    return (
      isSupportedOneOfficeUrl(url) &&
      parsedUrl.pathname === '/apps/work-project-project/view' &&
      /^\d+$/.test(parsedUrl.searchParams.get('ID') ?? '')
    );
  } catch {
    return false;
  }
}

export function isOneOfficeFilteredTaskListUrl(url: string): boolean {
  try {
    const parsedUrl = new URL(url);
    return (
      isSupportedOneOfficeUrl(url) &&
      parsedUrl.pathname === '/apps/work-task-task/list' &&
      (parsedUrl.searchParams.has('@sfilter') || parsedUrl.searchParams.has('project_id'))
    );
  } catch {
    return false;
  }
}

export function isOneOfficeReportPageUrl(url: string): boolean {
  try {
    const parsed = new URL(url);
    return isSupportedOneOfficeUrl(url) && parsed.pathname === '/apps/report-report/view'
      && parsed.searchParams.getAll('ID').length === 1 && /^[1-9]\d*$/.test(parsed.searchParams.get('ID') ?? '');
  } catch { return false; }
}

export async function getActiveTab(): Promise<ActiveTab | null> {
  if (!chromeApi?.tabs?.query) {
    throw new Error('Chrome tab access is unavailable.');
  }

  const tabs = await chromeApi.tabs.query({ active: true, currentWindow: true });
  const tab = tabs?.[0];

  if (typeof tab?.id !== 'number') {
    return null;
  }

  return {
    id: tab.id,
    title: typeof tab.title === 'string' ? tab.title : '',
    url: typeof tab.url === 'string' ? tab.url : '',
  };
}

export async function collectOneOfficeDiscussion(
  tabId: number,
  onProgress?: ProgressListener,
  expectedOrigin?: string,
): Promise<OneOfficeDiscussionExport> {
  if (!chromeApi?.scripting?.executeScript) {
    throw new Error('Chrome scripting access is unavailable. Reload the extension and try again.');
  }

  const allowedOrigin = await getOneOfficeTabOrigin(tabId, expectedOrigin);
  const collectionId = crypto.randomUUID();
  const unsubscribe = onProgress
    ? subscribeToOneOfficeProgress(onProgress, { tabId, collectionId, origin: allowedOrigin })
    : () => undefined;
  try {
    const injectionResults = await chromeApi.scripting.executeScript({
      target: { tabId },
      func: collectOneOfficeDiscussionFromPage,
      args: [
        {
          allowedOrigin,
          maxLoadMoreClicks: 100,
          noProgressTimeoutMs: 5000,
          progressMessageType: ONEOFFICE_PROGRESS_MESSAGE,
          settleTimeMs: 750,
          collectionId,
        },
      ],
    });
    const snapshot = injectionResults?.[0]?.result as OneOfficePageSnapshot | undefined;

    if (!snapshot?.discussionHtml) {
      throw new Error('The 1Office collector returned no discussion data.');
    }

    return parseOneOfficeSnapshot(snapshot);
  } finally {
    unsubscribe();
  }
}

export async function collectOneOfficeProjectInventory(
  tabId: number,
  expectedOrigin?: string,
): Promise<OneOfficeProjectInventory> {
  if (!chromeApi?.scripting?.executeScript) {
    throw new Error('Chrome scripting access is unavailable. Reload the extension and try again.');
  }

  const allowedOrigin = await getOneOfficeTabOrigin(tabId, expectedOrigin);
  let apiError: unknown;
  try {
    const apiResults = await chromeApi.scripting.executeScript({
      target: { tabId },
      func: collectOneOfficeProjectInventoryFromApi,
      args: [{ allowedOrigin }],
    });
    const apiInventory = apiResults?.[0]?.result as OneOfficeProjectInventory | undefined;
    if (apiInventory?.tasks) {
      return apiInventory;
    }
  } catch (error) {
    if (error instanceof Error && error.message.startsWith('Yêu cầu 1Office:')) throw error;
    apiError = error;
  }

  const domResults = await chromeApi.scripting.executeScript({
    target: { tabId },
    func: collectOneOfficeProjectInventoryFromPage,
    args: [{ allowedOrigin }],
  });
  const inventory = domResults?.[0]?.result as OneOfficeProjectInventory | undefined;

  if (!inventory?.tasks) {
    throw new Error('The 1Office project inventory collector returned no task data.');
  }

  if (!inventory.collection.complete) {
    const reason = apiError instanceof Error ? ` ${apiError.message}` : '';
    throw new Error(
      `1Office only exposed a partial project task table and the full Gantt inventory could not be loaded.${reason}`,
    );
  }

  return inventory;
}

export async function collectOneOfficeFilteredTaskInventory(
  tabId: number,
  expectedOrigin?: string,
): Promise<OneOfficeProjectInventory> {
  if (!chromeApi?.scripting?.executeScript) {
    throw new Error('Chrome scripting access is unavailable. Reload the extension and try again.');
  }

  const allowedOrigin = await getOneOfficeTabOrigin(tabId, expectedOrigin);
  const selectionResults = await chromeApi.scripting.executeScript({
    target: { tabId },
    func: collectOneOfficeFilteredTaskSelectionFromPage,
    args: [{ allowedOrigin }],
  });
  const selection = selectionResults?.[0]?.result;
  if (!selection?.tasks) {
    throw new Error('The 1Office filtered task collector returned no task data.');
  }

  const projectUrl = `${allowedOrigin}/apps/work-project-project/view?ID=${encodeURIComponent(selection.projectInternalId)}`;
  const projectResults = await chromeApi.scripting.executeScript({
    target: { tabId },
    func: collectOneOfficeProjectInventoryFromApi,
    args: [{ allowedOrigin, sourceUrl: projectUrl }],
  });
  const projectInventory = projectResults?.[0]?.result as OneOfficeProjectInventory | undefined;
  if (!projectInventory?.tasks) {
    throw new Error('The 1Office project inventory collector returned no task data.');
  }

  return buildOneOfficeFilteredTaskInventory(selection, projectInventory);
}

export function parseOneOfficeSnapshot(
  snapshot: OneOfficePageSnapshot,
): OneOfficeDiscussionExport {
  const parsedDocument = new DOMParser().parseFromString(
    `<main>${snapshot.detailFieldsHtml}${snapshot.discussionHtml}</main>`,
    'text/html',
  );
  const parsed = parseOneOfficeDiscussion(parsedDocument, {
    capturedAt: snapshot.capturedAt,
    loadMoreClicks: snapshot.loadMoreClicks,
    sourceUrl: snapshot.sourceUrl,
  });
  const warnings = Array.from(
    new Set([...snapshot.warnings, ...parsed.collection.warnings]),
  );

  return {
    ...parsed,
    collection: {
      ...parsed.collection,
      complete: snapshot.complete && parsed.collection.complete,
      expectedRootCommentCount:
        snapshot.expectedRootCommentCount ?? parsed.collection.expectedRootCommentCount,
      warnings,
    },
  };
}

export function subscribeToOneOfficeProgress(
  listener: ProgressListener,
  expected: { tabId: number; collectionId: string; origin?: string },
): () => void {
  if (!chromeApi?.runtime?.onMessage?.addListener) {
    return () => undefined;
  }

  const messageListener = (message: unknown, sender: any) => {
    if (!message || typeof message !== 'object' || Array.isArray(message) ||
        !chromeApi.runtime.id || sender?.id !== chromeApi.runtime.id ||
        sender?.tab?.id !== expected.tabId || sender?.frameId !== 0 ||
        !isSupportedOneOfficeUrl(sender?.url ?? '') ||
        (expected.origin !== undefined && new URL(sender.url).origin !== expected.origin)) return;
    const candidate = message as {
      type?: string;
      collectionId?: string;
      progress?: OneOfficeCollectionProgress;
    };
    const progress = candidate.progress;
    const count = (value: unknown) => Number.isSafeInteger(value) && Number(value) >= 0;
    if (candidate.type === ONEOFFICE_PROGRESS_MESSAGE && candidate.collectionId === expected.collectionId &&
        progress && typeof progress === 'object' && !Array.isArray(progress) &&
        count(progress.loadedRootCommentCount) && count(progress.loadMoreClicks) &&
        (progress.expectedRootCommentCount === null || count(progress.expectedRootCommentCount))) {
      listener({ loadedRootCommentCount: progress.loadedRootCommentCount,
        expectedRootCommentCount: progress.expectedRootCommentCount, loadMoreClicks: progress.loadMoreClicks });
    }
  };

  chromeApi.runtime.onMessage.addListener(messageListener);
  return () => chromeApi.runtime.onMessage.removeListener(messageListener);
}
