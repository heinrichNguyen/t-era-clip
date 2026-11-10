import { collectOneOfficeDiscussion } from './collectionClient';
import { getOneOfficeTabOrigin, oneOfficeOrigin } from './origins';
import { collectOneOfficeTaskFromCommentApi } from './projectCommentApi';
import { collectOneOfficeTaskFastPath } from './projectFastPathProbe';
import type { OneOfficeDiscussionExport } from './types';
import { taskCollectionKey } from './taskCollectionKey';
import { resolveReportTask } from './reportTaskLookup';
import type {
  OneOfficeProjectBatchResult,
  OneOfficeProjectInventory,
  OneOfficeProjectTaskCollectionResult,
  OneOfficeProjectTaskInventoryItem,
} from './types';

export interface OneOfficeTaskPageReadyOptions {
  allowedOrigin: string;
  taskId: string;
  taskPath?: string;
}

export interface OneOfficeTaskPageState {
  ready: boolean;
  rootCount: number;
  replyCount: number;
  signature: string;
}

export interface OneOfficeProjectBatchProgress {
  completedCount: number;
  currentTask: OneOfficeProjectTaskInventoryItem | null;
  result: OneOfficeProjectTaskCollectionResult | null;
  totalCount: number;
}

type BatchProgressListener = (progress: OneOfficeProjectBatchProgress) => void;
type CancellationCheck = () => boolean;

const FAST_PATH_CONCURRENCY = 3;

const chromeApi = typeof globalThis !== 'undefined' ? (globalThis as any).chrome : undefined;

export function readOneOfficeTaskPageState(
  options: OneOfficeTaskPageReadyOptions,
): OneOfficeTaskPageState {
  const notReady = (): OneOfficeTaskPageState => ({
    ready: false,
    rootCount: 0,
    replyCount: 0,
    signature: '',
  });

  if (window.location.origin !== options.allowedOrigin) {
    return notReady();
  }

  const currentUrl = new URL(window.location.href);
  if (
    currentUrl.pathname !== (options.taskPath ?? '/apps/work-task-task/view') ||
    currentUrl.searchParams.get('ID') !== options.taskId
  ) {
    return notReady();
  }

  const isExplicitlyHidden = (element: HTMLElement): boolean => {
    let current: HTMLElement | null = element;
    while (current) {
      const style = window.getComputedStyle(current);
      if (
        current.hidden ||
        current.classList.contains('hidden') ||
        current.getAttribute('aria-hidden') === 'true' ||
        style.display === 'none' ||
        style.visibility === 'hidden'
      ) {
        return true;
      }
      current = current.parentElement;
    }
    return false;
  };

  const discussionRoot = Array.from(
    document.querySelectorAll<HTMLElement>(
      `.cards.comments[comments-id="${options.taskPath === '/apps/work-normal-normal/view' ? 'work-normal-normal' : 'work-task-task'}@${options.taskId}"]`,
    ),
  ).find((element) => !isExplicitlyHidden(element));

  if (!discussionRoot || document.readyState !== 'complete') {
    return notReady();
  }

  const readCommentIds = (selector: string): string[] =>
    Array.from(discussionRoot.querySelectorAll<HTMLElement>(selector))
      .map(
        (comment) =>
          Array.from(comment.classList).find((className) => /^comment\d+$/.test(className)) ?? '',
      )
      .filter(Boolean);
  const rootIds = Array.from(new Set(readCommentIds(':scope > .cards-body > .comment')));
  const replyIds = Array.from(new Set(readCommentIds('.cards.sub-comments .comment')));

  return {
    ready: true,
    rootCount: rootIds.length,
    replyCount: replyIds.length,
    signature: `${rootIds.join(',')}|${replyIds.join(',')}|${discussionRoot.innerHTML.length}`,
  };
}

const getChromeLastError = (): string | undefined => chromeApi?.runtime?.lastError?.message;

const waitForTabNavigation = async (tabId: number, url: string): Promise<void> => {
  if (!chromeApi?.tabs?.update || !chromeApi?.tabs?.onUpdated?.addListener) {
    throw new Error('Chrome tab navigation is unavailable. Reload the extension and try again.');
  }

  return new Promise((resolve, reject) => {
    const timeoutId = globalThis.setTimeout(() => {
      cleanup();
      reject(new Error('The source tab did not finish loading within 30 seconds.'));
    }, 30000);

    const onUpdated = (updatedTabId: number, changeInfo: { status?: string }): void => {
      if (updatedTabId === tabId && changeInfo.status === 'complete') {
        cleanup();
        resolve();
      }
    };

    const cleanup = (): void => {
      globalThis.clearTimeout(timeoutId);
      chromeApi.tabs.onUpdated.removeListener(onUpdated);
    };

    chromeApi.tabs.onUpdated.addListener(onUpdated);
    chromeApi.tabs.update(tabId, { url }, () => {
      const error = getChromeLastError();
      if (error) {
        cleanup();
        reject(new Error(error));
      }
    });
  });
};

const waitForTaskPageReady = async (
  tabId: number,
  taskId: string,
  allowedOrigin: string,
  taskPath?: string,
): Promise<OneOfficeTaskPageState> => {
  if (!chromeApi?.scripting?.executeScript) {
    throw new Error('Chrome scripting access is unavailable. Reload the extension and try again.');
  }

  const startedAt = Date.now();
  let stableSince = 0;
  let lastSignature = '';
  let latestState: OneOfficeTaskPageState | null = null;

  while (Date.now() - startedAt < 30000) {
    try {
      const injectionResults = await chromeApi.scripting.executeScript({
        target: { tabId },
        func: readOneOfficeTaskPageState,
        args: [{ allowedOrigin, taskId, ...(taskPath ? { taskPath } : {}) }],
      });
      const pageState = injectionResults?.[0]?.result as OneOfficeTaskPageState | undefined;

      if (pageState?.ready) {
        latestState = pageState;
        if (pageState.signature !== lastSignature) {
          lastSignature = pageState.signature;
          stableSince = Date.now();
        } else if (Date.now() - stableSince >= 1000) {
          return pageState;
        }
      }
    } catch {
      // A same-origin navigation may briefly replace the main frame. Retry after it settles.
    }

    await new Promise((resolve) => globalThis.setTimeout(resolve, 250));
  }

  throw new Error(
    latestState
      ? `Task ${taskId} discussion did not stabilize within 30 seconds.`
      : `Task ${taskId} did not finish rendering within 30 seconds.`,
  );
};

const createTaskPageUrl = (taskId: string, origin: string): string =>
  `${origin}/apps/work-task-task/view?ID=${encodeURIComponent(taskId)}`;

const taskPageUrl = (task: OneOfficeProjectTaskInventoryItem): string => new URL(task.url).pathname === '/apps/work-normal-normal/view'
  ? task.url : createTaskPageUrl(task.internalId, oneOfficeOrigin(task.url));

const normalizeError = (error: unknown): string =>
  error instanceof Error ? error.message : 'Unknown task collection error.';

export async function collectOneOfficeProjectTaskDiscussions(
  tabId: number,
  inventory: OneOfficeProjectInventory,
  onProgress: BatchProgressListener,
  isCancelled: CancellationCheck,
): Promise<OneOfficeProjectBatchResult> {
  inventory = { ...inventory, tasks: inventory.tasks.map(task => ({...task})) };
  const allowedOrigin = oneOfficeOrigin(inventory.project.url);
  const report = inventory.source === '1office-report-dom';
  for (const task of inventory.tasks) {
    if (report && task.internalId === '' && task.url === ''
      && task.collectionKey === `report-row:${task.sourceOrder}`) continue;
    if (oneOfficeOrigin(task.url) !== allowedOrigin) throw new Error('Danh sách công việc chứa nhiều môi trường 1Office. Thu thập lại từ đúng dự án.');
    const url = new URL(task.url);
    if (!/^\/apps\/work-(task-task\/quickview|task-task\/view|normal-normal\/view)$/.test(url.pathname)
      || url.searchParams.getAll('ID').length !== 1 || url.searchParams.get('ID') !== task.internalId
      || !/^[1-9]\d*$/.test(task.internalId)) throw new Error('Danh sách có liên kết hoặc ID công việc không hợp lệ.');
  }
  if (new Set(inventory.tasks.map(taskCollectionKey)).size !== inventory.tasks.length) throw new Error('Danh sách có khóa dòng trùng. Hãy đọc lại báo cáo.');
  await getOneOfficeTabOrigin(tabId, allowedOrigin);
  const startedAt = new Date().toISOString();
  const resultsByTaskId = new Map<string, OneOfficeProjectTaskCollectionResult>();
  const reportIds = new Set<string>();
  const fallbackTasks: OneOfficeProjectTaskInventoryItem[] = [];
  const warnings: string[] = [];
  let completedCount = 0;
  let sourceTabNavigated = false;

  const normalizeDiscussion = (
    task: OneOfficeProjectTaskInventoryItem,
    discussion: OneOfficeDiscussionExport,
  ): OneOfficeDiscussionExport => ({
    ...discussion,
    entity: {
      ...discussion.entity,
      code: task.code,
      name: task.name,
      url: taskPageUrl(task),
    },
  });

  const emitCollecting = (task: OneOfficeProjectTaskInventoryItem): void => {
    onProgress({
      completedCount,
      currentTask: task,
      result: null,
      totalCount: inventory.tasks.length,
    });
  };

  const saveResult = (taskResult: OneOfficeProjectTaskCollectionResult): void => {
    resultsByTaskId.set(taskCollectionKey(taskResult.task), taskResult);
    completedCount += 1;
    onProgress({
      completedCount,
      currentTask: taskResult.task,
      result: taskResult,
      totalCount: inventory.tasks.length,
    });
  };

  const cancelUnresolvedTasks = (): void => {
    for (const task of inventory.tasks) {
      if (!resultsByTaskId.has(taskCollectionKey(task))) {
        saveResult({
          task,
          state: 'cancelled',
          discussion: null,
          error: null,
        });
      }
    }
  };

  try {
    for (const task of inventory.tasks.filter(task => task.identityError)) {
      saveResult({ task, state: 'unresolved', discussion: null, error: task.identityError! });
    }
    const eligibleTasks = inventory.tasks.filter(task => !task.identityError);
    if (report) {
      const taskIndexes = new Map(inventory.tasks.map((task, index) => [taskCollectionKey(task), index]));
      const updateTask = (task: OneOfficeProjectTaskInventoryItem) => {
        inventory.tasks[taskIndexes.get(taskCollectionKey(task))!] = task;
      };
      const collectReportTask = async (initial: OneOfficeProjectTaskInventoryItem) => {
        let task = initial;
        emitCollecting(task);
        try {
          if (!isCancelled() && !task.internalId) {
            task = await resolveReportTask(tabId, inventory.project.url, task);
            updateTask(task);
            if (task.identityError) {
              saveResult({task, state: 'unresolved', discussion: null, error: task.identityError});
              return;
            }
            if (task.identityWarning) warnings.push(`${task.code ?? task.name}: ${task.identityWarning}`);
          }
          if (isCancelled()) {
            saveResult({task, state: 'cancelled', discussion: null, error: null});
            return;
          }
          if (reportIds.has(task.internalId)) {
            task = {...task, internalId: '', url: '', identityError: 'ID đã xuất hiện ở dòng khác. Giữ dòng công việc, không thu thập thảo luận lặp.'};
            updateTask(task);
            saveResult({task, state: 'unresolved', discussion: null, error: task.identityError!});
            return;
          }
          reportIds.add(task.internalId);
          const direct = await collectOneOfficeTaskFromCommentApi(tabId, task);
          const discussion = direct ?? (new URL(task.url).pathname === '/apps/work-normal-normal/view'
            ? null : await collectOneOfficeTaskFastPath(tabId, task));
          if (!discussion) {
            saveResult({task, state: isCancelled() ? 'cancelled' : 'failed', discussion: null,
              error: isCancelled() ? null : 'Request ngầm chưa trả đủ thảo luận. Giữ dòng công việc, không mở trang chi tiết.'});
            return;
          }
          const normalized = normalizeDiscussion(task, discussion);
          saveResult({task, state: normalized.collection.complete ? 'complete' : 'incomplete', discussion: normalized, error: null});
        } catch (error) {
          saveResult({task, state: 'failed', discussion: null, error: normalizeError(error)});
        }
      };
      let next = 0;
      await Promise.all(Array.from({length: Math.min(FAST_PATH_CONCURRENCY, eligibleTasks.length)}, async () => {
        while (!isCancelled() && next < eligibleTasks.length) await collectReportTask(eligibleTasks[next++]);
      }));
      if (isCancelled()) cancelUnresolvedTasks();
    }
    if (!report) for (let index = 0; index < eligibleTasks.length; index += FAST_PATH_CONCURRENCY) {
      if (isCancelled()) {
        cancelUnresolvedTasks();
        break;
      }

      const batch = eligibleTasks.slice(index, index + FAST_PATH_CONCURRENCY);
      batch.forEach(emitCollecting);
      const fastPathResults = await Promise.all(
        batch.map(async (task) => {
          try {
            if (isCancelled()) return null;
            const directDiscussion = await collectOneOfficeTaskFromCommentApi(tabId, task);
            return directDiscussion ?? (new URL(task.url).pathname === '/apps/work-normal-normal/view'
              ? null : await collectOneOfficeTaskFastPath(tabId, task));
          } catch (error) {
            if (error instanceof Error && error.message.startsWith('Yêu cầu 1Office:')) {
              saveResult({ task, state: 'failed', discussion: null, error: error.message });
            }
            return null;
          }
        }),
      );

      batch.forEach((task, batchIndex) => {
        if (resultsByTaskId.has(taskCollectionKey(task))) return;
        const discussion = fastPathResults[batchIndex];
        if (!discussion) {
          fallbackTasks.push(task);
          return;
        }

        const normalizedDiscussion = normalizeDiscussion(task, discussion);
        saveResult({
          task,
          state: normalizedDiscussion.collection.complete ? 'complete' : 'incomplete',
          discussion: normalizedDiscussion,
          error: null,
        });
      });
    }

    if (!isCancelled()) {
      for (const task of fallbackTasks) {
        if (isCancelled()) {
          cancelUnresolvedTasks();
          break;
        }

        emitCollecting(task);
        let taskResult: OneOfficeProjectTaskCollectionResult;

        try {
          sourceTabNavigated = true;
          await getOneOfficeTabOrigin(tabId, allowedOrigin);
          await waitForTabNavigation(tabId, taskPageUrl(task));
          const pageState = await waitForTaskPageReady(tabId, task.internalId, allowedOrigin,
            new URL(task.url).pathname === '/apps/work-normal-normal/view' ? '/apps/work-normal-normal/view' : undefined);
          let discussion = await collectOneOfficeDiscussion(tabId, undefined, allowedOrigin);

          if (
            discussion.collection.loadedRootCommentCount !== pageState.rootCount ||
            discussion.collection.loadedReplyCount !== pageState.replyCount
          ) {
            await new Promise((resolve) => globalThis.setTimeout(resolve, 500));
            discussion = await collectOneOfficeDiscussion(tabId, undefined, allowedOrigin);
          }

          if (
            discussion.collection.loadedRootCommentCount !== pageState.rootCount ||
            discussion.collection.loadedReplyCount !== pageState.replyCount
          ) {
            throw new Error(
              `Task ${task.internalId} changed while being collected. Expected ${pageState.rootCount} root comment(s) and ${pageState.replyCount} reply/replies, but parsed ${discussion.collection.loadedRootCommentCount} and ${discussion.collection.loadedReplyCount}.`,
            );
          }

          if (discussion.entity.internalId !== task.internalId) {
            throw new Error(
              `Expected task ${task.internalId} but collected task ${discussion.entity.internalId}.`,
            );
          }

          const normalizedDiscussion = normalizeDiscussion(task, discussion);
          taskResult = {
            task,
            state: normalizedDiscussion.collection.complete ? 'complete' : 'incomplete',
            discussion: normalizedDiscussion,
            error: null,
          };
        } catch (error) {
          taskResult = {
            task,
            state: 'failed',
            discussion: null,
            error: normalizeError(error),
          };
        }

        saveResult(taskResult);
      }
    }
  } finally {
    if (sourceTabNavigated) {
      try {
        await getOneOfficeTabOrigin(tabId, allowedOrigin);
        await waitForTabNavigation(tabId, inventory.project.url);
      } catch (error) {
        warnings.push(`Could not restore the source project tab. ${normalizeError(error)}`);
      }
    }
  }

  const results = inventory.tasks.map(
    (task) =>
      resultsByTaskId.get(taskCollectionKey(task)) ?? {
        task,
        state: 'cancelled' as const,
        discussion: null,
        error: null,
      },
  );

  return {
    schemaVersion: 1,
    source: '1office-project-batch-dom',
    startedAt,
    completedAt: new Date().toISOString(),
    cancelled: results.some((result) => result.state === 'cancelled'),
    project: inventory.project,
    inventory,
    tasks: results,
    warnings,
  };
}
