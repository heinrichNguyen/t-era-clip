export const ONEOFFICE_PROGRESS_MESSAGE = 'oneoffice:collection-progress';

export interface OneOfficeCollectionProgress {
  loadedRootCommentCount: number;
  expectedRootCommentCount: number | null;
  loadMoreClicks: number;
}

export interface OneOfficePageSnapshot {
  capturedAt: string;
  complete: boolean;
  detailFieldsHtml: string;
  discussionHtml: string;
  expectedRootCommentCount: number | null;
  loadMoreClicks: number;
  sourceUrl: string;
  warnings: string[];
}

export interface OneOfficeCollectorOptions {
  allowedOrigin: string;
  maxLoadMoreClicks: number;
  noProgressTimeoutMs: number;
  progressMessageType: string;
  settleTimeMs: number;
  collectionId?: string;
}

export async function collectOneOfficeDiscussionFromPage(
  options: OneOfficeCollectorOptions,
): Promise<OneOfficePageSnapshot> {
  const discussionSelector = '.cards.comments[comments-id]';
  const rootCommentSelector = ':scope > .cards-body > .comment';
  const loadMoreSelector = ':scope > .prev .bt, :scope > .cards-body > .prev .bt';
  const supportedIdPattern = /^(work-project-project|work-task-task)@(\d+)$/;
  const warnings: string[] = [];

  const normalizeLabel = (value: string): string =>
    value
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/đ/gi, (character) => (character === 'Đ' ? 'D' : 'd'))
      .toLowerCase()
      .replace(/\s+/g, ' ')
      .trim();

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

  const findVisibleLoadMore = (root: HTMLElement): HTMLElement | null => {
    const candidate = root.querySelector<HTMLElement>(loadMoreSelector);
    return candidate && !isExplicitlyHidden(candidate) ? candidate : null;
  };

  const countRootComments = (root: HTMLElement): number => {
    const ids = new Set<string>();
    let commentsWithoutId = 0;

    for (const comment of root.querySelectorAll<HTMLElement>(rootCommentSelector)) {
      const idClass = Array.from(comment.classList).find((className) => /^comment\d+$/.test(className));
      if (idClass) {
        ids.add(idClass);
      } else {
        commentsWithoutId += 1;
      }
    }

    return ids.size + commentsWithoutId;
  };

  const parseExpectedRootCount = (label: string, loadedCount: number): number | null => {
    const match = /xem\s+them\s+\d+\s*\/\s*(\d+)/i.exec(normalizeLabel(label));
    return match ? loadedCount + Number(match[1]) : null;
  };

  const reportProgress = (progress: OneOfficeCollectionProgress): void => {
    const extensionApi = (
      globalThis as typeof globalThis & {
        chrome?: { runtime?: { sendMessage?: (message: unknown) => unknown } };
      }
    ).chrome;

    try {
      const result = extensionApi?.runtime?.sendMessage?.({
        type: options.progressMessageType,
        collectionId: options.collectionId,
        progress,
      });

      if (result && typeof (result as Promise<unknown>).catch === 'function') {
        void (result as Promise<unknown>).catch(() => undefined);
      }
    } catch {
      // Progress reporting is optional; collection must continue if the popup closes.
    }
  };

  const waitForRootCountChange = async (
    root: HTMLElement,
    previousCount: number,
  ): Promise<boolean> =>
    new Promise((resolve) => {
      const startedAt = Date.now();
      const intervalId = window.setInterval(() => {
        const currentCount = countRootComments(root);
        const buttonGone = !findVisibleLoadMore(root);
        const timedOut = Date.now() - startedAt >= options.noProgressTimeoutMs;

        if (currentCount > previousCount || buttonGone || timedOut) {
          window.clearInterval(intervalId);
          resolve(currentCount > previousCount || buttonGone);
        }
      }, 100);
    });

  const waitForFinalRootCount = async (
    root: HTMLElement,
    expectedCount: number | null,
  ): Promise<number> =>
    new Promise((resolve) => {
      const startedAt = Date.now();
      let lastCount = countRootComments(root);
      let lastCountChangeAt = startedAt;

      const intervalId = window.setInterval(() => {
        const now = Date.now();
        const currentCount = countRootComments(root);

        if (currentCount !== lastCount) {
          lastCount = currentCount;
          lastCountChangeAt = now;
          reportProgress({
            loadedRootCommentCount: currentCount,
            expectedRootCommentCount: expectedCount,
            loadMoreClicks,
          });
        }

        const reachedExpectedCount = expectedCount === null || currentCount >= expectedCount;
        const isStable = now - lastCountChangeAt >= options.settleTimeMs;
        const timedOut = now - startedAt >= options.noProgressTimeoutMs;

        if ((reachedExpectedCount && isStable) || timedOut) {
          window.clearInterval(intervalId);
          resolve(currentCount);
        }
      }, 100);
    });

  if (window.location.origin !== options.allowedOrigin) {
    throw new Error('This tab is not on the approved 1Office origin.');
  }

  const candidates = Array.from(document.querySelectorAll<HTMLElement>(discussionSelector)).filter(
    (candidate) =>
      !isExplicitlyHidden(candidate) &&
      supportedIdPattern.test(candidate.getAttribute('comments-id') ?? ''),
  );

  if (candidates.length === 0) {
    throw new Error('No supported visible 1Office discussion was found on this page.');
  }

  const activeCandidate = candidates.find((candidate) => candidate.closest('.active'));
  const discussionRoot = activeCandidate ?? candidates[candidates.length - 1];
  let loadMoreClicks = 0;
  let expectedRootCommentCount: number | null = null;

  while (loadMoreClicks < options.maxLoadMoreClicks) {
    const loadMoreButton = findVisibleLoadMore(discussionRoot);
    const loadedCount = countRootComments(discussionRoot);

    if (!loadMoreButton) {
      reportProgress({
        loadedRootCommentCount: loadedCount,
        expectedRootCommentCount: expectedRootCommentCount ?? loadedCount,
        loadMoreClicks,
      });
      break;
    }

    if (expectedRootCommentCount === null) {
      expectedRootCommentCount = parseExpectedRootCount(
        loadMoreButton.textContent ?? '',
        loadedCount,
      );
    }
    reportProgress({
      loadedRootCommentCount: loadedCount,
      expectedRootCommentCount,
      loadMoreClicks,
    });

    loadMoreButton.click();
    loadMoreClicks += 1;

    const progressed = await waitForRootCountChange(discussionRoot, loadedCount);
    if (!progressed) {
      warnings.push(
        `Stopped after ${loadMoreClicks} load-more attempt(s) because the root comment count did not change.`,
      );
      break;
    }
  }

  let remainingLoadMore = findVisibleLoadMore(discussionRoot);
  if (remainingLoadMore && loadMoreClicks >= options.maxLoadMoreClicks) {
    warnings.push(`Stopped after reaching the ${options.maxLoadMoreClicks}-click safety limit.`);
  }

  let finalRootCount = countRootComments(discussionRoot);
  if (!remainingLoadMore && loadMoreClicks > 0) {
    finalRootCount = await waitForFinalRootCount(
      discussionRoot,
      expectedRootCommentCount,
    );
    remainingLoadMore = findVisibleLoadMore(discussionRoot);
  }

  const reachedExpectedRootCount =
    expectedRootCommentCount === null || finalRootCount >= expectedRootCommentCount;
  if (!reachedExpectedRootCount) {
    warnings.push(
      `Collection stopped at ${finalRootCount} of ${expectedRootCommentCount} expected root comments.`,
    );
  }

  reportProgress({
    loadedRootCommentCount: finalRootCount,
    expectedRootCommentCount: expectedRootCommentCount ?? finalRootCount,
    loadMoreClicks,
  });

  const discussionClone = discussionRoot.cloneNode(true) as HTMLElement;
  for (const removable of discussionClone.querySelectorAll(
    'form, button, input, textarea, [contenteditable="true"], .comment-menu, .comment-panel-reply, .comment-panel-react',
  )) {
    removable.remove();
  }

  const allowedFieldLabels = new Set([
    'ma cong viec',
    'ma du an',
    'ten cong viec',
    'ten du an',
  ]);
  const detailFieldsHtml = Array.from(document.querySelectorAll<HTMLElement>('.detail-field'))
    .filter((field) => {
      const label = field.querySelector('.detail-field-label')?.textContent ?? '';
      return !isExplicitlyHidden(field) && allowedFieldLabels.has(normalizeLabel(label));
    })
    .map((field) => field.outerHTML)
    .join('');

  return {
    capturedAt: new Date().toISOString(),
    complete: !remainingLoadMore && reachedExpectedRootCount,
    detailFieldsHtml,
    discussionHtml: discussionClone.outerHTML,
    expectedRootCommentCount: expectedRootCommentCount ?? finalRootCount,
    loadMoreClicks,
    sourceUrl: window.location.href,
    warnings,
  };
}
