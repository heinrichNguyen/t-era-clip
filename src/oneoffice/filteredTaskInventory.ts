import type {
  OneOfficeProjectInventory,
  OneOfficeProjectTaskInventoryItem,
} from './types';

export interface OneOfficeFilteredTaskSelectionOptions {
  allowedOrigin: string;
  sourceUrl?: string;
}

export interface OneOfficeFilteredTaskSelection {
  sourceUrl: string;
  projectInternalId: string;
  projectName: string;
  expectedTaskCount: number | null;
  complete: boolean;
  warnings: string[];
  tasks: OneOfficeProjectTaskInventoryItem[];
}

export async function collectOneOfficeFilteredTaskSelectionFromPage(
  options: OneOfficeFilteredTaskSelectionOptions,
): Promise<OneOfficeFilteredTaskSelection> {
  const sourceUrl = options.sourceUrl ?? window.location.href;
  const parsedUrl = new URL(sourceUrl);

  if (parsedUrl.origin !== options.allowedOrigin) {
    throw new Error('This tab is not on the approved 1Office origin.');
  }
  if (parsedUrl.pathname !== '/apps/work-task-task/list') {
    throw new Error('Open a filtered 1Office task list before checking its tasks.');
  }

  const normalizeText = (value: string | null | undefined): string =>
    (value ?? '').replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();

  const projectIds = new Set<string>();
  const collectProjectIds = (value: unknown): void => {
    if (Array.isArray(value)) {
      value.forEach(collectProjectIds);
      return;
    }
    if (!value || typeof value !== 'object') {
      return;
    }

    const record = value as Record<string, unknown>;
    if (record.f === 'project_id' && Array.isArray(record.p)) {
      record.p.forEach((id) => {
        const normalizedId = normalizeText(typeof id === 'string' || typeof id === 'number' ? String(id) : '');
        if (/^\d+$/.test(normalizedId)) {
          projectIds.add(normalizedId);
        }
      });
    }
    Object.values(record).forEach(collectProjectIds);
  };

  const serializedFilter = parsedUrl.searchParams.get('@sfilter');
  if (serializedFilter) {
    try {
      collectProjectIds(JSON.parse(serializedFilter));
    } catch {
      throw new Error('The current 1Office task filter could not be read.');
    }
  }
  const directProjectId = normalizeText(parsedUrl.searchParams.get('project_id'));
  if (/^\d+$/.test(directProjectId)) {
    projectIds.add(directProjectId);
  }
  if (projectIds.size !== 1) {
    throw new Error('The filtered list must belong to exactly one project.');
  }
  const projectInternalId = Array.from(projectIds)[0];

  const readPageInfo = (doc: Document): { expected: number | null; totalPages: number } => {
    const title = normalizeText(doc.querySelector<HTMLElement>('.table-page-title')?.textContent);
    const expectedMatch = /\/\s*([\d,.]+)\s*bản ghi/i.exec(title);
    const expected = expectedMatch
      ? Number(expectedMatch[1].replace(/[,.]/g, ''))
      : null;
    const pagesText = normalizeText(doc.querySelector<HTMLElement>('.table-pages')?.textContent);
    const pagesMatch = /Trang:\s*\d+\s*\/\s*(\d+)/i.exec(pagesText);
    return {
      expected: Number.isFinite(expected) ? expected : null,
      totalPages: pagesMatch ? Math.max(1, Number(pagesMatch[1])) : 1,
    };
  };

  const parseTasks = (doc: Document): OneOfficeProjectTaskInventoryItem[] =>
    Array.from(doc.querySelectorAll<HTMLTableRowElement>('tr.table-row-body')).flatMap((row) => {
      const link = row.querySelector<HTMLAnchorElement>(
        'a.table-main-title[href*="/apps/work-normal-normal/view?ID="], a.table-main-title[href*="/apps/work-task-task/view?ID="], a.table-main-title[href*="/apps/work-task-task/quickview?ID="]',
      );
      if (!link) {
        return [];
      }

      const taskUrl = new URL(link.getAttribute('href') ?? '', parsedUrl.href);
      const internalId = normalizeText(taskUrl.searchParams.get('ID'));
      if (!/^\d+$/.test(internalId)) {
        return [];
      }

      const readColumn = (name: string): string =>
        normalizeText(row.querySelector<HTMLElement>(`[col="${name}"]`)?.textContent);
      const parentLink = row.querySelector<HTMLAnchorElement>(
        'td[col="title"] a[href*="/view?ID="]:not(.table-main-title)',
      );
      const parentId = parentLink
        ? normalizeText(new URL(parentLink.getAttribute('href') ?? '', parsedUrl.href).searchParams.get('ID'))
        : '';

      return [{
        sourceOrder: 0,
        internalId,
        no: '',
        code: readColumn('code') || null,
        name: normalizeText(link.textContent),
        url: `${parsedUrl.origin}/apps/work-task-task/view?ID=${encodeURIComponent(internalId)}`,
        depth: 0,
        parentNo: null,
        parentInternalId: /^\d+$/.test(parentId) ? parentId : null,
        taskStatus: readColumn('status'),
        plannedStart: readColumn('start_plan'),
        plannedEnd: readColumn('end_plan'),
        actualStart: readColumn('start'),
        actualEnd: readColumn('end'),
        overdueDays: readColumn('out_date'),
      }];
    });

  const pageInfo = readPageInfo(document);
  const readVisibleRange = (): { start: number; end: number; total: number } | null => {
    const title = normalizeText(
      document.querySelector<HTMLElement>('.table-page-title')?.textContent,
    );
    const match = /Hiển thị\s*([\d,.]+)\s*-\s*([\d,.]+)\s*\/\s*([\d,.]+)\s*bản ghi/i.exec(title);
    if (!match) {
      return null;
    }

    const [start, end, total] = match.slice(1).map((value) =>
      Number(value.replace(/[,.]/g, '')),
    );
    return [start, end, total].every(Number.isFinite) ? { start, end, total } : null;
  };
  const readCurrentPage = (): number => {
    const pagesText = normalizeText(document.querySelector<HTMLElement>('.table-pages')?.textContent);
    const match = /Trang:\s*(\d+)\s*\/\s*\d+/i.exec(pagesText);
    return match ? Math.max(1, Number(match[1])) : 1;
  };
  const initialPage = readCurrentPage();
  const warnings: string[] = [];
  const tasksById = new Map<string, OneOfficeProjectTaskInventoryItem>();
  let projectName = '';

  const collectCurrentPage = (): void => {
    if (!projectName) {
      projectName = normalizeText(
        document.querySelector<HTMLElement>('tr.table-row-body [col="project_title"]')?.textContent,
      );
    }
    parseTasks(document).forEach((task) => {
      if (!tasksById.has(task.internalId)) {
        tasksById.set(task.internalId, { ...task, sourceOrder: tasksById.size });
      }
    });
  };

  const taskSignature = (): string =>
    parseTasks(document)
      .map((task) => task.internalId)
      .sort((left, right) => left.localeCompare(right))
      .join(',');

  const moveOnePage = async (direction: 'left' | 'right'): Promise<void> => {
    const previousPage = readCurrentPage();
    const targetPage = previousPage + (direction === 'right' ? 1 : -1);
    const previousSignature = taskSignature();
    const button = document.querySelector<HTMLElement>(
      `.table-header-btn.icon-caret-${direction}`,
    );
    if (!button || button.classList.contains('disabled')) {
      throw new Error(`The 1Office ${direction} page control is unavailable.`);
    }

    button.click();
    const startedAt = Date.now();
    let stableSignature = '';
    let stableSince = 0;
    while (Date.now() - startedAt < 15000) {
      if (readCurrentPage() === targetPage) {
        const visibleRange = readVisibleRange();
        const currentTasks = parseTasks(document);
        const currentSignature = currentTasks
          .map((task) => task.internalId)
          .sort((left, right) => left.localeCompare(right))
          .join(',');
        const expectedVisibleCount = visibleRange
          ? Math.max(0, visibleRange.end - visibleRange.start + 1)
          : null;
        const pageIsFullyRendered =
          visibleRange !== null
          && visibleRange.total === pageInfo.expected
          && currentTasks.length === expectedVisibleCount
          && new Set(currentTasks.map((task) => task.internalId)).size === expectedVisibleCount
          && currentSignature !== previousSignature;

        if (pageIsFullyRendered) {
          if (currentSignature !== stableSignature) {
            stableSignature = currentSignature;
            stableSince = Date.now();
          } else if (Date.now() - stableSince >= 600) {
            return;
          }
        } else {
          stableSignature = '';
          stableSince = 0;
        }
      }
      await new Promise((resolve) => globalThis.setTimeout(resolve, 100));
    }
    throw new Error(`Timed out while opening filtered task page ${targetPage}.`);
  };

  const moveToPage = async (targetPage: number): Promise<void> => {
    while (readCurrentPage() > targetPage) {
      await moveOnePage('left');
    }
    while (readCurrentPage() < targetPage) {
      await moveOnePage('right');
    }
  };

  try {
    await moveToPage(1);
    for (let page = 1; page <= pageInfo.totalPages; page += 1) {
      collectCurrentPage();
      if (page < pageInfo.totalPages) {
        await moveOnePage('right');
      }
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Unknown pagination error.';
    warnings.push(`Could not load every filtered task page. ${message}`);
  } finally {
    try {
      await moveToPage(initialPage);
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown restore error.';
      warnings.push(`Could not restore the original filtered task page. ${message}`);
    }
  }

  const tasks = Array.from(tasksById.values());
  if (tasks.length === 0) {
    throw new Error('No tasks were found in the current filtered list.');
  }
  if (pageInfo.expected !== null && tasks.length !== pageInfo.expected) {
    warnings.push(`Loaded ${tasks.length} of ${pageInfo.expected} filtered tasks.`);
  }

  return {
    sourceUrl: parsedUrl.href,
    projectInternalId,
    projectName,
    expectedTaskCount: pageInfo.expected,
    complete: warnings.length === 0 && (pageInfo.expected === null || tasks.length === pageInfo.expected),
    warnings,
    tasks,
  };
}

export function buildOneOfficeFilteredTaskInventory(
  selection: OneOfficeFilteredTaskSelection,
  projectInventory: OneOfficeProjectInventory,
): OneOfficeProjectInventory {
  const detailedTaskById = new Map(
    projectInventory.tasks.map((task) => [task.internalId, task]),
  );
  const missingDetailedIds: string[] = [];
  const tasks = selection.tasks.map((fallbackTask, sourceOrder) => {
    const detailedTask = detailedTaskById.get(fallbackTask.internalId);
    if (!detailedTask) {
      missingDetailedIds.push(fallbackTask.internalId);
    }
    return { ...(detailedTask ?? fallbackTask), sourceOrder };
  });
  const warnings = [...selection.warnings];
  if (missingDetailedIds.length > 0) {
    warnings.push(
      `${missingDetailedIds.length} filtered task(s) were not present in the project inventory; list-row details were used.`,
    );
  }

  return {
    ...projectInventory,
    capturedAt: new Date().toISOString(),
    project: {
      ...projectInventory.project,
      internalId: selection.projectInternalId,
      name: selection.projectName || projectInventory.project.name,
      url: selection.sourceUrl,
    },
    collection: {
      complete: selection.complete,
      expectedTaskCount: selection.expectedTaskCount,
      loadedTaskCount: tasks.length,
      warnings: Array.from(new Set(warnings)),
    },
    tasks,
  };
}
