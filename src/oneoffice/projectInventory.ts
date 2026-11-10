import type {
  OneOfficeProjectInventory,
  OneOfficeProjectTaskInventoryItem,
} from './types';

export interface OneOfficeProjectInventoryOptions {
  allowedOrigin: string;
  sourceUrl?: string;
}

export function collectOneOfficeProjectInventoryFromPage(
  options: OneOfficeProjectInventoryOptions,
): OneOfficeProjectInventory {
  const sourceUrl = options.sourceUrl ?? window.location.href;
  const parsedUrl = new URL(sourceUrl);

  if (parsedUrl.origin !== options.allowedOrigin) {
    throw new Error('This tab is not on the approved 1Office origin.');
  }

  if (parsedUrl.pathname !== '/apps/work-project-project/view') {
    throw new Error('Open a 1Office project before checking its task inventory.');
  }

  const projectInternalId = parsedUrl.searchParams.get('ID')?.trim() ?? '';
  if (!/^\d+$/.test(projectInternalId)) {
    throw new Error('The project ID could not be read from the current page.');
  }

  const normalizeLabel = (value: string): string =>
    value
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/đ/gi, (character) => (character === 'Đ' ? 'D' : 'd'))
      .toLowerCase()
      .replace(/\s+/g, ' ')
      .trim();

  const normalizeText = (value: string | null | undefined): string =>
    (value ?? '').replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();

  const headerElements = Array.from(
    document.querySelectorAll<HTMLElement>('.gantt-header-left-field > .gantt-col'),
  );
  const headers = headerElements.map((element) => normalizeLabel(element.textContent ?? ''));

  if (headers.length === 0) {
    throw new Error("Open the project's Công việc tab before checking its task inventory.");
  }

  const readExpectedTaskCount = (): number | null => {
    for (const element of document.querySelectorAll<HTMLElement>('a, div, span')) {
      const ownText = normalizeText(
        Array.from(element.childNodes)
          .filter((node) => node.nodeType === Node.TEXT_NODE)
          .map((node) => node.textContent ?? '')
          .join(' '),
      );
      const match = /^cong viec\s*\(\s*(\d+)\s*\)$/i.exec(normalizeLabel(ownText));

      if (match) {
        return Number(match[1]);
      }
    }

    return null;
  };

  const readColumn = (columns: Map<string, string>, labels: string[]): string => {
    for (const label of labels) {
      const value = columns.get(label);
      if (value) {
        return value;
      }
    }

    return '';
  };

  const rawTasks: OneOfficeProjectTaskInventoryItem[] = [];
  const seenTaskIds = new Set<string>();
  const warnings: string[] = [];

  for (const row of document.querySelectorAll<HTMLElement>('.gantt-body-row')) {
    const link = row.querySelector<HTMLAnchorElement>(
      'a[href*="/apps/work-task-task/quickview?ID="], a[href*="/apps/work-task-task/view?ID="]',
    );

    if (!link) {
      continue;
    }

    const taskUrl = new URL(link.href, parsedUrl.href);
    if (
      taskUrl.origin !== options.allowedOrigin ||
      !['/apps/work-task-task/quickview', '/apps/work-task-task/view'].includes(taskUrl.pathname)
    ) {
      warnings.push('Ignored a task row because its URL was outside the approved task pages.');
      continue;
    }
    const internalId = taskUrl.searchParams.get('ID')?.trim() ?? '';
    if (!/^\d+$/.test(internalId)) {
      warnings.push('Ignored a task row because its internal ID was missing or invalid.');
      continue;
    }

    if (seenTaskIds.has(internalId)) {
      continue;
    }
    seenTaskIds.add(internalId);

    const values = Array.from(row.querySelectorAll<HTMLElement>(':scope > .gantt-col'));
    const columns = new Map<string, string>();
    headers.forEach((header, index) => {
      if (header) {
        columns.set(header, normalizeText(values[index]?.textContent));
      }
    });

    // 1Office exposes the hierarchy value as No. in the UI, while some DOM
    // layouts place the same value in the hidden "Thu tu" column.
    const taskNo = readColumn(columns, ['no.', 'no', 'thu tu']);
    const taskName = readColumn(columns, ['cong viec']) || normalizeText(link.textContent);
    const taskCode = readColumn(columns, ['ma cong viec']);
    const noParts = taskNo.split('.').map((part) => part.trim()).filter(Boolean);
    const parentNo = noParts.length > 1 ? noParts.slice(0, -1).join('.') : null;

    rawTasks.push({
      sourceOrder: rawTasks.length,
      internalId,
      no: taskNo,
      sourceNo: row.querySelector('.gantt-cf19') ? normalizeText(row.querySelector('.gantt-cf19')?.textContent) : undefined,
      code: taskCode || null,
      name: taskName,
      url: taskUrl.href,
      depth: noParts.length,
      parentNo,
      parentInternalId: null,
    });
  }

  if (rawTasks.length === 0) {
    throw new Error('No project tasks were found in the current Công việc table.');
  }

  const taskIdByNo = new Map(
    rawTasks.filter((task) => task.no).map((task) => [task.no, task.internalId]),
  );
  const tasks = rawTasks.map((task) => ({
    ...task,
    parentInternalId: task.parentNo ? taskIdByNo.get(task.parentNo) ?? null : null,
  }));
  const missingNoCount = tasks.filter((task) => !task.no).length;
  const duplicateNoCount = tasks.length - new Set(tasks.filter((task) => task.no).map((task) => task.no)).size - missingNoCount;
  const expectedTaskCount = readExpectedTaskCount();

  if (missingNoCount > 0) {
    warnings.push(`${missingNoCount} task(s) did not expose a No. value.`);
  }
  if (duplicateNoCount > 0) {
    warnings.push(`${duplicateNoCount} duplicate task No. value(s) were found.`);
  }
  if (expectedTaskCount !== null && tasks.length < expectedTaskCount) {
    warnings.push(`Found ${tasks.length} of ${expectedTaskCount} expected project tasks.`);
  }
  if (expectedTaskCount !== null && tasks.length > expectedTaskCount) {
    warnings.push(
      `Found ${tasks.length} visible tasks while the project counter shows ${expectedTaskCount}.`,
    );
  }

  const projectName = normalizeText(
    document.querySelector<HTMLElement>('#header-title')?.textContent,
  );

  return {
    schemaVersion: 1,
    source: '1office-project-dom',
    capturedAt: new Date().toISOString(),
    project: {
      type: 'project',
      internalId: projectInternalId,
      code: null,
      name: projectName,
      url: parsedUrl.href,
    },
    collection: {
      complete:
        missingNoCount === 0 &&
        duplicateNoCount === 0 &&
        (expectedTaskCount === null || tasks.length >= expectedTaskCount),
      expectedTaskCount,
      loadedTaskCount: tasks.length,
      warnings: Array.from(new Set(warnings)),
    },
    tasks,
  };
}
