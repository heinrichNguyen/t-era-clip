import type {
  OneOfficeProjectInventory,
  OneOfficeProjectTaskInventoryItem,
} from './types';

export interface OneOfficeProjectInventoryApiOptions {
  allowedOrigin: string;
  sourceUrl?: string;
}

interface GanttTask {
  ID?: unknown;
  parent_id?: unknown;
  phase_id?: unknown;
  title?: unknown;
  code?: unknown;
  start_plan?: unknown;
  end_plan?: unknown;
  start?: unknown;
  end?: unknown;
  status?: unknown;
  out_date?: unknown;
  fields?: unknown;
}

interface GanttPhase {
  phase_db_id?: unknown;
}

interface GanttStructureResponse {
  tasks?: unknown;
  phases?: unknown;
  groupType?: unknown;
}

interface GanttEnrichmentResponse {
  enriched?: unknown;
}

interface UserLookupResponse {
  items?: unknown;
}

export async function collectOneOfficeProjectInventoryFromApi(
  options: OneOfficeProjectInventoryApiOptions,
): Promise<OneOfficeProjectInventory> {
  const enrichmentBatchSize = 50;
  const enrichmentConcurrency = 4;
  const sourceUrl = options.sourceUrl ?? window.location.href;
  const parsedUrl = new URL(sourceUrl);

  if (window.location.origin !== options.allowedOrigin || parsedUrl.origin !== options.allowedOrigin) {
    throw new Error('This tab is not on the approved 1Office origin.');
  }

  if (parsedUrl.pathname !== '/apps/work-project-project/view') {
    throw new Error('Open a 1Office project before checking its task inventory.');
  }

  const projectInternalId = parsedUrl.searchParams.get('ID')?.trim() ?? '';
  if (!/^\d+$/.test(projectInternalId)) {
    throw new Error('The project ID could not be read from the current page.');
  }

  const normalizeText = (value: unknown): string =>
    (typeof value === 'string' || typeof value === 'number' ? String(value) : '')
      .replace(/\u00a0/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();

  const decodeHtml = (value: string): string => {
    const textarea = document.createElement('textarea');
    textarea.innerHTML = value;
    return textarea.value;
  };

  const readElementProps = (value: unknown): Record<string, unknown> | null => {
    if (typeof value !== 'string') {
      return null;
    }

    const match = /data-props=(['"])([\s\S]*?)\1/.exec(value);
    if (!match) {
      return null;
    }

    try {
      const parsed = JSON.parse(decodeHtml(match[2]));
      return parsed && typeof parsed === 'object' && !Array.isArray(parsed)
        ? parsed as Record<string, unknown>
        : null;
    } catch {
      return null;
    }
  };

  const stripMarkup = (value: unknown): string => {
    if (typeof value !== 'string' && typeof value !== 'number') {
      return '';
    }

    const container = document.createElement('div');
    container.innerHTML = decodeHtml(String(value))
      .replace(/<br\s*\/?\s*>/gi, '\n')
      .replace(/<\/(?:div|li|p)>/gi, '\n');
    return normalizeText(container.textContent);
  };

  const uniqueValues = (values: string[]): string[] =>
    Array.from(new Set(values.map(normalizeText).filter(Boolean)));

  const collectNamedValues = (value: unknown): string[] => {
    const names: string[] = [];
    const visit = (candidate: unknown): void => {
      if (Array.isArray(candidate)) {
        candidate.forEach(visit);
        return;
      }
      if (!candidate || typeof candidate !== 'object') {
        return;
      }

      const record = candidate as Record<string, unknown>;
      for (const key of ['name', 'label', 'title']) {
        const namedValue = normalizeText(record[key]);
        if (namedValue) {
          names.push(namedValue);
        }
      }
      for (const key of ['user', 'users', 'items', 'values']) {
        visit(record[key]);
      }
    };

    visit(readElementProps(value));
    return uniqueValues(names);
  };

  const readUserIds = (value: unknown): string[] => {
    const ids: string[] = [];
    const visit = (candidate: unknown): void => {
      if (Array.isArray(candidate)) {
        candidate.forEach(visit);
        return;
      }
      if (typeof candidate === 'string' || typeof candidate === 'number') {
        String(candidate).split(',').forEach((part) => {
          const id = part.trim();
          if (/^\d+$/.test(id)) {
            ids.push(id);
          }
        });
        return;
      }
      if (!candidate || typeof candidate !== 'object') {
        return;
      }

      const record = candidate as Record<string, unknown>;
      visit(record.ID);
      visit(record.id);
      for (const key of ['user', 'users', 'items', 'values', 'ids', 'value']) {
        visit(record[key]);
      }
    };

    visit(readElementProps(value));
    return uniqueValues(ids);
  };

  const readDisplayValues = (value: unknown): string[] => {
    const namedValues = collectNamedValues(value);
    if (namedValues.length > 0) {
      return namedValues;
    }

    const props = readElementProps(value);
    const content = props?.content;
    const renderedContent = stripMarkup(content);
    if (renderedContent && !/^\d+(?:\s*[,;]\s*\d+)*$/.test(renderedContent)) {
      return uniqueValues(renderedContent.split(/\s*[;\n]+\s*/));
    }

    const renderedValue = stripMarkup(value);
    return renderedValue && renderedValue !== '…'
      ? uniqueValues(renderedValue.split(/\s*[;\n]+\s*/))
      : [];
  };

  const readFieldRecord = (value: unknown): Record<string, unknown> =>
    value && typeof value === 'object' && !Array.isArray(value)
      ? value as Record<string, unknown>
      : {};

  const readFieldText = (fields: Record<string, unknown>, key: string): string => {
    const value = fields[key];
    const namedValues = collectNamedValues(value);
    if (namedValues.length > 0) {
      return namedValues.join('; ');
    }
    return stripMarkup(value);
  };

  const readExpectedTaskCount = (): number | null => {
    const normalizeLabel = (value: string): string =>
      value
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/đ/gi, (character) => (character === 'Đ' ? 'D' : 'd'))
        .toLowerCase()
        .replace(/\s+/g, ' ')
        .trim();

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

  const fetchJson = async (url: string, init?: RequestInit): Promise<unknown> => {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 30000);
    try {
      const response = await fetch(url, {
        credentials: 'include',
        ...init,
        redirect: 'error',
        signal: controller.signal,
        headers: {
          'X-Requested-With': 'XMLHttpRequest',
          ...(init?.headers ?? {}),
        },
      });

      if (!response.ok) {
        throw new Error(`Yêu cầu 1Office: HTTP ${response.status}. Hãy kiểm tra phiên đăng nhập và thử lại.`);
      }
      return await response.json();
    } catch (error) {
      if (error instanceof Error && error.message.startsWith('Yêu cầu 1Office:')) throw error;
      throw new Error(controller.signal.aborted
        ? 'Yêu cầu 1Office: máy chủ không phản hồi trong 30 giây. Hãy thử lại.'
        : 'Yêu cầu 1Office: lỗi kết nối, dữ liệu trả về hoặc chuyển hướng. Kiểm tra phiên đăng nhập và thử lại.');
    } finally {
      clearTimeout(timeout);
    }
  };

  const ganttUrl = new URL('/apps/work-project-project/gantt', parsedUrl.origin);
  ganttUrl.searchParams.set('project_id', projectInternalId);
  const structureUrl = new URL(ganttUrl.href);
  structureUrl.searchParams.set('structure', '1');
  structureUrl.searchParams.set('group_type', 'phase');

  const rawStructure = (await fetchJson(structureUrl.href)) as GanttStructureResponse;
  if (!Array.isArray(rawStructure.tasks) || !Array.isArray(rawStructure.phases)) {
    throw new Error('The 1Office Gantt response did not contain a project task structure.');
  }

  const warnings: string[] = [];
  const seenTaskIds = new Set<string>();
  const rawTasks = (rawStructure.tasks as GanttTask[]).flatMap((task, sourceIndex) => {
    const internalId = normalizeText(task.ID);
    if (!/^\d+$/.test(internalId)) {
      warnings.push('Ignored a Gantt task because its internal ID was missing or invalid.');
      return [];
    }
    if (seenTaskIds.has(internalId)) {
      warnings.push(`Ignored duplicate Gantt task ID ${internalId}.`);
      return [];
    }
    seenTaskIds.add(internalId);

    const normalizedParentId = normalizeText(task.parent_id);
    const normalizedPhaseId = normalizeText(task.phase_id);
    const parentInternalId = normalizedParentId && normalizedParentId !== '0'
      ? normalizedParentId
      : null;
    const phaseId = normalizedPhaseId && normalizedPhaseId !== '0' ? normalizedPhaseId : null;
    return [{
      sourceIndex,
      internalId,
      parentInternalId,
      phaseId,
      name: normalizeText(task.title),
      code: normalizeText(task.code) || null,
      startPlan: normalizeText(task.start_plan),
      endPlan: normalizeText(task.end_plan),
      actualStart: normalizeText(task.start),
      actualEnd: normalizeText(task.end),
      status: normalizeText(task.status),
      overdueDays: normalizeText(task.out_date),
      fields: readFieldRecord(task.fields),
    }];
  });

  if (rawTasks.length === 0) {
    throw new Error('No project tasks were found in the 1Office Gantt response.');
  }

  const enrichmentByTaskId = new Map<string, Record<string, unknown>>();
  const batches: string[][] = [];
  for (let index = 0; index < rawTasks.length; index += enrichmentBatchSize) {
    batches.push(rawTasks.slice(index, index + enrichmentBatchSize).map((task) => task.internalId));
  }

  const loadEnrichmentBatch = async (ids: string[]): Promise<void> => {
    const body = new URLSearchParams({
      mode: 'enrich',
      project_id: projectInternalId,
      ids: JSON.stringify(ids),
    });
    const rawEnrichment = (await fetchJson(ganttUrl.href, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8' },
      body: body.toString(),
    })) as GanttEnrichmentResponse;
    const enriched = rawEnrichment.enriched;
    if (!enriched || typeof enriched !== 'object' || Array.isArray(enriched)) {
      throw new Error('The 1Office Gantt enrichment response was invalid.');
    }

    for (const id of ids) {
      const record = (enriched as Record<string, unknown>)[id];
      if (record && typeof record === 'object' && !Array.isArray(record)) {
        enrichmentByTaskId.set(id, record as Record<string, unknown>);
      }
    }
  };

  try {
    for (let index = 0; index < batches.length; index += enrichmentConcurrency) {
      await Promise.all(batches.slice(index, index + enrichmentConcurrency).map(loadEnrichmentBatch));
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Unknown enrichment error.';
    if (message.startsWith('Yêu cầu 1Office:')) throw error;
    warnings.push(`Could not load all task No. values. ${message}`);
  }

  if (enrichmentByTaskId.size < rawTasks.length) {
    warnings.push(
      `Loaded detailed task fields for ${enrichmentByTaskId.size} of ${rawTasks.length} tasks.`,
    );
  }

  const userNameById = new Map<string, string>();
  const userIds = uniqueValues(Array.from(enrichmentByTaskId.values()).flatMap((record) => {
    const fields = readFieldRecord(record.fields);
    return [
      ...readUserIds(record.assign_names_html),
      ...readUserIds(record.assign_ids_html),
      ...readUserIds(fields.assign_ids),
      ...readUserIds(fields.owner_ids),
    ];
  }));
  const userLookupBatchSize = 100;
  try {
    for (let index = 0; index < userIds.length; index += userLookupBatchSize) {
      const query = new URLSearchParams();
      userIds.slice(index, index + userLookupBatchSize).forEach((id) => query.append('ids[]', id));
      const lookupUrl = new URL('/app/autofill/users/items', parsedUrl.origin);
      lookupUrl.search = query.toString();
      const response = (await fetchJson(lookupUrl.href)) as UserLookupResponse;
      if (!Array.isArray(response.items)) {
        continue;
      }
      for (const item of response.items) {
        const record = readFieldRecord(item);
        const id = normalizeText(record.ID || record.id);
        const name = normalizeText(record.name);
        if (/^\d+$/.test(id) && name) {
          userNameById.set(id, name);
        }
      }
    }
  } catch {
    // Keep the names embedded in the Gantt response when user lookup is unavailable.
  }

  const readUserNames = (value: unknown): string[] => {
    const namesById = readUserIds(value)
      .map((id) => userNameById.get(id) ?? '')
      .filter(Boolean);
    return uniqueValues([...namesById, ...collectNamedValues(value)]);
  };

  const tasksById = new Map(rawTasks.map((task) => [task.internalId, task]));
  const childrenByParentId = new Map<string, typeof rawTasks>();
  const roots: typeof rawTasks = [];

  for (const task of rawTasks) {
    if (task.parentInternalId && tasksById.has(task.parentInternalId)) {
      const siblings = childrenByParentId.get(task.parentInternalId) ?? [];
      siblings.push(task);
      childrenByParentId.set(task.parentInternalId, siblings);
    } else {
      if (task.parentInternalId) {
        warnings.push(`Task ${task.internalId} references missing parent ${task.parentInternalId}.`);
      }
      roots.push(task);
    }
  }

  const phaseIds = (rawStructure.phases as GanttPhase[])
    .map((phase) => normalizeText(phase.phase_db_id))
    .filter(Boolean);
  const rootsByPhase = new Map<string, typeof rawTasks>();
  const otherRoots: typeof rawTasks = [];
  for (const root of roots) {
    if (root.phaseId && phaseIds.includes(root.phaseId)) {
      const phaseRoots = rootsByPhase.get(root.phaseId) ?? [];
      phaseRoots.push(root);
      rootsByPhase.set(root.phaseId, phaseRoots);
    } else {
      otherRoots.push(root);
    }
  }

  const outlineByTaskId = new Map<string, string>();
  const orderedTaskIds: string[] = [];
  const visiting = new Set<string>();
  const visited = new Set<string>();

  const visitTask = (task: (typeof rawTasks)[number], outline: string): void => {
    if (visiting.has(task.internalId)) {
      warnings.push(`Task hierarchy cycle detected at task ${task.internalId}.`);
      return;
    }
    if (visited.has(task.internalId)) {
      return;
    }

    visiting.add(task.internalId);
    visited.add(task.internalId);
    outlineByTaskId.set(task.internalId, outline);
    orderedTaskIds.push(task.internalId);
    const children = childrenByParentId.get(task.internalId) ?? [];
    children.forEach((child, index) => visitTask(child, `${outline}.${index + 1}`));
    visiting.delete(task.internalId);
  };

  phaseIds.forEach((phaseId, phaseIndex) => {
    const phaseRoots = rootsByPhase.get(phaseId) ?? [];
    phaseRoots.forEach((root, rootIndex) => visitTask(root, `${phaseIndex + 1}.${rootIndex + 1}`));
  });
  if (otherRoots.length > 0) {
    const otherGroupIndex = phaseIds.length + 1;
    otherRoots.forEach((root, rootIndex) =>
      visitTask(root, `${otherGroupIndex}.${rootIndex + 1}`),
    );
  }
  for (const task of rawTasks) {
    if (!visited.has(task.internalId)) {
      warnings.push(`Task ${task.internalId} could not be placed in the project hierarchy.`);
      visitTask(task, String(orderedTaskIds.length + 1));
    }
  }

  const readCustomNo = (taskId: string): string => {
    const enrichment = enrichmentByTaskId.get(taskId);
    const fields = enrichment?.fields;
    if (!fields || typeof fields !== 'object' || Array.isArray(fields)) {
      return '';
    }
    const customNo = normalizeText((fields as Record<string, unknown>).cf19);
    return customNo === '…' || customNo === '-' ? '' : customNo;
  };

  const readEnrichment = (taskId: string): {
    fields: Record<string, unknown>;
    record: Record<string, unknown>;
  } => {
    const record = enrichmentByTaskId.get(taskId) ?? {};
    return { record, fields: readFieldRecord(record.fields) };
  };

  const chosenNoByTaskId = new Map(
    rawTasks.map((task) => [
      task.internalId,
      readCustomNo(task.internalId) || outlineByTaskId.get(task.internalId) || '',
    ]),
  );
  const tasks: OneOfficeProjectTaskInventoryItem[] = orderedTaskIds.map((taskId, sourceOrder) => {
    const task = tasksById.get(taskId)!;
    const { record: enrichment, fields: enrichedFields } = readEnrichment(taskId);
    const taskNo = chosenNoByTaskId.get(taskId) ?? '';
    const parentNo = task.parentInternalId
      ? chosenNoByTaskId.get(task.parentInternalId) ?? null
      : null;

    return {
      sourceOrder,
      internalId: task.internalId,
      no: taskNo,
      sourceNo: Object.prototype.hasOwnProperty.call(enrichedFields, 'cf19') ? readCustomNo(taskId) : undefined,
      code: task.code,
      name: task.name,
      url: `${parsedUrl.origin}/apps/work-task-task/view?ID=${encodeURIComponent(task.internalId)}`,
      depth: (outlineByTaskId.get(task.internalId) ?? taskNo).split('.').filter(Boolean).length,
      parentNo,
      parentInternalId: task.parentInternalId,
      assigneeNames: readUserNames(
        enrichment.assign_names_html || enrichment.assign_ids_html || enrichedFields.assign_ids,
      ),
      assignerNames: readUserNames(enrichedFields.owner_ids),
      departmentNames: readDisplayValues(enrichedFields.department_ids),
      taskStatus:
        readFieldText(enrichedFields, 'status') ||
        normalizeText(task.fields.status) ||
        task.status,
      baselineStart: readFieldText(enrichedFields, 'cf171'),
      baselineEnd: readFieldText(enrichedFields, 'cf174'),
      plannedStart: normalizeText(task.fields.start_plan) || task.startPlan,
      plannedEnd: normalizeText(task.fields.end_plan) || task.endPlan,
      actualStart: normalizeText(task.fields.start) || task.actualStart,
      actualEnd: normalizeText(task.fields.end) || task.actualEnd,
      overdueDays:
        readFieldText(enrichedFields, 'out_date') ||
        normalizeText(task.fields.out_date) ||
        task.overdueDays,
      completedOverdueDays: readFieldText(enrichedFields, 'out_date_completed'),
      labels: readDisplayValues(enrichedFields.tag_ids),
    };
  });

  const expectedTaskCount = readExpectedTaskCount();
  const missingSourceNo = tasks.filter((task) => task.sourceNo === undefined).length;
  if (missingSourceNo > 0) {
    warnings.push(`Chưa xác minh được No. gốc của ${missingSourceNo}/${tasks.length} công việc từ nguồn chi tiết. Số thứ tự gợi ý không thay thế No. gốc; chưa đủ dữ liệu để ghi No.`);
  }
  const missingCode = tasks.filter((task) => !task.code).length;
  if (missingCode > 0) {
    warnings.push(`Chưa đọc được mã công việc của ${missingCode}/${tasks.length} công việc từ nguồn danh sách. Không kết luận mã trên 1Office đang trống.`);
  }
  if (expectedTaskCount !== null && tasks.length !== expectedTaskCount) {
    warnings.push(`Loaded ${tasks.length} of ${expectedTaskCount} expected project tasks.`);
  }
  const duplicateNoCount = tasks.length - new Set(tasks.map((task) => task.no)).size;
  if (duplicateNoCount > 0) {
    warnings.push(`${duplicateNoCount} duplicate task No. value(s) were found.`);
  }

  const projectName = normalizeText(
    document.querySelector<HTMLElement>('#header-title')?.textContent,
  );
  // Inventory completeness describes whether all project tasks were loaded.
  // Optional enrichment can warn about missing export fields without blocking collection.
  const complete = expectedTaskCount === null || tasks.length === expectedTaskCount;

  return {
    schemaVersion: 1,
    source: '1office-project-api',
    capturedAt: new Date().toISOString(),
    project: {
      type: 'project',
      internalId: projectInternalId,
      code: null,
      name: projectName,
      url: parsedUrl.href,
    },
    collection: {
      complete,
      expectedTaskCount,
      loadedTaskCount: tasks.length,
      warnings: Array.from(new Set(warnings)),
    },
    tasks,
  };
}
