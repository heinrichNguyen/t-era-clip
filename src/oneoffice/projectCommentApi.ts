import { oneOfficeOrigin } from './origins';
import { parseOneOfficeTimestamp } from './timestamp';
import type {
  OneOfficeAttachment,
  OneOfficeComment,
  OneOfficeDiscussionExport,
  OneOfficeMention,
  OneOfficeProjectTaskInventoryItem,
} from './types';

interface OneOfficeCommentApiFile {
  ID?: number | string;
  name?: string;
  thumb?: string;
  url?: string;
  size?: string;
}

interface OneOfficeCommentApiUser {
  ID?: number | string;
  name?: string;
}

interface OneOfficeCommentApiItem {
  ID?: number | string;
  parent_id?: number | string;
  dateTime?: string;
  date_time?: string;
  content?: string;
  files?: OneOfficeCommentApiFile[];
  user?: OneOfficeCommentApiUser;
  subs?: OneOfficeCommentApiItem[];
  total_subs?: number;
}

export interface OneOfficeCommentApiResponse {
  requestError?: string;
  items?: OneOfficeCommentApiItem[];
  total?: number;
}

export interface OneOfficeCommentApiRequestOptions {
  allowedOrigin: string;
  taskId: string;
  offset: number;
  object?: 'work-task-task' | 'work-normal-normal';
}

const chromeApi = typeof globalThis !== 'undefined' ? (globalThis as any).chrome : undefined;

export async function fetchOneOfficeCommentsFromPage(
  options: OneOfficeCommentApiRequestOptions,
): Promise<OneOfficeCommentApiResponse | null> {
  if (window.location.origin !== options.allowedOrigin) {
    return null;
  }

  const url = new URL('/comment/post/gets', options.allowedOrigin);
  if (options.object && options.object !== 'work-task-task' && options.object !== 'work-normal-normal') return null;
  url.searchParams.set('object', options.object ?? 'work-task-task');
  url.searchParams.set('object_id', options.taskId);
  url.searchParams.set('_offset', String(options.offset));
  url.searchParams.set('_json', '1');
  url.searchParams.set('reloadCsrfToken', '0');
  url.searchParams.set('_dir', 'next');
  url.searchParams.set('_start', '0');
  url.searchParams.set('_ajax', '1');
  url.searchParams.set('inlineLogin', '1');

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 30000);
  try {
    const response = await fetch(url.href, {
      method: 'GET',
      credentials: 'include',
      redirect: 'error',
      signal: controller.signal,
      headers: {
        Accept: 'application/json, text/javascript, */*; q=0.01',
        'X-Requested-With': 'XMLHttpRequest',
      },
    });

    if (!response.ok) {
      return { requestError: `Yêu cầu 1Office: HTTP ${response.status}. Hãy kiểm tra phiên đăng nhập và thử lại.` };
    }

    return (await response.json()) as OneOfficeCommentApiResponse;
  } catch {
    return { requestError: controller.signal.aborted
      ? 'Yêu cầu 1Office: máy chủ không phản hồi trong 30 giây. Hãy thử lại.'
      : 'Yêu cầu 1Office: lỗi kết nối, dữ liệu trả về hoặc chuyển hướng. Kiểm tra phiên đăng nhập và thử lại.' };
  } finally {
    clearTimeout(timeout);
  }
}

const normalizeInlineText = (value: string | null | undefined): string =>
  (value ?? '').replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();

const readStructuredText = (value: string): string => {
  const parsed = new DOMParser().parseFromString(`<main>${value}</main>`, 'text/html');
  const root = parsed.querySelector('main');
  if (!root) {
    return normalizeInlineText(value);
  }

  const blockElements = new Set([
    'ADDRESS',
    'ARTICLE',
    'BLOCKQUOTE',
    'DIV',
    'H1',
    'H2',
    'H3',
    'H4',
    'H5',
    'H6',
    'LI',
    'P',
    'PRE',
  ]);
  let output = '';

  const visit = (node: Node): void => {
    if (node.nodeType === Node.TEXT_NODE) {
      output += node.textContent ?? '';
      return;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) {
      return;
    }

    const element = node as Element;
    if (element.tagName === 'BR') {
      output += '\n';
      return;
    }

    const isBlock = blockElements.has(element.tagName);
    if (isBlock && output && !output.endsWith('\n')) {
      output += '\n';
    }
    element.childNodes.forEach(visit);
    if (isBlock && !output.endsWith('\n')) {
      output += '\n';
    }
  };

  root.childNodes.forEach(visit);
  return output
    .replace(/\u00a0/g, ' ')
    .split('\n')
    .map((line) => line.replace(/[\t ]+/g, ' ').trim())
    .join('\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
};

const readMentions = (
  content: string,
  contentText: string,
  parentAuthor: OneOfficeCommentApiUser | null,
): OneOfficeMention[] => {
  const mentions = new Map<string, OneOfficeMention>();
  const parsed = new DOMParser().parseFromString(`<main>${content}</main>`, 'text/html');

  for (const anchor of parsed.querySelectorAll<HTMLAnchorElement>('a.userlink')) {
    const displayName = normalizeInlineText(anchor.textContent);
    if (!displayName) {
      continue;
    }
    const userId = /[?&]ID=(\d+)/i.exec(anchor.getAttribute('href') ?? '')?.[1] ?? null;
    mentions.set(`${userId ?? ''}:${displayName}`, { userId, displayName });
  }

  for (const match of contentText.matchAll(
    /@([\p{Lu}Đ][\p{L}'-]*(?:\s+[\p{Lu}Đ][\p{L}'-]*){0,4})/gu,
  )) {
    const displayName = normalizeInlineText(match[1]);
    if (displayName) {
      mentions.set(`:${displayName}`, { userId: null, displayName });
    }
  }

  const parentName = normalizeInlineText(parentAuthor?.name);
  const parentId = parentAuthor?.ID == null ? null : String(parentAuthor.ID);
  if (parentName && contentText.startsWith(parentName)) {
    mentions.set(`${parentId ?? ''}:${parentName}`, {
      userId: parentId,
      displayName: parentName,
    });
  }

  return Array.from(mentions.values());
};

const readAttachments = (
  files: OneOfficeCommentApiFile[] | undefined,
  sourceUrl: string,
): OneOfficeAttachment[] =>
  (files ?? []).map((file, index) => ({
    index,
    name: normalizeInlineText(file.name),
    sizeText: normalizeInlineText(file.size) || null,
    thumbnailUrl: file.thumb ? new URL(file.thumb, sourceUrl).href : null,
    originalDownloadAvailable: Boolean(file.url),
  }));

const toId = (value: number | string | undefined): string =>
  value == null ? '' : String(value).trim();

export function parseOneOfficeCommentApiResponse(
  response: OneOfficeCommentApiResponse,
  task: OneOfficeProjectTaskInventoryItem,
): OneOfficeDiscussionExport | null {
  if (!Array.isArray(response.items) || !Number.isInteger(response.total) || response.total! < 0) {
    return null;
  }

  if (response.items.length !== response.total) {
    return null;
  }

  const rootIds = new Set<string>();
  const allIds = new Set<string>();
  for (const root of response.items) {
    const rootId = toId(root.ID);
    const replies = Array.isArray(root.subs) ? root.subs : [];
    if (
      !/^\d+$/.test(rootId) ||
      rootIds.has(rootId) ||
      !Number.isInteger(root.total_subs ?? 0) ||
      replies.length !== (root.total_subs ?? 0)
    ) {
      return null;
    }
    rootIds.add(rootId);
    allIds.add(rootId);

    for (const reply of replies) {
      const replyId = toId(reply.ID);
      if (!/^\d+$/.test(replyId) || allIds.has(replyId)) {
        return null;
      }
      allIds.add(replyId);
    }
  }

  const comments: OneOfficeComment[] = [];
  const warnings: string[] = [];
  let sourceOrder = 0;
  const origin = oneOfficeOrigin(task.url);
  const taskUrl = new URL(task.url).pathname === '/apps/work-normal-normal/view' ? task.url
    : `${origin}/apps/work-task-task/view?ID=${encodeURIComponent(task.internalId)}`;

  const appendComment = (
    item: OneOfficeCommentApiItem,
    parentId: string | null,
    rootId: string,
    depth: number,
    parentAuthor: OneOfficeCommentApiUser | null,
  ): void => {
    const id = toId(item.ID);
    const createdAtRaw = normalizeInlineText(item.dateTime ?? item.date_time);
    const createdAtIso = parseOneOfficeTimestamp(createdAtRaw);
    const contentText = readStructuredText(item.content ?? '');

    if (!createdAtIso) {
      warnings.push(`Comment ${id} has an invalid or missing absolute timestamp.`);
    }

    comments.push({
      id,
      parentId,
      rootId,
      depth,
      sourceOrder: sourceOrder++,
      authorId: item.user?.ID == null ? null : String(item.user.ID),
      authorName: normalizeInlineText(item.user?.name),
      createdAtRaw,
      createdAtIso,
      contentText,
      mentions: readMentions(item.content ?? '', contentText, parentAuthor),
      attachments: readAttachments(item.files, taskUrl),
    });
  };

  for (const root of response.items) {
    const rootId = toId(root.ID);
    appendComment(root, null, rootId, 0, null);

    for (const reply of root.subs ?? []) {
      const replyParentId = toId(reply.parent_id) || rootId;
      appendComment(reply, replyParentId, rootId, replyParentId === rootId ? 1 : 2, root.user ?? null);
    }
  }

  return {
    schemaVersion: 1,
    source: '1office-comment-api',
    capturedAt: new Date().toISOString(),
    entity: {
      type: 'task',
      internalId: task.internalId,
      code: task.code,
      name: task.name,
      url: taskUrl,
    },
    collection: {
      complete: true,
      loadMoreClicks: 0,
      loadedRootCommentCount: response.items.length,
      loadedReplyCount: comments.length - response.items.length,
      expectedRootCommentCount: response.total,
      warnings,
    },
    comments,
  };
}

export async function collectOneOfficeTaskFromCommentApi(
  tabId: number,
  task: OneOfficeProjectTaskInventoryItem,
): Promise<OneOfficeDiscussionExport | null> {
  if (!chromeApi?.scripting?.executeScript) {
    throw new Error('Chrome scripting access is unavailable. Reload the extension and try again.');
  }

  const injectionResults = await chromeApi.scripting.executeScript({
    target: { tabId },
    func: fetchOneOfficeCommentsFromPage,
    args: [
      {
        allowedOrigin: oneOfficeOrigin(task.url),
        taskId: task.internalId,
        ...(new URL(task.url).pathname === '/apps/work-normal-normal/view' ? { object: 'work-normal-normal' } : {}),
        offset: 1000,
      },
    ],
  });
  const response = injectionResults?.[0]?.result as OneOfficeCommentApiResponse | null | undefined;
  if (response?.requestError) throw new Error(response.requestError);

  return response ? parseOneOfficeCommentApiResponse(response, task) : null;
}
