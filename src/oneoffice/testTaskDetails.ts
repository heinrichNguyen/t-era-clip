import { TEST_BASELINE_ORIGIN, TEST_BASELINE_PROJECT } from './testBaselineScope';

export interface TestDetailField {
  state: 'value' | 'empty' | 'placeholder' | 'missing' | 'ambiguous';
  displayValue: string | null;
}
export interface TestTaskDetails {
  taskId: string;
  sourceUrl: string;
  capturedAt: string;
  identityVerified: boolean;
  code: TestDetailField;
  no: TestDetailField;
  error: string | null;
  evidenceSource?: 'widget-response';
  failureStage?: string;
  progress?: { state: 'value' | 'missing' | 'ambiguous'; percent: number | null };
}

// Injected function: keep all checks and helpers inside this function.
export async function readTestTaskWidgetsFromPage(taskId: string): Promise<{ general?: string; custom?: string; error?: string }> {
  const current = new URL(window.location.href);
  if (current.origin !== 'https://testoffice.meygroup.vn:9443' || current.pathname !== '/apps/work-project-project/view' || current.searchParams.get('ID') !== '203' || !/^\d+$/.test(taskId)) {
    return { error: 'Tab nguồn không còn ở đúng dự án test 203. Chưa đọc chi tiết.' };
  }
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 30000);
  try {
    // These two read-only widgets were observed on project 203, not inferred from field IDs.
    const read = async (widgetId: string) => {
      if (window.location.href !== current.href) throw new Error('Source changed');
      const url = new URL('/features/core-app-features-board-objects-widgets/data', current.origin);
      const params = { ID: widgetId, module: 'work', mode: 'view', object: 'work-task-task', board_type: 'detail',
        quickview: '1', tab: '3', name: 'visual-3', object_id: taskId, _json: '1', inlineLogin: '1' };
      Object.entries(params).forEach(([key, value]) => url.searchParams.set(key, value));
      const response = await fetch(url.href, { method: 'GET', credentials: 'include', redirect: 'error', signal: controller.signal });
      if (!response.ok) throw new Error('Read failed');
      const reader = response.body?.getReader();
      if (!reader) throw new Error('Missing body');
      const decoder = new TextDecoder();
      let bytes = 0; let text = '';
      try {
        while (true) {
          const part = await reader.read();
          if (part.done) break;
          bytes += part.value.byteLength;
          if (bytes > 5_000_000) { await reader.cancel(); throw new Error('Response too large'); }
          text += decoder.decode(part.value, { stream: true });
        }
        return text + decoder.decode();
      } finally { reader.releaseLock(); }
    };
    const general = await read('23');
    const custom = await read('1915');
    if (window.location.href !== current.href) return { error: 'Tab nguồn đã đổi trong lúc đọc. Chưa sử dụng dữ liệu.' };
    return { general, custom };
  } catch {
    return { error: controller.signal.aborted ? 'Đọc chi tiết quá 30 giây. Không tự gửi lại.' : 'Không đọc được chi tiết: lỗi kết nối hoặc chuyển hướng.' };
  } finally { clearTimeout(timeout); }
}

// Parse declarative server markup inertly. Never render it or execute callbacks.
export function parseTestTaskWidgets(general: string, custom: string, taskId: string): TestTaskDetails {
  const missing = (): TestDetailField => ({ state: 'missing', displayValue: null });
  const result: TestTaskDetails = { taskId, sourceUrl: `${TEST_BASELINE_ORIGIN}/apps/work-task-task/view?ID=${encodeURIComponent(taskId)}`,
    capturedAt: new Date().toISOString(), identityVerified: false, code: missing(), no: missing(), error: null, evidenceSource: 'widget-response' };
  let stage = 'kiểm tra đầu vào';
  const fail = () => ({ ...result, identityVerified: false, code: missing(), no: missing(), progress: { state: 'missing' as const, percent: null }, failureStage: stage,
    error: `Không đọc được chi tiết tại bước ${stage}. Chưa xác nhận đúng công việc, dự án hoặc cấu trúc; không suy diễn dữ liệu.` });
  if (!/^\d+$/.test(taskId) || general.length > 5_000_000 || custom.length > 5_000_000) return fail();
  const parse = (html: string) => new DOMParser().parseFromString(html, 'text/html');
  const elements = (root: ParentNode, depth = 0): Element[] => {
    if (depth > 12) throw new Error('Markup nesting limit');
    root.querySelectorAll('script, style, iframe, object').forEach((node) => node.remove());
    if (root.querySelector('input[type="password"]')) throw new Error('Login');
    return [...root.querySelectorAll('elem')].concat([...root.querySelectorAll('template')]
      .flatMap((node) => elements((node as HTMLTemplateElement).content, depth + 1)));
  };
  const props = (node: Element): Record<string, unknown> => {
    const value: unknown = JSON.parse(node.getAttribute('data-props') ?? '{}');
    if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error('Invalid props');
    return value as Record<string, unknown>;
  };
  const decode = (value: string): string => {
    // Decode only the envelope, never recursively decode nested field attributes.
    const doc = parse(''); const textarea = doc.createElement('textarea');
    textarea.innerHTML = value;
    return textarea.value;
  };
  const bodyFields = (raw: string, id: number, name: string) => {
    stage = `đọc JSON khối ${id}`;
    const payload: unknown = JSON.parse(raw);
    if (!payload || typeof payload !== 'object' || !('content' in payload) || typeof payload.content !== 'string') throw new Error('Missing content');
    stage = `đọc cấu trúc khối ${id}`;
    const widgets = elements(parse(payload.content)).filter((node) => node.getAttribute('name') === 'visual.widget');
    if (widgets.length !== 1) throw new Error('Ambiguous widget');
    const data = props(widgets[0]);
    stage = `xác minh định danh khối ${id}`;
    if (String(data.postId) !== taskId || data.board_id !== id || data.name !== name || data.url !== '/apps/work-task-task/view/widget' || typeof data.body !== 'string') throw new Error('Widget mismatch');
    stage = `đọc các trường khối ${id}`;
    // Parsing data-props already decodes entities in the site's single-encoded envelope.
    // A second decode would turn nested &#039; into attribute delimiters and corrupt JSON.
    const body = data.body.trim();
    const markup = body.startsWith('<elem ') ? body : body.startsWith('&lt;elem ') ? decode(body) : null;
    if (markup === null) throw new Error('Unsupported body encoding');
    return elements(parse(markup)).filter((node) => node.getAttribute('name') === 'detail.field').map(props);
  };
  const normalize = (text: string) => text.replace(/\s+/g, ' ').trim();
  const read = (fields: Record<string, unknown>[], label: string): TestDetailField => {
    const matches = fields.filter((field) => field.label === label);
    if (matches.length > 1) return { state: 'ambiguous', displayValue: null };
    if (!matches.length || typeof matches[0].content !== 'string') return missing();
    const doc = parse(matches[0].content);
    // Code and No. are plain text in the observed wire format. Do not flatten arbitrary components.
    if (doc.body.children.length) return missing();
    const displayValue = normalize(doc.body.textContent ?? '');
    return { state: displayValue === '' ? 'empty' : ['--', '…', '-'].includes(displayValue) ? 'placeholder' : 'value', displayValue };
  };
  try {
    const fields = bodyFields(general, 23, 'Work-Task-TaskNormalInformation');
    const customFields = bodyFields(custom, 1915, 'customField');
    stage = 'xác minh liên kết dự án';
    const projects = fields.filter((field) => field.label === 'Dự án');
    if (projects.length !== 1 || typeof projects[0].content !== 'string') return fail();
    const links = elements(parse(projects[0].content)).filter((node) => node.getAttribute('name') === 'a');
    const projectUrls = links.map((node) => {
      const atts = JSON.parse(node.getAttribute('data-atts') ?? '{}');
      if (typeof atts.href !== 'string') throw new Error('Missing project href');
      return new URL(atts.href, TEST_BASELINE_ORIGIN);
    });
    if (projectUrls.length !== 1 || projectUrls[0].origin !== TEST_BASELINE_ORIGIN || projectUrls[0].pathname !== '/apps/work-project-project/view' || projectUrls[0].searchParams.getAll('ID').length !== 1 || projectUrls[0].searchParams.get('ID') !== TEST_BASELINE_PROJECT) return fail();
    result.identityVerified = true;
    stage = 'đọc mã và No.';
    result.code = read(fields, 'Mã công việc'); result.no = read(customFields, 'No.');
    const progressFields = fields.filter((field) => field.label === 'Tiến độ');
    result.progress = { state: progressFields.length > 1 ? 'ambiguous' : 'missing', percent: null };
    if (progressFields.length === 1 && typeof progressFields[0].content === 'string') {
      const components = elements(parse(progressFields[0].content)).filter((node) => node.getAttribute('name') === 'progress');
      if (components.length > 1) result.progress.state = 'ambiguous';
      if (components.length === 1) {
        const data = props(components[0]);
        if (data.type === 'progresstask' && typeof data.percent === 'number' && Number.isFinite(data.percent) && data.percent >= 0 && data.percent <= 100) {
          result.progress = { state: 'value', percent: data.percent };
        }
      }
    }
    if (result.code.state !== 'value' || !['value', 'empty'].includes(result.no.state)) result.error = 'Một số trường thiếu, không rõ hoặc chỉ là ký hiệu hiển thị. Chưa đủ dữ liệu để ghi.';
    return result;
  } catch { return fail(); }
}

export function parseTestTaskDetails(html: string, taskId: string): TestTaskDetails {
  const missing = (): TestDetailField => ({ state: 'missing', displayValue: null });
  const result: TestTaskDetails = { taskId, sourceUrl: `${TEST_BASELINE_ORIGIN}/apps/work-task-task/view?ID=${encodeURIComponent(taskId)}`,
    capturedAt: new Date().toISOString(), identityVerified: false, code: missing(), no: missing(), error: null };
  if (!/^\d+$/.test(taskId)) return { ...result, error: 'ID mẫu không hợp lệ.' };
  const doc = new DOMParser().parseFromString(html, 'text/html');
  doc.querySelectorAll('script, style, template, iframe, object').forEach((node) => node.remove());
  if (doc.querySelector('input[type="password"]')) return { ...result, error: 'Nguồn trả về trang đăng nhập, chưa đọc chi tiết.' };
  const discussions = [...doc.querySelectorAll('.cards.comments[comments-id]')];
  const expectedId = `work-task-task@${taskId}`;
  const matching = discussions.filter((node) => node.getAttribute('comments-id') === expectedId);
  const other = discussions.some((node) => /^work-task-task@/.test(node.getAttribute('comments-id') ?? '') && node.getAttribute('comments-id') !== expectedId);
  const projects = [...doc.querySelectorAll<HTMLAnchorElement>('a[href]')].flatMap((link) => {
    try { const url = new URL(link.getAttribute('href')!, result.sourceUrl);
      return url.pathname === '/apps/work-project-project/view' ? [url] : [];
    } catch { return []; }
  });
  if (matching.length !== 1 || other || !projects.some((url) => url.origin === TEST_BASELINE_ORIGIN && url.searchParams.get('ID') === TEST_BASELINE_PROJECT)) {
    return { ...result, error: 'HTML chưa xác nhận được đúng ID công việc và dự án 203. Không suy diễn trường còn thiếu.' };
  }
  result.identityVerified = true;
  const normalize = (text: string) => text.replace(/\s+/g, ' ').trim();
  const read = (label: string): TestDetailField => {
    const fields = [...doc.querySelectorAll('.detail-field')].filter((field) => normalize(field.querySelector('.detail-field-label')?.textContent ?? '') === label);
    if (!fields.length) return missing();
    if (fields.length !== 1) return { state: 'ambiguous', displayValue: null };
    const content = fields[0].querySelector('.detail-field-content');
    if (!content) return missing();
    content.querySelectorAll('form, button, input, textarea').forEach((node) => node.remove());
    const displayValue = normalize(content.textContent ?? '');
    return { state: ['', '--', '…', '-'].includes(displayValue) ? 'placeholder' : 'value', displayValue };
  };
  result.code = read('Mã công việc'); result.no = read('No.');
  if (result.code.state !== 'value' || result.no.state !== 'value') {
    result.error = 'Một số trường thiếu hoặc chỉ có ký hiệu hiển thị. Chưa xác minh giá trị gốc để ghi.';
  }
  return result;
}

export async function collectTestTaskDetails(tabId: number, taskId: string): Promise<TestTaskDetails> {
  const api = (globalThis as any).chrome;
  if (!api?.scripting?.executeScript) throw new Error('Không có quyền đọc tab nguồn.');
  const results = await api.scripting.executeScript({ target: { tabId }, func: readTestTaskWidgetsFromPage, args: [taskId] });
  const response = results?.[0]?.result as { general?: string; custom?: string; error?: string } | undefined;
  if (typeof response?.general !== 'string' || typeof response?.custom !== 'string') throw new Error(response?.error || 'Không nhận được chi tiết mẫu.');
  return parseTestTaskWidgets(response.general, response.custom, taskId);
}
