import { buildNoUpdatePlan } from './noUpdatePlan';
import type { OneOfficeProjectInventory } from './types';
import { isNoTestTask } from './noTestScope';
import { ONEOFFICE_ALLOWED_ORIGIN, oneOfficeOrigin } from './origins';
import { finishNoWriteJournal, prepareNoWriteJournal, readNoWriteJournal, reconcileNoWriteJournal, withNoWriteLock } from './noWriteJournal';

export interface NoWriteRequest {
  silentReads?: boolean;
  directEdit?: boolean;
  changes: { internalId: string; code: string; before: string; after: string }[];
  sourceTasks: { internalId: string; no: string }[];
}

export interface NoExperimentResult {
  mode?: 'experiment' | 'plan';
  readOnly: boolean;
  complete: boolean;
  error: string | null;
  phase: string;
  saveAttempts: number;
  metrics?: { durationMs: number; formLoads: number; readRequests?: number };
  diagnostics?: string[];
  tasks: { code: string; before: string; temporary: string; after?: string; finalized?: boolean; applied: boolean; restored: boolean; error: string | null }[];
}

// Self-contained page-world test runner. No extension APIs or persistent page bridge.
export async function runNoWriteExperimentFromPage(readOnly = false, request?: NoWriteRequest, inspectState = false, directProbe = false, selectedIds?: string[]): Promise<NoExperimentResult> {
  const startedAt = Date.now();
  const guard = () => {
    const url = new URL(window.location.href);
    if (url.origin !== 'https://office.meygroup.vn' || url.pathname !== '/apps/work-project-project/view' || url.searchParams.get('ID') !== '870') {
      throw new Error('Thử nghiệm ghi chỉ được phép trong dự án test, ID 870.');
    }
  };
  guard();
  if (inspectState && !readOnly) throw new Error('Đối soát nhật ký chỉ được đọc, không được ghi.');
  if (directProbe && (!readOnly || request || inspectState)) throw new Error('Kiểm tra mở trực tiếp chỉ được đọc, không được ghi.');
  const direct = directProbe || request?.directEdit;
  const requireVisible = () => {
    if (document.visibilityState !== 'visible') throw new Error('Tab 1Office đang ở nền hoặc bị thu nhỏ. Giữ tab dự án test hiển thị trong lúc kiểm thử; không tự chạy lại nếu đã có thao tác lưu.');
  };
  requireVisible();
  const lock = 'data-tera-no-experiment';
  if (document.documentElement.hasAttribute(lock)) throw new Error('Đang có một thử nghiệm ghi chạy trong tab này.');
  document.documentElement.setAttribute(lock, 'running');
  const result: NoExperimentResult = { mode: request ? 'plan' : 'experiment', readOnly, complete: false, error: null, phase: 'Chuẩn bị', saveAttempts: 0, diagnostics: [], tasks: [], metrics: { durationMs: 0, formLoads: 0, readRequests: 0 } };
  const allowedTargets = [
    { id: '100119', code: 'TEST01.3', temporary: '99001.1', editHref: 'https://office.meygroup.vn/apps/work-task-task/edit?ref=detail&_quickview=1&ID=100119' },
    { id: '100122', code: 'TEST01.INSERT', temporary: '99001.2', editHref: 'https://office.meygroup.vn/apps/work-task-task/edit?ref=detail&_quickview=1&ID=100122' },
    { id: '100113', code: 'TEST01.01', temporary: '99001.3', editHref: 'https://office.meygroup.vn/apps/work-task-task/edit?ref=detail&_quickview=1&ID=100113' },
    { id: '100116', code: 'TEST01.2', temporary: '99001.4', editHref: 'https://office.meygroup.vn/apps/work-task-task/edit?ref=detail&_quickview=1&ID=100116' },
    { id: '100143', code: 'TEST01.4', temporary: '99001.5', editHref: 'https://office.meygroup.vn/apps/work-task-task/edit?ref=detail&_quickview=1&ID=100143' },
  ];
  const targets = request ? allowedTargets.filter((target) => Array.isArray(request.changes) && request.changes.some((row) => row?.internalId === target.id))
    : selectedIds ? allowedTargets.filter((target) => selectedIds.includes(target.id)) : directProbe ? allowedTargets : allowedTargets.slice(0, 2);
  const apiParents = new Map<string, string>();
  const parentFor = (target: typeof targets[number]) => apiParents.get(target.id) ?? '100104';
  const pause = () => new Promise<void>((resolve) => setTimeout(resolve, 100));
  const visible = (element: Element) => element.getClientRects().length > 0 && getComputedStyle(element).visibility !== 'hidden';
  const activatePopup = async (element: HTMLElement) => {
    guard();
    requireVisible();
    if (!element.isConnected || !visible(element)) throw new Error('Không thao tác được nút popup đang ẩn hoặc đã bị thay thế.');
    const rect = element.getBoundingClientRect();
    const coordinates = {
      bubbles: true, cancelable: true, view: window,
      clientX: rect.left + rect.width / 2, clientY: rect.top + rect.height / 2,
      button: 0, buttons: 0,
    };
    // The observed popup manager resets its per-gesture WeakMap on pointerdown.
    element.dispatchEvent(new PointerEvent('pointerdown', { ...coordinates, buttons: 1, pointerId: 1, pointerType: 'mouse', isPrimary: true }));
    await new Promise<void>((resolve) => setTimeout(resolve, 25));
    guard();
    requireVisible();
    if (!element.isConnected || !visible(element)) return;
    // No pointerup: the page would derive another realclick from it.
    const event = new MouseEvent('realclick', coordinates);
    Object.defineProperty(event, 'which', { value: 1 });
    element.dispatchEvent(event);
  };
  const activateLink = async (element: HTMLAnchorElement, collapsedMenu = false) => {
    guard();
    requireVisible();
    const inCollapsedMenu = collapsedMenu && element.title === 'Sửa' && !!element.closest('.list-menu.hidden');
    if (!element.isConnected || (!visible(element) && !inCollapsedMenu)) throw new Error('Không thao tác được liên kết đang ẩn hoặc đã bị thay thế.');
    const previous = element.onclick;
    // MAIN-world addEventListener is wrapped asynchronously by 1Office. A native
    // event-handler property cancels document navigation during dispatch itself.
    const preventNavigation: typeof element.onclick = function (event) {
      event.preventDefault();
      previous?.call(this, event);
      return false;
    };
    element.onclick = preventNavigation;
    try {
      // The inspected core-app handler opens internal links on click.
      element.click();
    } finally {
      if (element.onclick === preventNavigation) element.onclick = previous;
    }
  };
  const wait = async <T>(phase: string, read: () => T | null | false): Promise<T> => {
    result.phase = phase;
    const deadline = Date.now() + 30000;
    while (Date.now() < deadline) {
      guard();
      requireVisible();
      const value = read();
      if (value) return value;
      await pause();
    }
    throw new Error(`Hết thời hạn 30 giây tại bước: ${phase}. Không tự gửi lại yêu cầu ghi; cần kiểm tra trạng thái hiện tại.`);
  };
  const forms = () => Array.from(document.querySelectorAll<HTMLFormElement>('form')).filter(visible);
  const read = (form: HTMLFormElement, name: string) => String(new FormData(form).get(name) ?? '');
  const fingerprint = (form: HTMLFormElement) => JSON.stringify(
    Array.from(new FormData(form).entries())
      .filter(([name, value]) => name && name !== 'cf19' && !/csrf|token/i.test(name) && typeof value === 'string')
      .sort(([a, av], [b, bv]) => a.localeCompare(b) || String(av).localeCompare(String(bv))),
  );
  const check = (form: HTMLFormElement, target: typeof targets[number]) => {
    guard();
    const action = new URL(form.action, window.location.href);
    if (action.origin !== window.location.origin || action.pathname !== '/apps/work-task-task/edit' || form.method.toLowerCase() !== 'post'
      || read(form, 'ID') !== target.id || read(form, 'code') !== target.code || read(form, 'project_id') !== '870' || read(form, 'parent_id') !== parentFor(target)) {
      throw new Error('Biểu mẫu không khớp ID, mã, dự án hoặc công việc cha đã được cho phép. Dừng ghi.');
    }
  };
  const closeForm = async (form: HTMLFormElement) => {
    const dialog = form.closest('.dialog');
    const close = dialog && Array.from(dialog.querySelectorAll<HTMLElement>('.dialog-close[pop-close]'))
      .find((button) => visible(button) && button.closest('.dialog') === dialog);
    if (!close) throw new Error('Không tìm thấy nút đóng biểu mẫu thử nghiệm.');
    await activatePopup(close);
    await wait('Đóng biểu mẫu vừa đọc', () => !visible(form));
  };
  const closeDetails = async () => {
    for (const dialog of Array.from(document.querySelectorAll<HTMLElement>('.task-quick-view')).filter(visible)) {
      const close = Array.from(dialog.querySelectorAll<HTMLElement>('.dialog-close[pop-close]'))
        .find((button) => visible(button) && button.closest('.dialog') === dialog);
      if (!close) throw new Error('Không tìm thấy nút đóng chi tiết công việc.');
      await activatePopup(close);
      await wait('Đóng hộp chi tiết cũ', () => !visible(dialog));
    }
  };
  const open = async (target: typeof targets[number]): Promise<HTMLFormElement> => {
    if (forms().some((form) => form.querySelector('[name="cf19"]'))) throw new Error('Đang mở biểu mẫu chưa đóng. Dừng để tránh ảnh hưởng thao tác của bạn.');
    await closeDetails();
    const link = Array.from(document.querySelectorAll<HTMLAnchorElement>('a[href]')).find((a) => {
      const url = new URL(a.href, window.location.href);
      return visible(a) && url.origin === window.location.origin && url.pathname === '/apps/work-task-task/quickview' && url.searchParams.get('ID') === target.id;
    });
    if (!link) throw new Error(`Không tìm thấy công việc ${target.code} trong dự án test.`);
    let detail: HTMLElement | null = null;
    if (direct) {
      if ((!directProbe && !request?.silentReads) || !apiOriginals.has(target.id)) throw new Error('Chưa có dữ liệu API xác minh để mở trực tiếp biểu mẫu.');
      result.phase = `Mở trực tiếp biểu mẫu ${target.code}`;
    } else {
      result.phase = `Mở chi tiết ${target.code}`;
      await activateLink(link);
      detail = await wait(`Mở chi tiết ${target.code}`, () => Array.from(document.querySelectorAll<HTMLElement>('.task-quick-view')).find((dialog) => {
      if (!visible(dialog)) return false;
      const code = Array.from(dialog.querySelectorAll('.detail-field')).find((field) =>
        field.querySelector('.detail-field-label')?.textContent?.trim() === 'Mã công việc')
        ?.querySelector('.detail-field-content')?.textContent?.trim();
      if (code !== target.code) return false;
      // Header actions vary by task status/permissions. Do not depend on member.
      return Array.from(dialog.querySelectorAll<HTMLAnchorElement>('.detail-header-btns a[href]')).some((anchor) => {
        const url = new URL(anchor.href, window.location.href);
        return url.origin === window.location.origin
          && ['/apps/work-task-task/member', '/apps/work-task-task/econtract', '/apps/work-task-task/progress'].includes(url.pathname)
          && url.searchParams.get('ID') === target.id;
      });
      }) ?? null);
    }
    const diagnosticIndex = result.diagnostics!.length;
    const candidates = Array.from(document.querySelectorAll<HTMLAnchorElement>('.list-menu a[href]')).filter((a) => {
      const url = new URL(a.href, window.location.href);
      return url.pathname === '/apps/work-task-task/edit' && url.searchParams.get('ID') === target.id;
    });
    const links = candidates.filter((a) => {
        const url = new URL(a.href, window.location.href);
        return a.title === 'Sửa' && url.origin === window.location.origin
          && !a.hasAttribute('hidden') && !a.classList.contains('hidden') && !a.hasAttribute('disabled')
          && a.getAttribute('aria-disabled') !== 'true' && a.getAttribute('aria-hidden') !== 'true'
          && getComputedStyle(a).display !== 'none' && getComputedStyle(a).visibility !== 'hidden'
          && (visible(a) || !!a.closest('.list-menu.hidden'));
    });
    if (candidates.length && !links.length) {
      result.diagnostics![diagnosticIndex] = `${target.code}: Liên kết Sửa hiện có không hợp lệ hoặc bị vô hiệu hóa. Không dùng URL dự phòng.`;
      throw new Error('Liên kết Sửa hiện có không hợp lệ hoặc bị vô hiệu hóa. Dừng kiểm thử.');
    }
    let temporaryLink: HTMLAnchorElement | null = null;
    let edit = links[links.length - 1];
    if (!edit) {
      // Exact URLs observed live for these two test tasks only, not a general API template.
      temporaryLink = document.createElement('a');
      temporaryLink.href = target.editHref;
      temporaryLink.textContent = 'T-eraClip: mở biểu mẫu kiểm thử';
      temporaryLink.setAttribute('data-tera-no-form-link', '');
      // Keep the native app root of the existing task link, even without a detail dialog.
      (detail ?? link.parentElement!).append(temporaryLink);
      edit = temporaryLink;
    }
    result.diagnostics![diagnosticIndex] = `${target.code}: ${temporaryLink ? 'Dùng URL biểu mẫu test đã xác minh, qua liên kết tạm.' : 'Dùng liên kết Sửa sẵn có.'} Không mở dropdown. Đây là bước mở biểu mẫu, không phải kết quả lưu.`;
    if (direct) result.diagnostics![diagnosticIndex] += ' Không mở chi tiết công việc; kiểm tra định danh trên biểu mẫu trước khi lưu.';
    let form: HTMLFormElement;
    try {
      result.metrics!.formLoads++;
      await activateLink(edit, !temporaryLink);
      form = await wait(`Tải biểu mẫu ${target.code}`, () => forms().find((f) => f.querySelector('[name="cf19"]') && read(f, 'ID') === target.id) ?? null);
    } finally {
      temporaryLink?.remove();
    }
    // Selects are enriched asynchronously. Require a stable fully initialized form.
    let prior = '';
    let stableSince = Date.now();
    await wait(`Chờ các trường dự án/cha của ${target.code}`, () => {
      const state = fingerprint(form);
      if (state !== prior) { prior = state; stableSince = Date.now(); }
      return form.classList.contains('ajax-form-done') && read(form, 'project_id') === '870' && read(form, 'parent_id') === parentFor(target) && Date.now() - stableSince >= 600 ? true : null;
    });
    check(form, target);
    return form;
  };
  const save = async (form: HTMLFormElement, target: typeof targets[number], expected: string, next: string, original: string) => {
    requireVisible();
    check(form, target);
    if (read(form, 'cf19') !== expected || fingerprint(form) !== original) throw new Error('Dữ liệu đã thay đổi ngoài dự kiến. Không ghi đè.');
    const input = form.querySelector<HTMLInputElement>('input[name="cf19"]');
    const button = Array.from(form.querySelectorAll<HTMLButtonElement>('button')).find((b) => visible(b) && b.textContent?.trim().toLowerCase() === 'cập nhật');
    if (!input || !button || button.disabled) throw new Error('Không tìm thấy ô No. hoặc nút cập nhật sẵn sàng.');
    Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!.call(input, next);
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.dispatchEvent(new Event('change', { bubbles: true }));
    if (read(form, 'cf19') !== next || fingerprint(form) !== original) throw new Error('Không xác nhận được thay đổi riêng No. Chưa gửi yêu cầu lưu.');
    uncertainWrite = true;
    result.saveAttempts++;
    const previous = form.onsubmit;
    const preventDocumentSubmit: typeof form.onsubmit = function (event) {
      event.preventDefault();
      previous?.call(this, event);
      return false;
    };
    form.onsubmit = preventDocumentSubmit;
    try {
      button.click();
      await wait(`Chờ lưu No. của ${target.code}`, () => !visible(form));
    } finally {
      if (form.onsubmit === preventDocumentSubmit) form.onsubmit = previous;
    }
  };
  const originals = new Map<string, { no: string; fingerprint: string }>();
  const apiOriginals = new Map<string, string>();
  // Read-only Gantt endpoints already used by the inventory collector. No custom write request.
  const apiJson = async (url: URL, init?: RequestInit): Promise<any> => {
    guard();
    if (url.origin !== window.location.origin || url.pathname !== '/apps/work-project-project/gantt' || url.searchParams.get('project_id') !== '870') throw new Error('Địa chỉ đọc ngầm ngoài phạm vi test.');
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 30000);
    result.metrics!.readRequests!++;
    try {
      const response = await fetch(url.href, { ...init, credentials: 'same-origin', cache: 'no-store', redirect: 'error', signal: controller.signal,
        headers: { 'X-Requested-With': 'XMLHttpRequest', ...(init?.headers ?? {}) } });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const data = await response.json();
      guard();
      return data;
    } catch {
      throw new Error(controller.signal.aborted ? 'Đọc ngầm: máy chủ không phản hồi trong 30 giây. Dừng, không gửi lại yêu cầu ghi.' : 'Đọc ngầm thất bại hoặc phiên bị chuyển hướng. Dừng, không tự gửi lại yêu cầu ghi.');
    } finally { clearTimeout(timer); }
  };
  const canonical = (value: any): any => Array.isArray(value) ? value.map(canonical)
    : value && typeof value === 'object' ? Object.fromEntries(Object.keys(value).sort().map((key) => [key, canonical(value[key])])) : value;
  const readApi = async (ids: string[]) => {
    result.phase = 'Đọc ngầm dữ liệu mới từ 1Office';
    const structureUrl = new URL('/apps/work-project-project/gantt?project_id=870&structure=1&group_type=phase', window.location.origin);
    const structure = await apiJson(structureUrl);
    if (!Array.isArray(structure?.tasks)) throw new Error('Đọc ngầm thiếu cấu trúc công việc.');
    const tasks = new Map<string, any>();
    for (const task of structure.tasks) {
      const id = String(task?.ID ?? '');
      if (!/^\d+$/.test(id) || tasks.has(id)) throw new Error('Đọc ngầm có ID thiếu hoặc trùng.');
      tasks.set(id, task);
    }
    const states = new Map<string, { no: string; fingerprint: string }>();
    for (let offset = 0; offset < ids.length; offset += 50) {
      const batch = ids.slice(offset, offset + 50);
      const endpoint = new URL('/apps/work-project-project/gantt?project_id=870', window.location.origin);
      const data = await apiJson(endpoint, { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8' },
        body: new URLSearchParams({ mode: 'enrich', project_id: '870', ids: JSON.stringify(batch) }).toString() });
      for (const id of batch) {
        const task = tasks.get(id);
        const fields = data?.enriched?.[id]?.fields;
        if (!task || !fields || typeof fields !== 'object' || Array.isArray(fields)
          || !Object.prototype.hasOwnProperty.call(fields, 'cf19') || !['string', 'number'].includes(typeof fields.cf19)) throw new Error(`Đọc ngầm thiếu No. gốc của ID ${id}.`);
        const target = allowedTargets.find((row) => row.id === id);
        if (target) {
          const parent = task.parent_id === null || String(task.parent_id) === '0' ? '' : String(task.parent_id ?? '');
          if (task.code !== target.code || typeof task.title !== 'string' || !Object.prototype.hasOwnProperty.call(task, 'parent_id') || (parent !== '' && !/^\d+$/.test(parent))) throw new Error('Đọc ngầm không khớp mã, tên hoặc công việc cha của task test.');
          if (apiParents.has(id) && apiParents.get(id) !== parent) throw new Error('Quan hệ cha/con đã thay đổi trong lúc ghi. Dừng kiểm tra.');
          apiParents.set(id, parent);
        }
        // Protect all returned raw fields, not HTML display snippets or metadata.
        const enriched = Object.fromEntries(Object.entries(fields).filter(([key]) => key !== 'cf19'));
        const rawFields = task.fields && typeof task.fields === 'object' && !Array.isArray(task.fields)
          ? Object.fromEntries(Object.entries(task.fields).filter(([key]) => key !== 'cf19')) : {};
        const identity = Object.fromEntries(['ID', 'code', 'title', 'parent_id', 'phase_id', 'start_plan', 'end_plan', 'start', 'end', 'status'].map((key) => [key, task[key] ?? null]));
        states.set(id, { no: String(fields.cf19).trim(), fingerprint: JSON.stringify(canonical({ identity, rawFields, enriched })) });
      }
    }
    return { states, taskIds: [...tasks.keys()] };
  };
  let activeForm: HTMLFormElement | null = null;
  let uncertainWrite = false;
  try {
    if (selectedIds && (!readOnly || request || !selectedIds.length || selectedIds.length > 5 || new Set(selectedIds).size !== selectedIds.length || targets.length !== selectedIds.length)) throw new Error('Phạm vi đọc phục hồi không hợp lệ.');
    if (directProbe) {
      const snapshot = await readApi(targets.map((target) => target.id));
      for (const [id, state] of snapshot.states) apiOriginals.set(id, state.fingerprint);
    }
    if (request) {
      if (readOnly || !Array.isArray(request.changes) || !request.changes.length || request.changes.length > 5
        || new Set(request.changes.map((row) => row.internalId)).size !== request.changes.length
        || request.changes.some((row) => !allowedTargets.some((target) => target.id === row.internalId && target.code === row.code)
          || (row.before !== '' && !/^[1-9]\d*(?:\.[1-9]\d*)*$/.test(row.before)) || !/^[1-9]\d*(?:\.[1-9]\d*)*$/.test(row.after) || row.before === row.after)
        || new Set(request.changes.map((row) => row.after)).size !== request.changes.length) {
        throw new Error('Danh sách ghi không hợp lệ hoặc vượt phạm vi năm mã công việc test cho phép.');
      }
      if (!Array.isArray(request.sourceTasks) || !request.sourceTasks.length
        || request.sourceTasks.some((row) => !/^\d+$/.test(row.internalId) || typeof row.no !== 'string')
        || new Set(request.sourceTasks.map((row) => row.internalId)).size !== request.sourceTasks.length) {
        throw new Error('Bản nguồn thiếu ID hoặc No. gốc đã xác minh. Kiểm tra lại danh sách dự án.');
      }
      const live = new Map<string, string>();
      if (request.silentReads) {
        const snapshot = await readApi(request.sourceTasks.map((row) => row.internalId));
        if (snapshot.taskIds.length !== request.sourceTasks.length || snapshot.taskIds.some((id) => !request.sourceTasks.some((row) => row.internalId === id))) throw new Error('Danh sách máy chủ đã khác bản xem trước. Chưa ghi dữ liệu.');
        for (const [id, state] of snapshot.states) { live.set(id, state.no); apiOriginals.set(id, state.fingerprint); }
      } else for (const element of Array.from(document.querySelectorAll('.gantt-body-row[data-side="left"][item-type="item"]'))) {
        const id = element.getAttribute('item-id')?.match(/^item_(\d+)$/)?.[1];
        const cell = element.querySelector('.gantt-cf19');
        if (!id || !cell || live.has(id)) throw new Error('Không nhận diện được đầy đủ ID và ô No. gốc trong bảng dự án. Chưa ghi dữ liệu.');
        live.set(id, cell.textContent?.trim() || '');
      }
      if (live.size !== request.sourceTasks.length) {
        throw new Error(`Danh sách nguồn chưa khớp: trang có ${live.size}/${request.sourceTasks.length} công việc. Kiểm tra lại danh sách dự án, chưa ghi dữ liệu.`);
      }
      for (const row of request.sourceTasks) {
        if (!live.has(row.internalId) || live.get(row.internalId) !== row.no) throw new Error(`No. gốc của công việc ID ${row.internalId} khác bản xem trước. Kiểm tra lại danh sách dự án, chưa ghi dữ liệu.`);
      }
      const remaining = request.sourceTasks.filter((row) => !request.changes.some((change) => change.internalId === row.internalId)).map((row) => row.no);
      for (const row of request.changes) {
        if (live.get(row.internalId) !== row.before) throw new Error('No. nguồn không khớp bản xem trước.');
      }
      const finalNumbers = [...remaining.filter(Boolean), ...request.changes.map((row) => row.after)];
      if (new Set(finalNumbers).size !== finalNumbers.length || request.changes.some((row) => allowedTargets.some((target) => target.temporary === row.after))) {
        throw new Error('No. mới bị trùng hoặc dùng số tạm dành riêng cho kiểm thử.');
      }
    }
    if (!inspectState && forms().some((f) => f.querySelector('[name="cf19"]'))) throw new Error('Hãy đóng biểu mẫu đang sửa trước khi chạy thử nghiệm.');
    const existingNos = Array.from(document.querySelectorAll('.gantt-cf19')).map((e) => e.textContent?.trim());
    if (!inspectState && targets.some((t) => existingNos.includes(t.temporary))) throw new Error('No. tạm đã có trong danh sách, không thực hiện thử nghiệm.');
    if (inspectState) {
      const snapshot = await readApi(targets.map((target) => target.id));
      for (const target of targets) result.tasks.push({ code: target.code, before: snapshot.states.get(target.id)!.no,
        temporary: target.temporary, applied: false, restored: false, error: null });
    }
    // Read both before the first write, then reopen each immediately before submitting.
    for (const target of (inspectState ? [] : targets)) {
      if (request?.silentReads) {
        const planned = request.changes.find((row) => row.internalId === target.id)!;
        originals.set(target.id, { no: planned.before, fingerprint: '' });
        result.tasks.push({ code: target.code, before: planned.before, temporary: target.temporary, after: planned.after, finalized: false, applied: false, restored: false, error: null });
        continue;
      }
      activeForm = await open(target);
      const no = read(activeForm, 'cf19');
      if (!inspectState && !directProbe && !/^[1-9]\d*(?:\.[1-9]\d*)*$/.test(no)) throw new Error('No. ban đầu không hợp lệ.');
      const planned = request?.changes.find((row) => row.internalId === target.id);
      if (planned && no !== planned.before) throw new Error(`${target.code}: No. hiện tại đã khác bản xem trước. Chưa ghi dữ liệu.`);
      originals.set(target.id, { no, fingerprint: fingerprint(activeForm) });
      result.tasks.push({ code: target.code, before: no, temporary: target.temporary, after: planned?.after, finalized: false, applied: false, restored: false, error: null });
      await closeForm(activeForm); activeForm = null;
    }
    for (const [index, target] of (readOnly || request ? [] : targets).entries()) {
      const original = originals.get(target.id)!;
      activeForm = await open(target);
      await save(activeForm, target, original.no, target.temporary, original.fingerprint);
      activeForm = null;
      activeForm = await open(target);
      if (read(activeForm, 'cf19') !== target.temporary || fingerprint(activeForm) !== original.fingerprint) throw new Error('Đọc lại không khớp No. mới hoặc trường khác thay đổi. Cần kiểm tra thủ công.');
      result.tasks[index].applied = true;
      uncertainWrite = false;
      await closeForm(activeForm); activeForm = null;
    }
    if (request) {
      // Free destinations first. Break only actual cycles with a temporary No.
      // The schedule is built from the complete, ID-matched raw snapshot.
      const occupied = new Map(request.sourceTasks.filter((row) => row.no).map((row) => [row.no, row.internalId]));
      const current = new Map(request.changes.map((row) => [row.internalId, row.before]));
      const pending = [...targets];
      const steps: { target: typeof targets[number]; from: string; to: string }[] = [];
      while (pending.length) {
        const free = pending.find((target) => !occupied.has(request.changes.find((row) => row.internalId === target.id)!.after));
        const target = free ?? pending[0];
        const from = current.get(target.id)!;
        const to = free ? request.changes.find((row) => row.internalId === target.id)!.after : target.temporary;
        if (occupied.has(to)) throw new Error('No. tạm bị trùng. Dừng trước khi ghi.');
        steps.push({ target, from, to });
        if (from) occupied.delete(from);
        occupied.set(to, target.id); current.set(target.id, to);
        if (free) pending.splice(pending.indexOf(target), 1);
      }
      for (const { target, from, to } of steps) {
        const index = targets.indexOf(target);
        const original = originals.get(target.id)!;
        const row = result.tasks[index];
        activeForm = await open(target);
        if (request.silentReads && !original.fingerprint) original.fingerprint = fingerprint(activeForm);
        await save(activeForm, target, from, to, original.fingerprint);
        activeForm = null;
        if (request.silentReads) {
          const state = (await readApi([target.id])).states.get(target.id)!;
          if (state.no !== to || state.fingerprint !== apiOriginals.get(target.id)) throw new Error('Đọc ngầm sau lưu không khớp No. hoặc trường dữ liệu API đã thay đổi. Dừng, kiểm tra thủ công.');
        } else {
          activeForm = await open(target);
          if (read(activeForm, 'cf19') !== to || fingerprint(activeForm) !== original.fingerprint) throw new Error('Chưa xác nhận được No. vừa ghi hoặc trường khác đã thay đổi. Kiểm tra thủ công.');
        }
        row.applied = true;
        row.finalized = to === row.after;
        uncertainWrite = false;
        if (activeForm) { await closeForm(activeForm); activeForm = null; }
      }
    }
  } catch (error) {
    result.error = error instanceof Error ? error.message : 'Thử nghiệm không hoàn tất.';
  } finally {
    // Restore only confirmed writes whose current form still exactly matches our change.
    if (request) {
      // A plan is not a transaction. Never replay or roll back a partially applied plan automatically.
      if (result.error && result.saveAttempts > 0) result.error += ' Có thể đã cập nhật một phần hoặc còn No. tạm. Không tự chạy lại; kiểm tra từng công việc trước.';
      if (!uncertainWrite && activeForm && visible(activeForm)) {
        try { await closeForm(activeForm); activeForm = null; } catch { /* Keep the original failure. */ }
      }
    } else if (!uncertainWrite) {
      if (activeForm && visible(activeForm)) { try { await closeForm(activeForm); activeForm = null; } catch { /* Report restoration failures below. */ } }
      for (const [index, target] of targets.entries()) {
        const row = result.tasks[index];
        if (!row?.applied) continue;
        try {
          const original = originals.get(target.id)!;
          activeForm = await open(target);
          await save(activeForm, target, target.temporary, original.no, original.fingerprint);
          activeForm = null;
          activeForm = await open(target);
          if (read(activeForm, 'cf19') !== original.no || fingerprint(activeForm) !== original.fingerprint) throw new Error('Chưa xác nhận được hoàn nguyên.');
          row.restored = true;
          uncertainWrite = false;
          await closeForm(activeForm); activeForm = null;
        } catch (error) {
          row.error = error instanceof Error ? error.message : 'Hoàn nguyên thất bại.';
          if (uncertainWrite) break;
        }
      }
    } else {
      result.error = `${result.error} Có yêu cầu ghi chưa xác định kết quả. Không tự gửi lại hoặc hoàn nguyên; kiểm tra các công việc test trong lượt ghi trước.`;
    }
    result.complete = result.tasks.length === targets.length && !result.error && (request ? result.tasks.every((row) => row.finalized) : readOnly || result.tasks.every((row) => row.applied && row.restored));
    if (result.complete) result.phase = request ? request.silentReads ? 'Đã đọc ngầm và xác nhận toàn bộ No. mới trong danh sách' : 'Đã đọc lại và xác nhận toàn bộ No. mới trong danh sách' : inspectState ? `Đã đọc ngầm trạng thái ${targets.length} công việc, không ghi` : readOnly ? `Đã đọc ${targets.length} biểu mẫu, không kích hoạt lưu` : 'Đã xác nhận ghi và hoàn nguyên cả hai công việc';
    result.metrics!.durationMs = Date.now() - startedAt;
    document.documentElement.removeAttribute(lock);
  }
  return result;
}

async function executeNoRunner(tabId: number, readOnly = false, request?: NoWriteRequest, inspectState = false, directProbe = false, selectedIds?: string[]): Promise<NoExperimentResult> {
  const api = (globalThis as any).chrome;
  if (!Number.isSafeInteger(tabId) || tabId <= 0 || !api?.scripting?.executeScript || !api.tabs?.get || !api.tabs?.update) throw new Error('Mở lại đối chiếu từ popup của dự án test để cấp quyền thao tác tab nguồn.');
  const tab = await api.tabs.get(tabId);
  let url: URL;
  try { url = new URL(tab.url); } catch { throw new Error('Không đọc được địa chỉ tab nguồn. Mở lại đối chiếu từ popup của dự án test.'); }
  if (url.origin !== 'https://office.meygroup.vn' || url.pathname !== '/apps/work-project-project/view' || url.searchParams.get('ID') !== '870') {
    throw new Error('Thử nghiệm ghi chỉ được phép trong dự án test, ID 870.');
  }
  // 1Office's renderer waits for animation frames. Keep the source visible;
  // do not change Chrome throttling settings or repeatedly steal focus.
  const active = await api.tabs.update(tabId, { active: true });
  if (!active?.active) throw new Error('Không thể đưa tab dự án test lên trước. Chưa chạy kiểm thử.');
  // Page-world event properties stay in the same realm as 1Office's popup handlers.
  // Do not expose an arbitrary-command listener or extension APIs to the page.
  const results = await api.scripting.executeScript({ target: { tabId }, world: 'MAIN', func: runNoWriteExperimentFromPage, args: selectedIds ? [true, null, true, false, selectedIds] : directProbe ? [true, null, false, true] : inspectState ? [true, null, true] : request ? [readOnly, request] : [readOnly] });
  if (!results?.[0]?.result) throw new Error('Không nhận được kết quả. Không chạy lại trước khi kiểm tra No. các công việc test trong lượt ghi.');
  return results[0].result as NoExperimentResult;
}

export async function probeDirectNoForms(tabId: number) {
  return withNoWriteLock(() => executeNoRunner(tabId, true, undefined, false, true));
}

export async function runNoWriteExperiment(tabId: number, readOnly = false): Promise<NoExperimentResult> {
  if (readOnly) return executeNoRunner(tabId, true);
  return withNoWriteLock(async () => {
    if ((await readNoWriteJournal())?.status === 'pending') throw new Error('Có lượt ghi chưa đối soát. Chưa cho phép thử ghi hoặc hoàn nguyên.');
    return executeNoRunner(tabId, false);
  });
}

export async function inspectNoWriteJournal(tabId: number) {
  return withNoWriteLock(async () => {
    const journal = await readNoWriteJournal();
    if (!journal) throw new Error('Không có nhật ký để đối soát.');
    const result = await executeNoRunner(tabId, true, undefined, true, false, journal.rows.map((row) => row.internalId));
    return reconcileNoWriteJournal(journal, result);
  });
}

export async function runNoUpdatePlan(tabId: number, inventory: OneOfficeProjectInventory, text: string): Promise<NoExperimentResult> {
  const plan = buildNoUpdatePlan(inventory, text);
  if (plan.projectId !== '870' || plan.errors.length || !plan.changes.length) throw new Error('Cần bản xem trước hợp lệ, có thay đổi và thuộc dự án test ID 870.');
  if (plan.changes.length > 5 || plan.changes.some((row) => !isNoTestTask(row))) {
    throw new Error('Chỉ cho phép ghi năm mã công việc đã định danh trong dự án test.');
  }
  if (inventory.tasks.some((task) => typeof task.sourceNo !== 'string')) throw new Error('Bản nguồn cũ chưa có No. gốc đã xác minh. Kiểm tra lại danh sách dự án và mở đối chiếu mới.');
  if (oneOfficeOrigin(inventory.project.url) !== ONEOFFICE_ALLOWED_ORIGIN) {
    throw new Error('Chưa bật ghi No. cho môi trường này. Không dùng bản nguồn test để ghi vào site chính.');
  }
  const request: NoWriteRequest = {
    silentReads: true,
    directEdit: true,
    changes: plan.changes.map(({ internalId, code, before, after }) => ({ internalId, code, before, after })),
    sourceTasks: inventory.tasks.map((task) => ({ internalId: task.internalId, no: task.sourceNo! })),
  };
  return withNoWriteLock(async () => {
    const journal = await prepareNoWriteJournal(request);
    const result = await executeNoRunner(tabId, false, request);
    await finishNoWriteJournal(journal, result);
    return result;
  });
}
