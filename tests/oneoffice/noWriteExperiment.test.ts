import { afterEach, describe, expect, it, vi } from 'vitest';
import { runNoUpdatePlan, runNoWriteExperiment, runNoWriteExperimentFromPage, type NoWriteRequest } from '../../src/oneoffice/noWriteExperiment';
import type { OneOfficeProjectInventory } from '../../src/oneoffice/types';

let stopRecording = () => {};
afterEach(() => { stopRecording(); vi.restoreAllMocks(); vi.unstubAllGlobals(); vi.useRealTimers(); document.body.innerHTML = ''; document.documentElement.removeAttribute('data-tera-no-experiment'); });

function setup(allFive = false) {
  vi.useFakeTimers();
  const visibility = vi.spyOn(document, 'visibilityState', 'get').mockReturnValue('visible');
  const NativeMouseEvent = MouseEvent;
  class TestMouseEvent extends NativeMouseEvent {
    constructor(type: string, init: MouseEventInit) { super(type, { ...init, view: null }); }
  }
  vi.stubGlobal('MouseEvent', TestMouseEvent);
  vi.stubGlobal('PointerEvent', TestMouseEvent);
  vi.stubGlobal('window', { location: { href: 'https://office.meygroup.vn/apps/work-project-project/view?ID=870', origin: 'https://office.meygroup.vn' } });
  vi.spyOn(Element.prototype, 'getClientRects').mockImplementation(function (this: Element) {
    return { length: this.isConnected && !(this as HTMLElement).closest('.hidden') ? 1 : 0 } as DOMRectList;
  });
  const data = new Map([['100119', { code: 'TEST01.3', no: '1.3' }], ['100122', { code: 'TEST01.INSERT', no: '1.4' }]]);
  if (allFive) {
    data.set('100113', { code: 'TEST01.01', no: '1.1' });
    data.set('100116', { code: 'TEST01.2', no: '1.2' });
    data.set('100143', { code: 'TEST01.4', no: '' });
  }
  const parents = new Map([...data.keys()].map((id) => [id, id === '100143' ? '100116' : '100104']));
  const formOpeners = new Map<string, () => void>();
  const writes: string[] = [];
  const formLoads: string[] = [];
  const savedNumbers: string[][] = [];
  const anchorClicks: Event[] = [];
  const submitEvents: Event[] = [];
  const realClicks: string[] = [];
  const menuPointerEvents: string[] = [];
  const closeEvents: string[] = [];
  let failSecond = false;
  let failFinal = false;
  let stallMenu = false;
  let hideEdit = false;
  let otherTitle = false;
  let disabledEdit = false;
  let foreignEdit = false;
  let wrongSecondId = false;
  let wrongSecondCode = false;
  let denySecondForm = false;
  let hideOnSecondWrite = false;
  let navigateOnClick = false;
  document.body.innerHTML = [...data].map(([id, task]) => `<div class="gantt-body-row" data-side="left" item-type="item" item-id="item_${id}"><div class="gantt-cf19">${task.no}</div></div>`).join('');
  const recorder = (event: Event) => {
    const anchor = (event.target as Element).closest<HTMLAnchorElement>('a[href]');
    if (anchor) {
      anchorClicks.push(event);
      if (navigateOnClick && !event.defaultPrevented) window.location.href = anchor.href;
      if (anchor.hasAttribute('data-tera-no-form-link')) {
        formOpeners.get(new URL(anchor.href).searchParams.get('ID')!)?.();
      }
    }
  };
  document.body.addEventListener('click', recorder);
  let pressTarget: EventTarget | null = null;
  const pointerdown = (event: Event) => {
    // Model the async 1Office event wrapper, not a synchronous onclick-only fixture.
    setTimeout(() => { pressTarget = event.target; }, 5);
  };
  const pointerup = (event: Event) => {
    if (pressTarget === event.target) {
      event.target!.dispatchEvent(new MouseEvent('realclick', { bubbles: true, cancelable: true, button: 0 }));
    }
    pressTarget = null;
  };
  document.addEventListener('pointerdown', pointerdown);
  document.addEventListener('pointerup', pointerup);
  stopRecording = () => {
    document.body.removeEventListener('click', recorder);
    document.removeEventListener('pointerdown', pointerdown);
    document.removeEventListener('pointerup', pointerup);
  };
  for (const [id, task] of data) {
    const openForm = () => {
      formLoads.push(id);
      if (denySecondForm && id === '100122') return false;
      const dialog = document.createElement('div'); dialog.className = 'dialog';
      dialog.innerHTML = `<i class="dialog-close" pop-close></i><form class="ajax-form-done" action="https://office.meygroup.vn/apps/work-task-task/edit" method="post"><input name="ID" value="${id}"><input name="project_id" value="870"><input name="parent_id" value="${parents.get(id)}"><input name="code" value="${task.code}"><input name="title" value="Original title"><input name="cf19" value="${task.no}"><button type="button">Cập nhật</button></form>`;
      dialog.querySelector<HTMLElement>('.dialog-close')!.addEventListener('realclick', (event) => {
        closeEvents.push('form');
        if (pressTarget === event.target && (event as MouseEvent).which === 1) dialog.remove();
      });
      dialog.querySelector('button')!.onclick = () => {
        const event = new Event('submit', { bubbles: true, cancelable: true });
        dialog.querySelector('form')!.dispatchEvent(event);
        submitEvents.push(event);
        writes.push(id);
        if (failFinal && writes.length === 3) return;
        if (hideOnSecondWrite && id === '100122') { visibility.mockReturnValue('hidden'); return; }
        if (failSecond && id === '100122') return;
        task.no = dialog.querySelector<HTMLInputElement>('[name="cf19"]')!.value;
        savedNumbers.push([...data.values()].map((row) => row.no));
        dialog.remove();
      };
      document.body.append(dialog);
      return false;
    };
    formOpeners.set(id, openForm);
    const link = document.createElement('a');
    link.href = `https://office.meygroup.vn/apps/work-task-task/quickview?ID=${id}`;
    link.textContent = task.code;
    link.onclick = () => {
      const detail = document.createElement('div'); detail.className = 'dialog task-quick-view';
      const detailId = wrongSecondId && id === '100122' ? '1001220' : id;
      const detailCode = wrongSecondCode && id === '100122' ? 'TEST01.3' : task.code;
      detail.innerHTML = `<div class="dialog hidden"><i class="dialog-close" pop-close></i></div><i class="dialog-close" pop-close></i><div class="detail-field"><div class="detail-field-label">Mã công việc</div><div class="detail-field-content">${detailCode}</div></div><div class="detail-header-btns"><a href="https://office.meygroup.vn/apps/work-task-task/${id === '100122' ? 'progress' : 'member'}?ID=${detailId}">Action</a><a class="btn-icon-circle pop-done">Menu</a></div>`;
      detail.querySelector<HTMLElement>(':scope > .dialog-close')!.addEventListener('realclick', (event) => {
        closeEvents.push('detail');
        if (pressTarget === event.target && (event as MouseEvent).which === 1) detail.remove();
      });
      const menu = detail.querySelector<HTMLElement>('.btn-icon-circle')!;
      for (const type of ['pointerdown', 'pointerup', 'click']) {
        menu.addEventListener(type, () => menuPointerEvents.push(type));
      }
      menu.addEventListener('realclick', () => {
        realClicks.push(menu.className);
      });
      {
        const popup = document.createElement('div'); popup.className = 'list-menu hidden';
        const edit = document.createElement('a'); edit.title = otherTitle ? 'Chỉnh sửa' : 'Sửa';
        if (hideEdit) edit.classList.add('hidden');
        if (disabledEdit) edit.setAttribute('aria-disabled', 'true');
        edit.href = `https://${foreignEdit ? 'example.com' : 'office.meygroup.vn'}/apps/work-task-task/edit?ID=${id}`;
        edit.onclick = openForm;
        popup.append(edit);
        // Cold second task: its dropdown has never been opened, so no edit link exists.
        if (!stallMenu && id === '100119') document.body.append(popup);
      }
      document.body.append(detail);
      return false;
    };
    document.body.append(link);
  }
  const enableApi = (fault?: 'missing' | 'wrong-id' | 'changed-field' | 'stale-no' | 'redirect' | 'timeout') => {
    const fetch = vi.fn(async (url: string, init?: RequestInit) => {
      if (fault === 'redirect') throw new Error('redirect');
      if (fault === 'timeout') return new Promise<Response>((_, reject) => init!.signal!.addEventListener('abort', () => reject(new Error('aborted'))));
      if (!init?.method) return { ok: true, json: async () => ({ tasks: [...data].map(([ID, task]) => ({ ID, code: task.code, parent_id: parents.get(ID), title: 'Original title', fields: { status: 'waiting' } })) }) };
      const ids = JSON.parse(new URLSearchParams(init.body as string).get('ids')!);
      return { ok: true, json: async () => ({ enriched: Object.fromEntries(ids.map((id: string) => [fault === 'wrong-id' ? '999' : id, { fields: {
        ...(fault === 'missing' ? {} : { cf19: fault === 'stale-no' && writes.length ? '1.3' : data.get(id)!.no }),
        status: fault === 'changed-field' && writes.length ? 'done' : 'waiting',
      } }])) }) };
    });
    vi.stubGlobal('fetch', fetch);
    return fetch;
  };
  return { data, writes, formLoads, savedNumbers, anchorClicks, submitEvents, realClicks, menuPointerEvents, closeEvents, enableApi, failFinal: () => { failFinal = true; }, failSecond: () => { failSecond = true; }, stallMenu: () => { stallMenu = true; }, hideEdit: () => { hideEdit = true; }, otherTitle: () => { otherTitle = true; }, disableEdit: () => { disabledEdit = true; }, foreignEdit: () => { foreignEdit = true; }, navigateOnClick: () => { navigateOnClick = true; }, wrongSecondId: () => { wrongSecondId = true; }, wrongSecondCode: () => { wrongSecondCode = true; }, denySecondForm: () => { denySecondForm = true; }, hide: () => visibility.mockReturnValue('hidden'), hideOnSecondWrite: () => { hideOnSecondWrite = true; } };
}

describe('test-only UI write runner', () => {
  const swapPlan = (): NoWriteRequest => ({
    sourceTasks: [{ internalId: '100119', no: '1.3' }, { internalId: '100122', no: '1.4' }],
    changes: [
      { internalId: '100119', code: 'TEST01.3', before: '1.3', after: '1.4' },
      { internalId: '100122', code: 'TEST01.INSERT', before: '1.4', after: '1.3' },
    ],
  });
  it('updates all five allowed tasks, including a raw blank and a different existing parent', async () => {
    const app = setup(true); app.enableApi();
    const sourceTasks = [...app.data].map(([internalId, row]) => ({ internalId, no: row.no }));
    const changes = [...app.data].map(([internalId, row], index) => ({ internalId, code: row.code, before: row.no, after: `2.${index + 1}` }));
    const pending = runNoWriteExperimentFromPage(false, { silentReads: true, directEdit: true, sourceTasks, changes });
    await vi.runAllTimersAsync();
    const result = await pending;
    expect(result.error).toBeNull(); expect(result.complete).toBe(true);
    expect(result.tasks).toHaveLength(5); expect(result.saveAttempts).toBe(5);
    expect(app.formLoads).toHaveLength(5);
    expect(result.tasks.every((row) => row.finalized)).toBe(true);
    expect(result.tasks.find((row) => row.code === 'TEST01.4')!.before).toBe('');
    expect([...app.data.values()].map((row) => row.no)).toEqual(['2.1', '2.2', '2.3', '2.4', '2.5']);
    expect(document.querySelector('.task-quick-view')).toBeNull();
  });
  it('breaks a four-task cycle with one temporary save while also numbering the blank fifth task', async () => {
    const app = setup(true); app.enableApi();
    const after = ['1.4', '1.1', '1.2', '1.3', '1.5'];
    const sourceTasks = [...app.data].map(([internalId, row]) => ({ internalId, no: row.no }));
    const changes = [...app.data].map(([internalId, row], index) => ({ internalId, code: row.code, before: row.no, after: after[index] }));
    const pending = runNoWriteExperimentFromPage(false, { silentReads: true, directEdit: true, sourceTasks, changes });
    await vi.runAllTimersAsync();
    const result = await pending;
    expect(result.complete).toBe(true); expect(result.saveAttempts).toBe(6);
    expect([...app.data.values()].map((row) => row.no)).toEqual(after);
    expect(app.savedNumbers.every((numbers) => new Set(numbers.filter(Boolean)).size === numbers.filter(Boolean).length)).toBe(true);
  });
  it('recovers only the selected journal tasks rather than assuming there are two entries', async () => {
    const app = setup(true); app.enableApi();
    const pending = runNoWriteExperimentFromPage(true, undefined, true, false, ['100113', '100116', '100143']);
    await vi.runAllTimersAsync();
    const result = await pending;
    expect(result.complete).toBe(true); expect(result.tasks).toHaveLength(3);
    expect(result.tasks.map((row) => row.code)).toEqual(['TEST01.01', 'TEST01.2', 'TEST01.4']);
    expect(app.writes).toEqual([]); expect(app.formLoads).toEqual([]);
  });
  it('uses background reads for preflight and verification, opening only three save forms for a swap', async () => {
    const app = setup(); const fetch = app.enableApi(); const plan = { ...swapPlan(), silentReads: true };
    const pending = runNoWriteExperimentFromPage(false, plan);
    await vi.runAllTimersAsync();
    const result = await pending;
    expect(result.complete).toBe(true);
    expect(result.saveAttempts).toBe(3);
    expect(app.formLoads).toHaveLength(3);
    expect(result.metrics?.readRequests).toBe(8);
    expect(app.data.get('100119')!.no).toBe('1.4');
    expect(fetch.mock.calls.every(([url, init]) => url.startsWith('https://office.meygroup.vn/apps/work-project-project/gantt?project_id=870') && init?.redirect === 'error' && init?.cache === 'no-store')).toBe(true);
  });
  it('opens verified edit URLs directly without activating task quickview links', async () => {
    const app = setup(); app.enableApi();
    const pending = runNoWriteExperimentFromPage(false, { ...swapPlan(), silentReads: true, directEdit: true });
    await vi.runAllTimersAsync();
    const result = await pending;
    expect(result.complete).toBe(true);
    expect(app.formLoads).toHaveLength(3);
    expect(app.anchorClicks).toHaveLength(3);
    expect(app.anchorClicks.every((event) => (event.target as HTMLAnchorElement).href.includes('/apps/work-task-task/edit?'))).toBe(true);
    expect(document.querySelector('.task-quick-view')).toBeNull();
    expect(document.querySelector('[data-tera-no-form-link]')).toBeNull();
    expect(app.closeEvents).toEqual([]);
    expect(app.data.get('100119')!.no).toBe('1.4');
    expect(app.data.get('100122')!.no).toBe('1.3');
  });
  it('probes both direct forms with zero writes before the new transport is used live', async () => {
    const app = setup(true); app.enableApi();
    const pending = runNoWriteExperimentFromPage(true, undefined, false, true);
    await vi.runAllTimersAsync();
    const result = await pending;
    expect(result.complete).toBe(true);
    expect(result.saveAttempts).toBe(0);
    expect(app.writes).toEqual([]);
    expect(app.formLoads).toEqual(['100119', '100122', '100113', '100116', '100143']);
    expect(app.anchorClicks).toHaveLength(5);
    expect(document.querySelector('.task-quick-view')).toBeNull();
    expect(document.querySelector('form')).toBeNull();
  });
  it('does not bypass a disabled edit action on the direct path', async () => {
    const app = setup(); app.enableApi();
    document.body.insertAdjacentHTML('beforeend', '<div class="list-menu hidden"><a title="Sửa" aria-disabled="true" href="https://office.meygroup.vn/apps/work-task-task/edit?ID=100119">Sửa</a></div>');
    const pending = runNoWriteExperimentFromPage(false, { ...swapPlan(), silentReads: true, directEdit: true });
    await vi.runAllTimersAsync();
    const result = await pending;
    expect(result.error).toContain('vô hiệu hóa'); expect(app.writes).toEqual([]);
    expect(app.anchorClicks).toEqual([]);
  });
  it('stops on a direct form timeout without reverting to quickview or resending a save', async () => {
    const app = setup(); app.enableApi(); app.denySecondForm();
    const pending = runNoWriteExperimentFromPage(false, { ...swapPlan(), silentReads: true, directEdit: true });
    await vi.runAllTimersAsync();
    const result = await pending;
    expect(result.complete).toBe(false); expect(result.error).toContain('30 giây');
    expect(app.writes).toEqual(['100119']);
    expect(app.formLoads).toEqual(['100119', '100122']);
    expect(document.querySelector('.task-quick-view')).toBeNull();
    expect(document.querySelector('[data-tera-no-form-link]')).toBeNull();
  });
  it('uses only one save form when the destination is free', async () => {
    const app = setup(); app.enableApi(); const plan = { ...swapPlan(), silentReads: true };
    plan.changes = [{ ...plan.changes[0], after: '1.5' }];
    const pending = runNoWriteExperimentFromPage(false, plan); await vi.runAllTimersAsync();
    expect((await pending).complete).toBe(true);
    expect(app.formLoads).toHaveLength(1); expect(app.writes).toHaveLength(1);
  });
  it.each(['missing', 'wrong-id', 'redirect', 'timeout'] as const)('refuses %s background data before any save', async (fault) => {
    const app = setup(); app.enableApi(fault);
    const pending = runNoWriteExperimentFromPage(false, { ...swapPlan(), silentReads: true });
    await vi.runAllTimersAsync();
    const result = await pending;
    expect(result.complete).toBe(false); expect(result.error).toBeTruthy();
    expect(app.writes).toEqual([]); expect(app.formLoads).toEqual([]);
    if (fault === 'timeout') expect(result.error).toContain('30 giây');
  });
  it.each(['changed-field', 'stale-no'] as const)('stops after one save if background verification reports %s', async (fault) => {
    const app = setup(); app.enableApi(fault);
    const pending = runNoWriteExperimentFromPage(false, { ...swapPlan(), silentReads: true });
    await vi.runAllTimersAsync();
    const result = await pending;
    expect(result.complete).toBe(false); expect(result.error).toContain('Đọc ngầm sau lưu');
    expect(app.writes).toEqual(['100119']); expect(result.tasks[0].finalized).toBe(false);
  });
  it('applies a previewed swap through temporary numbers and retains the verified final values', async () => {
    const app = setup();
    const pending = runNoWriteExperimentFromPage(false, swapPlan());
    await vi.runAllTimersAsync();
    const result = await pending;
    expect(result.complete).toBe(true);
    expect(result.mode).toBe('plan');
    expect(result.saveAttempts).toBe(3);
    expect(app.writes).toEqual(['100119', '100122', '100119']);
    expect(app.data.get('100119')!.no).toBe('1.4');
    expect(app.data.get('100122')!.no).toBe('1.3');
    expect(result.tasks.every((row) => row.finalized && !row.restored)).toBe(true);
    expect(app.formLoads).toHaveLength(8);
    expect(app.savedNumbers.every((numbers) => new Set(numbers).size === numbers.length)).toBe(true);
  });
  it.each(['stale-list', 'stale-form', 'unknown-id', 'duplicate-id', 'collision', 'temporary'] as const)('refuses a %s plan before saving', async (kind) => {
    const app = setup(); const plan = swapPlan();
    if (kind === 'stale-list') plan.sourceTasks.push({ internalId: '100123', no: '1.5' });
    if (kind === 'stale-form') app.data.get('100122')!.no = '1.5';
    if (kind === 'unknown-id') plan.changes[1].internalId = '100123';
    if (kind === 'duplicate-id') plan.changes[1] = { ...plan.changes[0] };
    if (kind === 'collision') { plan.changes.pop(); }
    if (kind === 'temporary') plan.changes[0].after = '99001.1';
    const pending = runNoWriteExperimentFromPage(false, plan);
    await vi.runAllTimersAsync();
    const result = await pending;
    expect(result.complete).toBe(false);
    expect(result.error).toBeTruthy();
    expect(app.writes).toEqual([]);
    expect(document.documentElement.hasAttribute('data-tera-no-experiment')).toBe(false);
  });
  it('stops a partially applied plan without retrying or automatically rolling back', async () => {
    const app = setup(); app.failSecond();
    const pending = runNoWriteExperimentFromPage(false, swapPlan());
    await vi.runAllTimersAsync();
    const result = await pending;
    expect(result.complete).toBe(false);
    expect(result.error).toContain('Không tự chạy lại');
    expect(app.writes).toEqual(['100119', '100122']);
    expect(app.data.get('100119')!.no).toBe('99001.1');
    expect(result.tasks.every((row) => !row.restored && !row.finalized)).toBe(true);
  });
  it('reports verified final and temporary values separately after an uncertain final save', async () => {
    const app = setup(); app.failFinal();
    const pending = runNoWriteExperimentFromPage(false, swapPlan());
    await vi.runAllTimersAsync();
    const result = await pending;
    expect(result.complete).toBe(false);
    expect(result.saveAttempts).toBe(3);
    expect(result.tasks[0].finalized).toBe(false);
    expect(result.tasks[1].finalized).toBe(true);
    expect(app.data.get('100119')!.no).toBe('99001.1');
    expect(app.data.get('100122')!.no).toBe('1.3');
    expect(app.writes).toHaveLength(3);
    expect(result.error).toContain('Không tự chạy lại');
  });
  it('updates a single allowed task without touching the other task', async () => {
    const app = setup(); const plan = swapPlan();
    plan.changes = [{ ...plan.changes[0], after: '1.5' }];
    const pending = runNoWriteExperimentFromPage(false, plan);
    await vi.runAllTimersAsync();
    expect((await pending).complete).toBe(true);
    expect(app.writes).toEqual(['100119']);
    expect(app.formLoads).toHaveLength(3);
    expect(app.data.get('100122')!.no).toBe('1.4');
  });
  it('orders an acyclic shift from the free end without temporary writes', async () => {
    const app = setup(); const plan = swapPlan(); plan.changes[1].after = '1.5';
    const pending = runNoWriteExperimentFromPage(false, plan);
    await vi.runAllTimersAsync();
    expect((await pending).complete).toBe(true);
    expect(app.writes).toEqual(['100122', '100119']);
    expect(app.formLoads).toHaveLength(6);
    expect(app.savedNumbers.every((numbers) => new Set(numbers).size === numbers.length)).toBe(true);
  });
  it('reads temporary or unexpected values during recovery without triggering saves', async () => {
    const app = setup(); app.data.get('100119')!.no = '99001.1'; app.data.get('100122')!.no = '';
    app.enableApi();
    document.querySelector('.gantt-cf19')!.textContent = '99001.1';
    const pending = runNoWriteExperimentFromPage(true, undefined, true);
    await vi.runAllTimersAsync();
    const result = await pending;
    expect(result.complete).toBe(true);
    expect(result.tasks.map((row) => row.before)).toEqual(['99001.1', '']);
    expect(result.saveAttempts).toBe(0);
    expect(app.writes).toEqual([]);
    expect(app.formLoads).toEqual([]);
  });
  it('preserves a real blank No. on an untouched task instead of comparing a generated outline', async () => {
    const app = setup(); const plan = swapPlan();
    document.body.insertAdjacentHTML('beforeend', '<div class="gantt-body-row" data-side="left" item-type="item" item-id="item_100143"><div class="gantt-cf19"></div></div>');
    plan.sourceTasks.push({ internalId: '100143', no: '' });
    const pending = runNoWriteExperimentFromPage(false, plan);
    await vi.runAllTimersAsync();
    expect((await pending).complete).toBe(true);
    expect(app.writes).toHaveLength(3);
    expect(document.querySelector('[item-id="item_100143"] .gantt-cf19')!.textContent).toBe('');
  });
  it('rejects swapped source IDs even if the overall set of numbers is unchanged', async () => {
    const app = setup(); const plan = swapPlan();
    plan.sourceTasks[0].no = '1.4'; plan.sourceTasks[1].no = '1.3';
    const result = await runNoWriteExperimentFromPage(false, plan);
    expect(result.error).toContain('ID 100119');
    expect(app.writes).toEqual([]);
  });
  it('refuses an old inventory without verified raw No. before activating the source', async () => {
    const executeScript = vi.fn(); vi.stubGlobal('chrome', { scripting: { executeScript } });
    const inventory = {
      project: { internalId: '870' }, collection: { complete: true }, tasks: [
        { internalId: '100119', code: 'TEST01.3', name: 'Test', no: '1.3' },
      ],
    } as OneOfficeProjectInventory;
    await expect(runNoUpdatePlan(42, inventory, 'TEST01.3\t1.5')).rejects.toThrow('No. gốc');
    expect(executeScript).not.toHaveBeenCalled();
  });
  it('rejects a missing No. cell instead of assuming it is blank', async () => {
    const app = setup(); document.querySelector('.gantt-cf19')!.remove();
    const result = await runNoWriteExperimentFromPage(false, swapPlan());
    expect(result.error).toContain('Không nhận diện');
    expect(app.writes).toEqual([]);
  });
  it('refuses production projects before touching the DOM', async () => {
    vi.stubGlobal('window', { location: { href: 'https://office.meygroup.vn/apps/work-project-project/view?ID=248' } });
    await expect(runNoWriteExperimentFromPage()).rejects.toThrow('ID 870');
    expect(document.documentElement.hasAttribute('data-tera-no-experiment')).toBe(false);
  });
  it('writes two allowed tasks through their actual forms and restores both', async () => {
    const { data, writes, anchorClicks, submitEvents } = setup();
    const promise = runNoWriteExperimentFromPage();
    await vi.runAllTimersAsync();
    const result = await promise;
    expect(result.error).toBeNull();
    expect(result.complete).toBe(true);
    expect(writes).toEqual(['100119', '100122', '100119', '100122']);
    expect(data.get('100119')!.no).toBe('1.3');
    expect(data.get('100122')!.no).toBe('1.4');
    expect(result.tasks.every((row) => row.applied && row.restored)).toBe(true);
    expect(result.saveAttempts).toBe(4);
    expect(anchorClicks.length).toBeGreaterThan(0);
    expect(anchorClicks.every((event) => event.defaultPrevented)).toBe(true);
    expect(submitEvents).toHaveLength(4);
    expect(submitEvents.every((event) => event.defaultPrevented)).toBe(true);
  });
  it('can verify both forms with zero save attempts, preventing dialog-link navigation', async () => {
    const app = setup();
    const promise = runNoWriteExperimentFromPage(true);
    await vi.runAllTimersAsync();
    const result = await promise;
    expect(result.complete).toBe(true);
    expect(result.readOnly).toBe(true);
    expect(result.saveAttempts).toBe(0);
    expect(result.tasks).toHaveLength(2);
    expect(app.writes).toEqual([]);
    expect(app.realClicks).toEqual([]);
    expect(app.menuPointerEvents).toEqual([]);
    expect(result.diagnostics?.[0]).toContain('Không mở dropdown');
    expect(app.closeEvents).toEqual(['form', 'detail', 'form']);
    expect(app.anchorClicks).toHaveLength(4);
    expect(app.anchorClicks.every((event) => event.defaultPrevented)).toBe(true);
    expect(app.data.get('100119')!.no).toBe('1.3');
    expect(app.data.get('100122')!.no).toBe('1.4');
    expect(document.querySelector('.task-quick-view a[href*="progress?ID=100122"]')).not.toBeNull();
    expect(document.querySelector('.task-quick-view a[href*="member?ID=100122"]')).toBeNull();
  });
  it.each(['id', 'code'] as const)('rejects a second detail with the wrong %s before any save', async (kind) => {
    const app = setup();
    if (kind === 'id') app.wrongSecondId(); else app.wrongSecondCode();
    const pending = runNoWriteExperimentFromPage(true);
    await vi.runAllTimersAsync();
    const result = await pending;
    expect(result.complete).toBe(false);
    expect(result.phase).toBe('Mở chi tiết TEST01.INSERT');
    expect(result.tasks).toHaveLength(1);
    expect(result.tasks[0].code).toBe('TEST01.3');
    expect(result.saveAttempts).toBe(0);
    expect(app.writes).toEqual([]);
  });
  it('does not retry an uncertain write or claim restoration', async () => {
    const app = setup(); app.failSecond();
    const promise = runNoWriteExperimentFromPage();
    await vi.runAllTimersAsync();
    const result = await promise;
    expect(result.complete).toBe(false);
    expect(result.error).toContain('chưa xác định kết quả');
    expect(app.writes).toEqual(['100119', '100122']);
    expect(result.tasks[0].restored).toBe(false);
    expect(result.saveAttempts).toBe(2);
  });
  it('cancels native link and form defaults even when addEventListener defers handlers', async () => {
    const app = setup(); app.navigateOnClick();
    const nativeAdd = HTMLElement.prototype.addEventListener;
    vi.spyOn(HTMLElement.prototype, 'addEventListener').mockImplementation(function (this: HTMLElement, type, listener, options) {
      if ((type === 'click' || type === 'submit') && typeof listener === 'function') {
        nativeAdd.call(this, type, (event: Event) => { setTimeout(() => listener.call(this, event), 5); }, options);
      } else nativeAdd.call(this, type, listener, options);
    });
    const before = window.location.href;
    const pending = runNoWriteExperimentFromPage(true);
    await vi.runAllTimersAsync();
    const result = await pending;
    expect(result.complete).toBe(true);
    expect(result.saveAttempts).toBe(0);
    expect(app.anchorClicks).toHaveLength(4);
    expect(app.anchorClicks.every((event) => event.defaultPrevented)).toBe(true);
    expect(window.location.href).toBe(before);
    expect(app.writes).toEqual([]);
    const writing = runNoWriteExperimentFromPage(false);
    await vi.runAllTimersAsync();
    expect((await writing).complete).toBe(true);
    expect(app.submitEvents).toHaveLength(4);
    expect(app.submitEvents.every((event) => event.defaultPrevented)).toBe(true);
    expect(window.location.href).toBe(before);
  });
  it('opens both cold forms via the exact verified URLs without opening either dropdown', async () => {
    const app = setup(); app.stallMenu();
    const promise = runNoWriteExperimentFromPage(true);
    await vi.runAllTimersAsync();
    const result = await promise;
    expect(result.complete).toBe(true);
    expect(result.saveAttempts).toBe(0);
    expect(result.tasks).toHaveLength(2);
    expect(app.writes).toEqual([]);
    expect(result.diagnostics?.every((line) => line.includes('URL biểu mẫu test đã xác minh'))).toBe(true);
    expect(app.menuPointerEvents).toEqual([]);
    expect(document.querySelector('[data-tera-no-form-link]')).toBeNull();
  });
  it('refuses an individually hidden edit action even in a collapsed dropdown', async () => {
    const app = setup(); app.hideEdit();
    const pending = runNoWriteExperimentFromPage(true);
    await vi.runAllTimersAsync();
    const result = await pending;
    expect(result.complete).toBe(false);
    expect(result.diagnostics?.[0]).toContain('không hợp lệ hoặc bị vô hiệu hóa');
    expect(result.saveAttempts).toBe(0);
    expect(app.writes).toEqual([]);
  });
  it('refuses an edit link with an unexpected title without saving', async () => {
    const app = setup(); app.otherTitle();
    const pending = runNoWriteExperimentFromPage(true);
    await vi.runAllTimersAsync();
    const result = await pending;
    expect(result.complete).toBe(false);
    expect(result.diagnostics?.[0]).toContain('không hợp lệ hoặc bị vô hiệu hóa');
    expect(result.saveAttempts).toBe(0);
    expect(app.writes).toEqual([]);
    expect(result.diagnostics?.join(' ')).not.toMatch(/https:|<a|token=/);
  });
  it.each(['disabled', 'foreign'] as const)('refuses a %s edit action without opening it', async (kind) => {
    const app = setup();
    if (kind === 'disabled') app.disableEdit(); else app.foreignEdit();
    const pending = runNoWriteExperimentFromPage(true);
    await vi.runAllTimersAsync();
    const result = await pending;
    expect(result.complete).toBe(false);
    expect(result.saveAttempts).toBe(0);
    expect(app.anchorClicks).toHaveLength(1);
    expect(app.writes).toEqual([]);
    expect(app.menuPointerEvents).toEqual([]);
  });
  it('stops without saving and removes its temporary link when the server form never appears', async () => {
    const app = setup(); app.denySecondForm();
    const pending = runNoWriteExperimentFromPage(true);
    await vi.runAllTimersAsync();
    const result = await pending;
    expect(result.complete).toBe(false);
    expect(result.phase).toBe('Tải biểu mẫu TEST01.INSERT');
    expect(result.saveAttempts).toBe(0);
    expect(app.writes).toEqual([]);
    expect(document.querySelector('[data-tera-no-form-link]')).toBeNull();
  });
  it('blocks a pre-existing temporary No. before any write', async () => {
    const app = setup(); document.querySelector('.gantt-cf19')!.textContent = '99001.1';
    expect((await runNoWriteExperimentFromPage()).error).toContain('No. tạm đã có');
    expect(app.writes).toEqual([]);
  });
  it('refuses concurrent runs', async () => {
    setup(); document.documentElement.setAttribute('data-tera-no-experiment', 'running');
    await expect(runNoWriteExperimentFromPage()).rejects.toThrow('Đang có');
  });
  it('refuses to start while the source document is hidden', async () => {
    const app = setup(); app.hide();
    await expect(runNoWriteExperimentFromPage(true)).rejects.toThrow('đang ở nền');
    expect(document.documentElement.hasAttribute('data-tera-no-experiment')).toBe(false);
    expect(app.writes).toEqual([]);
  });
  it('does not resend or restore an uncertain write after the source becomes hidden', async () => {
    const app = setup(); app.hideOnSecondWrite();
    const pending = runNoWriteExperimentFromPage(false);
    await vi.runAllTimersAsync();
    const result = await pending;
    expect(result.complete).toBe(false);
    expect(result.error).toContain('đang ở nền');
    expect(result.error).toContain('Có yêu cầu ghi chưa xác định kết quả');
    expect(app.writes).toEqual(['100119', '100122']);
    expect(result.tasks[0].restored).toBe(false);
    expect(document.documentElement.hasAttribute('data-tera-no-experiment')).toBe(false);
  });
  it('injects the self-contained runner in the page realm without an arbitrary bridge', async () => {
    const result = { readOnly: true, complete: true, tasks: [], saveAttempts: 0 };
    const executeScript = vi.fn().mockResolvedValue([{ result }]);
    const get = vi.fn().mockResolvedValue({ url: 'https://office.meygroup.vn/apps/work-project-project/view?ID=870' });
    const update = vi.fn().mockResolvedValue({ active: true });
    vi.stubGlobal('chrome', { scripting: { executeScript }, tabs: { get, update } });
    expect(await runNoWriteExperiment(42, true)).toBe(result);
    expect(executeScript).toHaveBeenCalledExactlyOnceWith({
      target: { tabId: 42 }, world: 'MAIN', func: runNoWriteExperimentFromPage, args: [true],
    });
    expect(get).toHaveBeenCalledExactlyOnceWith(42);
    expect(update).toHaveBeenCalledExactlyOnceWith(42, { active: true });
    expect(get.mock.invocationCallOrder[0]).toBeLessThan(update.mock.invocationCallOrder[0]);
    expect(update.mock.invocationCallOrder[0]).toBeLessThan(executeScript.mock.invocationCallOrder[0]);
  });
  it('does not activate or inject into a non-test source tab', async () => {
    const get = vi.fn().mockResolvedValue({ url: 'https://office.meygroup.vn/apps/work-project-project/view?ID=871' });
    const update = vi.fn(); const executeScript = vi.fn();
    vi.stubGlobal('chrome', { scripting: { executeScript }, tabs: { get, update } });
    await expect(runNoWriteExperiment(42, true)).rejects.toThrow('ID 870');
    expect(update).not.toHaveBeenCalled();
    expect(executeScript).not.toHaveBeenCalled();
  });
});
