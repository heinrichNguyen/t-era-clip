import React, { act, useState } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import { useTemporaryData } from '../../src/oneoffice/temporaryData';
import { ClearTemporaryDataButton, TemporaryDataCleared } from '../../src/shared/TemporaryDataControls';

let root: Root;
let container: HTMLDivElement;
let changes: Array<(value: Record<string, unknown>, area: string) => void>;
let stored: Record<string, unknown>;
let api: any;
const cleared = vi.fn();

function Page({ name }: { name: string }) {
  const [data, setData] = useState('private report');
  const temporary = useTemporaryData(() => { setData(''); cleared(name); });
  return React.createElement('section', { 'data-page': name },
    temporary.cleared ? React.createElement(TemporaryDataCleared) : React.createElement(React.Fragment, null,
      React.createElement('p', null, data), React.createElement('button', null, 'Xuất Excel'),
      React.createElement(ClearTemporaryDataButton, { disabled: temporary.clearing, error: temporary.clearError,
        onClick: () => void temporary.clearTemporaryData() })));
}

beforeEach(() => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
  vi.spyOn(window, 'confirm').mockReturnValue(true);
  cleared.mockClear(); changes = [];
  stored = { oneOfficeDiscussionReview: { private: 'discussion' }, oneOfficeProjectBatchSession: { private: 'inventory' }, unrelated: 'preserve' };
  api = { runtime: {}, storage: {
    local: { remove: vi.fn(), clear: vi.fn() },
    onChanged: { addListener: (fn: typeof changes[number]) => changes.push(fn),
      removeListener: (fn: typeof changes[number]) => { changes = changes.filter((value) => value !== fn); } },
    session: {
      remove: vi.fn((keys: string[], done: () => void) => { keys.forEach((key) => delete stored[key]); done(); }),
      set: vi.fn((value: Record<string, unknown>, done: () => void) => {
        Object.assign(stored, value);
        changes.forEach((listener) => listener(Object.fromEntries(Object.keys(value).map((key) => [key, { newValue: value[key] }])), 'session'));
        done();
      }),
    },
  } };
  vi.stubGlobal('chrome', api);
  container = document.createElement('div'); document.body.append(container); root = createRoot(container);
});

afterEach(async () => {
  await act(async () => root.unmount()); container.remove();
  expect(changes).toHaveLength(0); vi.restoreAllMocks(); vi.unstubAllGlobals();
});

const render = () => act(async () => root.render(React.createElement(React.Fragment, null,
  React.createElement(Page, { name: 'discussion' }), React.createElement(Page, { name: 'access' }))));
const clear = () => act(async () => [...container.querySelectorAll('button')].find((button) => button.textContent === 'Xóa dữ liệu tạm')!.click());

it('clears session report keys and all listening pages without touching Capture or unrelated data', async () => {
  await render(); await clear();
  expect(stored).not.toHaveProperty('oneOfficeDiscussionReview');
  expect(stored).not.toHaveProperty('oneOfficeProjectBatchSession');
  expect(stored.unrelated).toBe('preserve');
  expect(api.storage.local.remove).not.toHaveBeenCalled(); expect(api.storage.local.clear).not.toHaveBeenCalled();
  expect(container.textContent).not.toContain('private report');
  expect(container.textContent).not.toContain('Xuất Excel');
  expect(cleared).toHaveBeenCalledWith('discussion'); expect(cleared).toHaveBeenCalledWith('access');
});

it('does nothing when the user cancels', async () => {
  vi.mocked(window.confirm).mockReturnValue(false);
  await render(); await clear();
  expect(api.storage.session.remove).not.toHaveBeenCalled();
  expect(cleared).not.toHaveBeenCalled(); expect(container.textContent).toContain('private report');
});

it('ignores unrelated storage events', async () => {
  await render();
  await act(async () => {
    changes.forEach((listener) => listener({ unrelated: {} }, 'session'));
    changes.forEach((listener) => listener({ oneOfficeTemporaryClearRevision: {} }, 'local'));
  });
  expect(cleared).not.toHaveBeenCalled();
});

it('reports deletion failures instead of claiming that the reports were erased', async () => {
  api.storage.session.remove.mockImplementation((_keys: string[], done: () => void) => {
    api.runtime.lastError = { message: 'Storage unavailable' }; done(); delete api.runtime.lastError;
  });
  await render(); await clear();
  expect(container.textContent).toContain('Storage unavailable');
  expect(container.textContent).toContain('private report'); expect(cleared).not.toHaveBeenCalled();
  expect(api.storage.session.set).not.toHaveBeenCalled();
});
