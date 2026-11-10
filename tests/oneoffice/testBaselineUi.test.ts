import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { afterEach, expect, it, vi } from 'vitest';
import { TestBaselinePanel } from '../../src/import-review/TestBaselinePanel';
import { collectTestBaseline, downloadTestBaseline, type TestBaseline } from '../../src/oneoffice/testBaseline';
import { findTestPreservationSample } from '../../src/oneoffice/testPreservation';

vi.mock('../../src/oneoffice/testBaseline', () => ({ collectTestBaseline: vi.fn(), downloadTestBaseline: vi.fn() }));
vi.mock('../../src/oneoffice/testPreservation', () => ({ findTestPreservationSample: vi.fn() }));
afterEach(() => { vi.resetAllMocks(); vi.unstubAllGlobals(); document.body.innerHTML = ''; });

it('waits for an explicit click, shows coverage gaps, and downloads only on request', async () => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
  const baseline = { mode: 'read-only-test-baseline', origin: 'https://testoffice.meygroup.vn:9443',
    startedAt: '2026-10-21', completedAt: '2026-10-21', inventory: { tasks: Array.from({ length: 864 }, () => ({ code: null })), collection: { warnings: ['Chưa đọc được mã từ nguồn danh sách.'] } },
    samples: [{ task: { internalId: '7459', code: 'TEST', name: 'Sample', sourceNo: '', parentInternalId: null },
      reason: 'Mẫu cha', discussion: { collection: { complete: true }, comments: [] }, error: null,
      details: { identityVerified: true, code: { state: 'value', displayValue: '.18902' }, no: { state: 'empty', displayValue: '' } } }],
    limitations: ['Chưa thu thập phần trăm hoàn thành.'] } as unknown as TestBaseline;
  vi.mocked(collectTestBaseline).mockResolvedValue(baseline);
  const container = document.createElement('div'); document.body.append(container); const root = createRoot(container);
  const onBusyChange = vi.fn();
  try {
    await act(async () => root.render(React.createElement(TestBaselinePanel, { sourceTabId: 42, onBusyChange })));
    expect(collectTestBaseline).not.toHaveBeenCalled();
    await act(async () => container.querySelector('button')!.click());
    expect(collectTestBaseline).toHaveBeenCalledExactlyOnceWith(42, expect.any(Function), expect.any(Function));
    expect(onBusyChange.mock.calls).toEqual([[true], [false]]);
    expect(container.textContent).toContain('864 công việc');
    expect(container.textContent).toContain('No. gốc đã xác minh 0/864');
    expect(container.textContent).toContain('đọc được mã 0/864');
    expect(container.textContent).toContain('Bản 4');
    expect(container.textContent).toContain('mã .18902; No. trong phản hồi (trả về trống)');
    expect(container.textContent).toContain('Chưa đọc được mã từ nguồn danh sách.');
    expect(container.textContent).toContain('Chưa đủ điều kiện kiểm thử bảo toàn thảo luận');
    expect(container.querySelector('tbody')?.textContent).toContain('(trống)');
    expect(downloadTestBaseline).not.toHaveBeenCalled();
    await act(async () => [...container.querySelectorAll('button')].find((button) => button.textContent?.startsWith('Tải bản hiện trạng'))!.click());
    expect(downloadTestBaseline).toHaveBeenCalledExactlyOnceWith({ ...baseline, preservation: null });
    await act(async () => root.unmount());
    expect(container.textContent).toBe('');
  } finally { if (container.textContent) await act(async () => root.unmount()); }
});

it('prevents duplicate reads and discards a late result after unmount', async () => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
  let resolve!: (value: TestBaseline) => void;
  vi.mocked(collectTestBaseline).mockReturnValue(new Promise((done) => { resolve = done; }));
  const container = document.createElement('div'); document.body.append(container); const root = createRoot(container);
  const onBusyChange = vi.fn();
  await act(async () => root.render(React.createElement(TestBaselinePanel, { sourceTabId: 42, onBusyChange })));
  await act(async () => { container.querySelector('button')!.click(); container.querySelector('button')!.click(); });
  expect(collectTestBaseline).toHaveBeenCalledTimes(1);
  const cancelled = vi.mocked(collectTestBaseline).mock.calls[0][2];
  await act(async () => root.unmount());
  expect(cancelled()).toBe(true);
  await act(async () => resolve({} as TestBaseline));
  expect(onBusyChange).toHaveBeenCalledTimes(1);
  expect(downloadTestBaseline).not.toHaveBeenCalled();
});

it('starts the bounded search only on click, shows progress evidence separately and exports it locally', async () => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
  const baseline = { startedAt: 'start', completedAt: 'end', origin: 'test', inventory: { tasks: [], collection: { warnings: [] } }, samples: [], limitations: [] } as unknown as TestBaseline;
  const probe = { mode: 'read-only-test-preservation', attempts: [{ taskId: '7698', state: 'found' }], remaining: 863,
    sample: { task: { internalId: '7698', name: 'Sample' }, discussion: { comments: [{ id: '1' }] }, details: { identityVerified: true, progress: { state: 'value', percent: 0 } } } } as any;
  vi.mocked(collectTestBaseline).mockResolvedValue(baseline);
  vi.mocked(findTestPreservationSample).mockResolvedValue(probe);
  const container = document.createElement('div'); document.body.append(container); const root = createRoot(container);
  try {
    await act(async () => root.render(React.createElement(TestBaselinePanel, { sourceTabId: 42, onBusyChange: vi.fn() })));
    await act(async () => container.querySelector('button')!.click());
    expect(findTestPreservationSample).not.toHaveBeenCalled();
    const find = () => [...container.querySelectorAll('button')].find((button) => button.textContent?.startsWith('Tìm mẫu có thảo luận'))!;
    await act(async () => { find().click(); find().click(); });
    expect(findTestPreservationSample).toHaveBeenCalledExactlyOnceWith(42, baseline, null, null, expect.any(Function), expect.any(Function));
    expect(container.textContent).toContain('0% (phản hồi chi tiết)');
    expect(container.textContent).toContain('Báo cáo thực tế: chưa thu thập');
    expect(container.textContent).not.toContain('Chưa có mẫu xác minh được thảo luận có sẵn');
    expect(downloadTestBaseline).not.toHaveBeenCalled();
    await act(async () => [...container.querySelectorAll('button')].find((button) => button.textContent?.startsWith('Tải bản hiện trạng'))!.click());
    expect(downloadTestBaseline).toHaveBeenCalledWith({ ...baseline, preservation: probe });
  } finally { await act(async () => root.unmount()); }
});

it('supports stopping a pending search and discards results when the panel is closed', async () => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
  const baseline = { inventory: { tasks: [], collection: { warnings: [] } }, samples: [], limitations: [] } as unknown as TestBaseline;
  vi.mocked(collectTestBaseline).mockResolvedValue(baseline);
  let resolve!: (result: any) => void;
  vi.mocked(findTestPreservationSample).mockReturnValue(new Promise((done) => { resolve = done; }));
  const container = document.createElement('div'); document.body.append(container); const root = createRoot(container);
  await act(async () => root.render(React.createElement(TestBaselinePanel, { sourceTabId: 42, onBusyChange: vi.fn() })));
  await act(async () => container.querySelector('button')!.click());
  await act(async () => [...container.querySelectorAll('button')].find((button) => button.textContent?.startsWith('Tìm mẫu có thảo luận'))!.click());
  const cancelled = vi.mocked(findTestPreservationSample).mock.calls[0][5];
  expect(cancelled()).toBe(false);
  await act(async () => [...container.querySelectorAll('button')].find((button) => button.textContent === 'Dừng tìm mẫu')!.click());
  expect(cancelled()).toBe(true);
  await act(async () => root.unmount());
  await act(async () => resolve({ sample: null }));
  expect(container.textContent).toBe('');
  expect(downloadTestBaseline).not.toHaveBeenCalled();
});
