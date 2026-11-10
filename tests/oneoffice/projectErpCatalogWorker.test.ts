import { afterEach, expect, it, vi } from 'vitest';

const reader = vi.hoisted(() => ({ readSheet: vi.fn() }));
vi.mock('read-excel-file/browser', () => reader);
afterEach(() => { vi.unstubAllGlobals(); vi.resetModules(); reader.readSheet.mockReset(); });

async function worker() {
  const scope = { postMessage: vi.fn(), onmessage: null as null | ((event: { data: { buffer: ArrayBuffer; sheet: string } }) => Promise<void>) };
  vi.stubGlobal('self', scope);
  await import('../../src/oneoffice/projectErpCatalog.worker');
  return scope;
}

it('reads only the selected sheet and posts only the allowed catalog entries', async () => {
  reader.readSheet.mockResolvedValue([
    ['1. Danh mục đối tượng'], ['Mã', 'Tên'], ['PRIVATE', 'Sensitive name'],
    ['2. Danh mục gói thầu'], ['Mã', 'Tên'], ['GT.001', 'Package'],
    ['17. Danh mục hạng mục'], ['Mã', 'Tên'], ['P', 'Project'],
  ]);
  const scope = await worker(); const buffer = new ArrayBuffer(12);
  await scope.onmessage!({ data: { buffer, sheet: ' TDR ' } });
  expect(reader.readSheet).toHaveBeenCalledExactlyOnceWith(buffer, 'TDR');
  const result = scope.postMessage.mock.calls[0][0];
  expect(result.catalog.packages).toHaveLength(1); expect(result.catalog.projects).toHaveLength(1);
  expect(JSON.stringify(result)).not.toContain('PRIVATE'); expect(JSON.stringify(result)).not.toContain('Sensitive');
});

it('reports missing sheets and unsupported catalog structure instead of returning an empty success', async () => {
  const scope = await worker();
  const missing = new Error('Missing'); missing.name = 'SheetNotFoundError'; reader.readSheet.mockRejectedValueOnce(missing);
  await scope.onmessage!({ data: { buffer: new ArrayBuffer(0), sheet: 'Wrong' } });
  expect(scope.postMessage.mock.lastCall![0].error).toContain('Không tìm thấy sheet');
  reader.readSheet.mockResolvedValueOnce([['Mã', 'Tên'], ['A', 'Person']]);
  await scope.onmessage!({ data: { buffer: new ArrayBuffer(0), sheet: 'Wrong' } });
  expect(scope.postMessage.mock.lastCall![0].error).toContain('Không tìm thấy Danh mục gói thầu');
});
