import type { NoExperimentResult, NoWriteRequest } from './noWriteExperiment';
import { isNoTestTask, NO_TEST_TASKS } from './noTestScope';

const KEY = 'oneOfficeNoWriteJournalV1';
export interface NoWriteJournal {
  version: 1;
  id: string;
  origin: 'https://office.meygroup.vn';
  projectId: '870';
  createdAt: string;
  updatedAt: string;
  status: 'pending' | 'verified' | 'unchanged';
  rows: { internalId: string; code: string; before: string; after: string; temporary: string; observed: string | null; state: 'unknown' | 'before' | 'temporary' | 'after' | 'conflict' }[];
}

function storageCall(method: 'get' | 'set' | 'remove', value: unknown): Promise<any> {
  const api = (globalThis as any).chrome;
  if (!api?.storage?.local?.[method]) return Promise.reject(new Error('Không truy cập được nhật ký cục bộ. Chưa cho phép ghi.'));
  return new Promise((resolve, reject) => {
    api.storage.local[method](value, (result: unknown) => {
      if (api.runtime?.lastError) reject(new Error('Không lưu/đọc được nhật ký cục bộ. Kiểm tra trạng thái trước khi ghi tiếp.'));
      else resolve(result);
    });
  });
}

export async function readNoWriteJournal(): Promise<NoWriteJournal | null> {
  const value = (await storageCall('get', KEY))?.[KEY];
  if (value === undefined) return null;
  // Treat corrupt/incompatible records as unresolved, never silently replace them.
  if (!value || value.version !== 1 || value.origin !== 'https://office.meygroup.vn' || value.projectId !== '870'
    || typeof value.id !== 'string' || !['pending', 'verified', 'unchanged'].includes(value.status)
    || !Array.isArray(value.rows) || !value.rows.length || value.rows.length > 5
    || value.rows.some((row: any) => !row || !isNoTestTask(row)
      || typeof row.before !== 'string' || (row.before !== '' && !/^[1-9]\d*(?:\.[1-9]\d*)*$/.test(row.before))
      || ['after', 'temporary'].some((field) => typeof row[field] !== 'string' || !/^[1-9]\d*(?:\.[1-9]\d*)*$/.test(row[field]))
      || row.temporary !== NO_TEST_TASKS.find((task) => task.internalId === row.internalId)?.temporary
      || !['unknown', 'before', 'temporary', 'after', 'conflict'].includes(row.state)
      || !(row.observed === null || typeof row.observed === 'string'))
    || new Set(value.rows.map((row: any) => row.internalId)).size !== value.rows.length) {
    throw new Error('Nhật ký ghi không hợp lệ. Dừng ghi, cần kiểm tra thủ công; không tự xóa nhật ký.');
  }
  if ((value.status === 'verified' && !value.rows.every((row: any) => row.state === 'after' && row.observed === row.after))
    || (value.status === 'unchanged' && !value.rows.every((row: any) => row.state === 'before' && row.observed === row.before))) {
    throw new Error('Trạng thái nhật ký không khớp các công việc. Dừng ghi.');
  }
  // No extra page-provided fields are returned or persisted.
  return {
    version: 1, id: value.id, origin: value.origin, projectId: value.projectId,
    createdAt: String(value.createdAt), updatedAt: String(value.updatedAt), status: value.status,
    rows: value.rows.map((row: any) => ({ internalId: row.internalId, code: row.code, before: row.before, after: row.after,
      temporary: row.temporary, observed: row.observed, state: row.state })),
  };
}

async function persist(journal: NoWriteJournal) {
  await storageCall('set', { [KEY]: journal });
  const saved = await readNoWriteJournal();
  if (JSON.stringify(saved) !== JSON.stringify(journal)) throw new Error('Chưa xác nhận lưu nhật ký. Dừng ghi.');
  return journal;
}

export async function prepareNoWriteJournal(request: NoWriteRequest) {
  const previous = await readNoWriteJournal();
  if (previous?.status === 'pending') throw new Error('Có lượt ghi chưa đối soát. Bấm kiểm tra trạng thái, không gửi lại bản cũ.');
  const now = new Date().toISOString();
  return persist({
    version: 1, id: crypto.randomUUID(), origin: 'https://office.meygroup.vn', projectId: '870',
    createdAt: now, updatedAt: now, status: 'pending',
    rows: request.changes.map((row) => ({ internalId: row.internalId, code: row.code, before: row.before, after: row.after,
      temporary: NO_TEST_TASKS.find((task) => task.internalId === row.internalId)!.temporary, observed: null, state: 'unknown' })),
  });
}

export async function finishNoWriteJournal(journal: NoWriteJournal, result: NoExperimentResult) {
  // Only a successful read-back permits marking a row verified. On interruption,
  // intent stays pending until a fresh read, even if no save was reported.
  const rows = journal.rows.map((row) => {
    const outcome = result.tasks.find((task) => task.code === row.code && task.before === row.before && task.after === row.after);
    return outcome?.finalized ? { ...row, observed: row.after, state: 'after' as const } : row;
  });
  return persist({ ...journal, updatedAt: new Date().toISOString(), rows,
    status: result.complete && rows.every((row) => row.state === 'after') ? 'verified' : 'pending' });
}

export async function reconcileNoWriteJournal(journal: NoWriteJournal, result: NoExperimentResult) {
  if (!result.readOnly || !result.complete || result.saveAttempts !== 0) throw new Error('Chưa đọc đầy đủ trạng thái. Giữ nguyên nhật ký và khóa ghi.');
  const current = await readNoWriteJournal();
  if (!current || current.id !== journal.id) throw new Error('Nhật ký đã thay đổi trong lúc kiểm tra.');
  const rows = journal.rows.map((row) => {
    const observed = result.tasks.find((task) => task.code === row.code)?.before;
    if (observed === undefined) throw new Error('Thiếu trạng thái công việc trong kết quả kiểm tra.');
    const state = observed === row.after ? 'after' : observed === row.before ? 'before' : observed === row.temporary ? 'temporary' : 'conflict';
    return { ...row, observed, state } as NoWriteJournal['rows'][number];
  });
  return persist({ ...journal, updatedAt: new Date().toISOString(), rows,
    status: rows.every((row) => row.state === 'after') ? 'verified' : rows.every((row) => row.state === 'before') ? 'unchanged' : 'pending' });
}

export async function clearResolvedNoWriteJournal() {
  return withNoWriteLock(async () => {
    const journal = await readNoWriteJournal();
    if (journal?.status === 'pending') throw new Error('Không xóa nhật ký khi lượt ghi chưa đối soát.');
    await storageCall('remove', KEY);
  });
}

// Web Locks serialize all extension review tabs, but do not lock other 1Office users.
export async function withNoWriteLock<T>(action: () => Promise<T>): Promise<T> {
  if (!navigator.locks?.request) throw new Error('Trình duyệt chưa hỗ trợ khóa lượt ghi. Chưa cho phép ghi.');
  return navigator.locks.request('tera-no-write-870', { ifAvailable: true }, async (lock) => {
    if (!lock) throw new Error('Một trang T-eraClip khác đang ghi hoặc kiểm tra lượt ghi này.');
    return action();
  });
}
