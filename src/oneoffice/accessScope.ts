import { collectAccessReport, reportFromAccessSnapshot, type AccessPage, type AccessPersonnel, type AccessRecord, type AccessReport, type AccessSnapshot } from './accessHistory';
import { matchesCompanyScope } from './companyScopes';

export function activeAccessSnapshot(snapshot: AccessSnapshot, personnel: AccessPersonnel): AccessSnapshot {
  if (!Array.isArray(personnel.activeAccounts)) throw new Error('Chưa xác minh được tài khoản active. Không xuất báo cáo.');
  const active = new Set(personnel.activeAccounts);
  const records = snapshot.records.filter((row) => active.has(row.account));
  const oldest = records.reduce<AccessRecord | undefined>((previous, row) => !previous || row.timeIso < previous.timeIso ? row : previous, undefined);
  return { ...snapshot, records, recordCount: records.length, oldestTime: oldest?.timeRaw ?? '', personnel,
    scope: snapshot.scope ?? 'Nhật ký chỉ của tài khoản hiện còn active, xác minh từ danh sách người dùng 1Office.' };
}

export interface AccessFilters { company: string; position: string; department: string }

export function selectAccessAccounts(personnel: AccessPersonnel, filters: AccessFilters): string[] {
  if (!Array.isArray(personnel.activeAccounts)) throw new Error('Chưa xác minh được tài khoản active. Không đọc nhật ký.');
  const active = new Set(personnel.activeAccounts);
  const normalize = (value: string) => value.normalize('NFC').toLocaleLowerCase('vi').trim();
  const company = normalize(filters.company);
  const position = normalize(filters.position);
  const department = normalize(filters.department);
  if (company && !personnel.users.some((user) => user.companyCode || user.companyGroup)) {
    throw new Error('Chưa xác minh được cơ cấu công ty. Không quét thay thế toàn tập đoàn.');
  }
  return personnel.users.filter((user) => active.has(user.account) && (!company ||
    (matchesCompanyScope(company, user.companyCode ?? '', user.companyGroup) ??
      (normalize(user.companyCode ?? '') === company || normalize(user.company ?? '').includes(company)))) && (!position || normalize(user.position).includes(position)) &&
    (!department || normalize(user.department).includes(department) || normalize(user.departmentCode ?? '').includes(department)))
    .map((user) => user.account);
}

// Multiple account searches share one six-request limit, including collector rechecks.
export async function collectScopedAccessReport(
  cutoff: string, accounts: string[], read: (page: number, account: string) => Promise<AccessPage>,
  signal: AbortSignal, progress: (loaded: number, total: number) => void, scope: string,
): Promise<AccessReport> {
  const selected = [...new Set(accounts)];
  if (!selected.length) throw new Error('Không có tài khoản khớp bộ lọc. Chưa đọc nhật ký.');
  let active = 0;
  const waiting: Array<() => void> = [];
  let failure: unknown;
  const limited = async (page: number, account: string) => {
    if (active >= 6) await new Promise<void>((resolve) => waiting.push(resolve));
    else active++;
    try {
      if (signal.aborted) throw new Error('Đã hủy thu thập nhật ký.');
      if (failure) throw failure;
      return await read(page, account);
    } finally {
      const next = waiting.shift();
      if (next) next(); else active--;
    }
  };
  const records = new Map<string, AccessRecord>();
  let index = 0;
  let finished = 0;
  progress(0, selected.length);
  const workers = Array.from({ length: Math.min(6, selected.length) }, async () => {
    try {
      while (index < selected.length && !failure && !signal.aborted) {
        const account = selected[index++];
        const result = await collectAccessReport(cutoff, (page) => limited(page, account), signal, () => {});
        // Keyword search can match substrings or other fields. Only exact accounts belong to this scope.
        for (const row of result.snapshot?.records ?? []) {
          if (row.account !== account) continue;
          const previous = records.get(row.id);
          if (previous && JSON.stringify(previous) !== JSON.stringify(row)) throw new Error('ID nhật ký có dữ liệu mâu thuẫn.');
          records.set(row.id, row);
          if (records.size > 100000) throw new Error('Phạm vi vượt giới hạn 100.000 bản ghi.');
        }
        progress(++finished, selected.length);
      }
    } catch (error) { failure ??= error; throw error; }
  });
  await Promise.allSettled(workers);
  if (failure) throw failure;
  if (signal.aborted) throw new Error('Đã hủy thu thập nhật ký.');
  const rows = [...records.values()];
  const oldest = rows.reduce<AccessRecord | undefined>((previous, row) => !previous || row.timeIso < previous.timeIso ? row : previous, undefined);
  return reportFromAccessSnapshot({ scope, collectedAt: new Date().toISOString(), recordCount: rows.length,
    oldestTime: oldest?.timeRaw ?? '', records: rows }, cutoff);
}
