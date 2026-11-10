export const NO_TEST_TASKS = [
  { internalId: '100119', code: 'TEST01.3', temporary: '99001.1' },
  { internalId: '100122', code: 'TEST01.INSERT', temporary: '99001.2' },
  { internalId: '100113', code: 'TEST01.01', temporary: '99001.3' },
  { internalId: '100116', code: 'TEST01.2', temporary: '99001.4' },
  { internalId: '100143', code: 'TEST01.4', temporary: '99001.5' },
] as const;

export function isNoTestTask(row: { internalId: string; code: string | null }) {
  return NO_TEST_TASKS.some((task) => task.internalId === row.internalId && task.code === row.code);
}
