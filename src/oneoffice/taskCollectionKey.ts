import type { OneOfficeProjectTaskInventoryItem } from './types';

// A report row key is local bookkeeping, never a 1Office task ID.
export const taskCollectionKey = (task: OneOfficeProjectTaskInventoryItem): string =>
  task.collectionKey ?? task.internalId;
