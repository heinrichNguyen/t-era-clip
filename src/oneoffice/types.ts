export type OneOfficeEntityType = 'project' | 'task' | 'report';

export interface OneOfficeEntityContext {
  type: OneOfficeEntityType;
  internalId: string;
  code: string | null;
  name: string;
  url: string;
}

export interface OneOfficeAttachment {
  index: number;
  name: string;
  sizeText: string | null;
  thumbnailUrl: string | null;
  originalDownloadAvailable: boolean;
}

export interface OneOfficeMention {
  userId: string | null;
  displayName: string;
}

export interface OneOfficeComment {
  id: string;
  parentId: string | null;
  rootId: string;
  depth: number;
  sourceOrder: number;
  authorId: string | null;
  authorName: string;
  createdAtRaw: string;
  createdAtIso: string | null;
  contentText: string;
  mentions: OneOfficeMention[];
  attachments: OneOfficeAttachment[];
}

export interface OneOfficeCollectionStatus {
  complete: boolean;
  loadMoreClicks: number;
  loadedRootCommentCount: number;
  loadedReplyCount: number;
  expectedRootCommentCount: number | null;
  warnings: string[];
}

export interface OneOfficeDiscussionExport {
  schemaVersion: 1;
  source: '1office-dom' | '1office-comment-api';
  capturedAt: string;
  entity: OneOfficeEntityContext;
  collection: OneOfficeCollectionStatus;
  comments: OneOfficeComment[];
}

export interface OneOfficeProjectTaskInventoryItem {
  collectionKey?: string;
  identityError?: string;
  identityWarning?: string;
  reportLookupUrl?: string;
  projectInternalId?: string;
  projectName?: string;
  parentName?: string;
  sourceOrder: number;
  internalId: string;
  no: string;
  // Exact cf19 value, including blank. Undefined means it was not verified.
  sourceNo?: string;
  code: string | null;
  name: string;
  url: string;
  depth: number;
  parentNo: string | null;
  parentInternalId: string | null;
  assigneeNames?: string[];
  assignerNames?: string[];
  departmentNames?: string[];
  taskStatus?: string;
  baselineStart?: string;
  baselineEnd?: string;
  plannedStart?: string;
  plannedEnd?: string;
  actualStart?: string;
  actualEnd?: string;
  overdueDays?: string;
  completedOverdueDays?: string;
  labels?: string[];
}

export interface OneOfficeProjectInventoryStatus {
  complete: boolean;
  expectedTaskCount: number | null;
  loadedTaskCount: number;
  warnings: string[];
}

export interface OneOfficeProjectInventory {
  schemaVersion: 1;
  source: '1office-project-dom' | '1office-project-api' | '1office-report-dom';
  capturedAt: string;
  project: OneOfficeEntityContext;
  collection: OneOfficeProjectInventoryStatus;
  tasks: OneOfficeProjectTaskInventoryItem[];
}

export type OneOfficeProjectTaskCollectionState =
  | 'pending'
  | 'collecting'
  | 'complete'
  | 'incomplete'
  | 'failed'
  | 'unresolved'
  | 'cancelled';

export interface OneOfficeProjectTaskCollectionResult {
  task: OneOfficeProjectTaskInventoryItem;
  state: Exclude<OneOfficeProjectTaskCollectionState, 'pending' | 'collecting'>;
  discussion: OneOfficeDiscussionExport | null;
  error: string | null;
}

export interface OneOfficeProjectBatchResult {
  schemaVersion: 1;
  source: '1office-project-batch-dom';
  startedAt: string;
  completedAt: string;
  cancelled: boolean;
  project: OneOfficeEntityContext;
  inventory: OneOfficeProjectInventory;
  tasks: OneOfficeProjectTaskCollectionResult[];
  warnings: string[];
}

export interface ParseOneOfficeDiscussionOptions {
  capturedAt?: Date | string;
  loadMoreClicks?: number;
  sourceUrl?: string;
  timeZoneOffset?: string;
}
