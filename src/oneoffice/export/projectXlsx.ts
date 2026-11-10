import writeExcelFile, {
  type CellObject,
  type Sheet,
  type SheetData,
} from 'write-excel-file/browser';
import type {
  OneOfficeComment,
  OneOfficeProjectBatchResult,
  OneOfficeProjectTaskCollectionResult,
} from '../types';
import { describeDiscussionFilter, filterDiscussionComments, type DiscussionFilter } from '../discussionFilter';

const HEADER_STYLE: Omit<CellObject, 'value'> = {
  backgroundColor: '#0E7490',
  borderColor: '#155E75',
  borderStyle: 'thin',
  fontWeight: 'bold',
  textColor: '#FFFFFF',
  alignVertical: 'center',
  wrap: true,
};

const LABEL_STYLE: Omit<CellObject, 'value'> = {
  backgroundColor: '#E2E8F0',
  borderColor: '#CBD5E1',
  borderStyle: 'thin',
  fontWeight: 'bold',
  textColor: '#0F172A',
  alignVertical: 'top',
};

const DATA_STYLE: Omit<CellObject, 'value'> = {
  borderColor: '#CBD5E1',
  borderStyle: 'thin',
  alignVertical: 'top',
};

const headerCell = (value: string): CellObject => ({ value, ...HEADER_STYLE });
const labelCell = (value: string): CellObject => ({ value, ...LABEL_STYLE });
const dataCell = (value: string | number | boolean): CellObject => ({
  value,
  ...DATA_STYLE,
  wrap: typeof value === 'string',
});
const textCell = (value: string): CellObject => ({
  value,
  type: String,
  format: '@',
  ...DATA_STYLE,
  wrap: true,
});

const sanitizeFilenamePart = (value: string): string =>
  value
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/gi, (character) => (character === 'Đ' ? 'D' : 'd'))
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 60);

const countAttachments = (result: OneOfficeProjectTaskCollectionResult): number =>
  result.discussion?.comments.reduce(
    (total, comment) => total + comment.attachments.length,
    0,
  ) ?? 0;

const collectedTasks = (result: OneOfficeProjectBatchResult) =>
  result.tasks.filter((task) => task.discussion !== null);

const readNoParts = (value: string): string[] =>
  value.split('.').map((part) => part.trim()).filter(Boolean);

const normalizeNo = (value: string): string => readNoParts(value).join('.');

const formatList = (values: string[] | undefined): string =>
  Array.from(new Set((values ?? []).map((value) => value.trim()).filter(Boolean))).join('; ');

const formatBoolean = (value: boolean): string => (value ? 'Có' : 'Không');

const formatCollectionState = (
  state: OneOfficeProjectTaskCollectionResult['state'],
): string => {
  const labels: Record<OneOfficeProjectTaskCollectionResult['state'], string> = {
    complete: 'Hoàn thành',
    incomplete: 'Chưa đầy đủ',
    failed: 'Thất bại',
    unresolved: 'Chưa lấy được thảo luận',
    cancelled: 'Đã hủy',
  };

  return labels[state];
};

const createProjectSheet = (result: OneOfficeProjectBatchResult, filter?: DiscussionFilter): SheetData => {
  const collected = collectedTasks(result);
  const rootCount = collected.reduce(
    (total, task) => total + (task.discussion?.collection.loadedRootCommentCount ?? 0),
    0,
  );
  const replyCount = collected.reduce(
    (total, task) => total + (task.discussion?.collection.loadedReplyCount ?? 0),
    0,
  );
  const attachmentCount = collected.reduce(
    (total, task) => total + countAttachments(task),
    0,
  );
  const isReport = result.project.type === 'report';
  const rows: Array<[string, string | number | boolean]> = [
    ['ID dự án', isReport ? '' : result.project.internalId],
    ['Mã dự án', result.project.code ?? ''],
    ['Tên dự án', isReport ? '' : result.project.name],
    ['Đường dẫn dự án', isReport ? '' : result.project.url],
    ['Bắt đầu thu thập', result.startedAt],
    ['Hoàn tất thu thập', result.completedAt],
    ['Đã hủy thu thập', formatBoolean(result.cancelled)],
    ['Tổng số công việc', result.inventory.tasks.length],
    ['Công việc đã thu thập', collected.length],
    ['Công việc thất bại', result.tasks.filter((task) => task.state === 'failed').length],
    ['Công việc chưa đầy đủ', result.tasks.filter((task) => task.state === 'incomplete' || task.state === 'unresolved').length],
    ['Công việc đã hủy', result.tasks.filter((task) => task.state === 'cancelled').length],
    ['Bình luận gốc', rootCount],
    ['Phản hồi', replyCount],
    ['Tệp đính kèm', attachmentCount],
    [
      'Cảnh báo',
      Array.from(
        new Set([
          ...result.inventory.collection.warnings,
          ...result.warnings,
          ...collected.flatMap((task) => task.discussion?.collection.warnings ?? []),
        ]),
      ).join('\n'),
    ],
  ];
  if (isReport) rows.push(['Loại nguồn', 'Danh sách trong báo cáo đã lọc'], ['ID báo cáo', result.project.internalId],
    ['Tên báo cáo', result.project.name], ['Đường dẫn báo cáo và bộ lọc', result.project.url]);
  if (filter) {
    rows.push(['Bộ lọc thảo luận', describeDiscussionFilter(filter)],
      ['Phạm vi bộ lọc', 'Chỉ áp dụng sheet Thảo luận và Tệp đính kèm. Sheet Công việc và số liệu thu thập gốc không thay đổi.'],
      ['Thảo luận khớp bộ lọc', collected.reduce((sum, task) => sum + filterDiscussionComments(task.discussion!.comments, filter).length, 0)],
      ['Thời gian không xác định', 'Bản ghi không có thời gian hợp lệ bị loại khi lọc ngày hoặc chọn mới nhất.']);
  }

  return [
    [headerCell('Trường thông tin'), headerCell('Giá trị')],
    ...rows.map(([label, value]) => [labelCell(label), dataCell(value)]),
  ];
};

const createTasksSheet = (result: OneOfficeProjectBatchResult): SheetData => {
  const headers = [
    'TT',
    'Mã công việc',
    'Tên công việc',
    'No.',
    'No. công việc cha',
    'Tên công việc cha',
    'No. công việc cấp 3',
    'Tên công việc cấp 3',
    'Cấp công việc',
    'ID công việc cha',
    'ID công việc',
    'Người thực hiện',
    'Người giao việc',
    'Phòng ban thực hiện',
    'Trạng thái (của công việc)',
    'Bắt đầu (KH gốc)',
    'Kết thúc (KH gốc)',
    'Bắt đầu dự kiến',
    'Kết thúc dự kiến',
    'Bắt đầu thực tế',
    'Kết thúc thực tế',
    'Quá hạn (ngày)',
    'Số ngày hoàn thành quá hạn',
    'Nhãn (nếu có)',
    'Trạng thái thu thập',
    'Bình luận gốc',
    'Phản hồi',
    'Tệp đính kèm',
    'Lỗi',
  ];
  const inventoryTaskByNo = new Map(
    result.inventory.tasks
      .filter((task) => task.no)
      .map((task) => [normalizeNo(task.no), task]),
  );
  const rows = [...result.tasks]
    .sort((first, second) => first.task.sourceOrder - second.task.sourceOrder)
    .map((taskResult, rowIndex) => {
      const noParts = result.project.type === 'report' ? [] : readNoParts(taskResult.task.no);
      const parentNo = noParts.length > 1 ? noParts.slice(0, -1).join('.') : '';
      const levelThreeNo = noParts.length >= 3 ? noParts.slice(0, 3).join('.') : '';
      const parentTask = parentNo ? inventoryTaskByNo.get(parentNo) : null;
      const levelThreeTask = levelThreeNo ? inventoryTaskByNo.get(levelThreeNo) : null;

      return [
        dataCell(rowIndex + 1),
        textCell(taskResult.task.code ?? ''),
        textCell(taskResult.task.name),
        textCell(taskResult.task.no),
        textCell(parentNo),
        textCell(result.project.type === 'report' ? taskResult.task.parentName ?? '' : parentTask?.name ?? ''),
        textCell(levelThreeNo),
        textCell(levelThreeTask?.name ?? ''),
        dataCell(noParts.length),
        textCell(parentTask?.internalId ?? ''),
        textCell(taskResult.task.internalId),
        textCell(formatList(taskResult.task.assigneeNames)),
        textCell(formatList(taskResult.task.assignerNames)),
        textCell(formatList(taskResult.task.departmentNames)),
        textCell(taskResult.task.taskStatus ?? ''),
        textCell(taskResult.task.baselineStart ?? ''),
        textCell(taskResult.task.baselineEnd ?? ''),
        textCell(taskResult.task.plannedStart ?? ''),
        textCell(taskResult.task.plannedEnd ?? ''),
        textCell(taskResult.task.actualStart ?? ''),
        textCell(taskResult.task.actualEnd ?? ''),
        textCell(taskResult.task.overdueDays ?? ''),
        textCell(taskResult.task.completedOverdueDays ?? ''),
        textCell(formatList(taskResult.task.labels)),
        dataCell(formatCollectionState(taskResult.state)),
        dataCell(taskResult.discussion?.collection.loadedRootCommentCount ?? 0),
        dataCell(taskResult.discussion?.collection.loadedReplyCount ?? 0),
        dataCell(countAttachments(taskResult)),
        textCell(taskResult.error ?? ''),
      ];
    });

  return [headers.map(headerCell), ...rows];
};

const createDiscussionSheet = (result: OneOfficeProjectBatchResult, filter?: DiscussionFilter): SheetData => {
  const headers = [
    'STT',
    'No.',
    'No. công việc cha',
    'Tên công việc cha',
    'No. công việc cấp 3',
    'Tên công việc cấp 3',
    'Mã công việc',
    'Tên công việc',
    'Loại thảo luận',
    'Nội dung bình luận cha',
    'Cấp phản hồi',
    'Người bình luận',
    'Thời gian bình luận',
    'Nội dung bình luận',
    'Người được nhắc đến',
    'Số tệp đính kèm',
    'Tên tệp đính kèm',
  ];
  const rows: CellObject[][] = [];
  let rowNumber = 1;
  const inventoryTaskByNo = new Map(
    result.inventory.tasks
      .filter((task) => task.no)
      .map((task) => [normalizeNo(task.no), task]),
  );

  for (const taskResult of [...result.tasks].sort(
    (first, second) => first.task.sourceOrder - second.task.sourceOrder,
  )) {
    const discussion = taskResult.discussion;
    if (!discussion) {
      continue;
    }

    const comments = [...discussion.comments].sort(
      (first, second) => first.sourceOrder - second.sourceOrder,
    );
    const commentById = new Map(comments.map((comment) => [comment.id, comment]));
    const noParts = result.project.type === 'report' ? [] : readNoParts(taskResult.task.no);
    const parentNo = noParts.length > 1 ? noParts.slice(0, -1).join('.') : '';
    const levelThreeNo = noParts.length >= 3 ? noParts.slice(0, 3).join('.') : '';
    const parentTask = parentNo ? inventoryTaskByNo.get(parentNo) : null;
    const levelThreeTask = levelThreeNo ? inventoryTaskByNo.get(levelThreeNo) : null;

    for (const comment of filter ? filterDiscussionComments(comments, filter) : comments) {
      const parent = comment.parentId ? commentById.get(comment.parentId) : null;
      rows.push([
        dataCell(rowNumber++),
        textCell(taskResult.task.no),
        textCell(parentNo),
        textCell(result.project.type === 'report' ? taskResult.task.parentName ?? '' : parentTask?.name ?? ''),
        textCell(levelThreeNo),
        textCell(levelThreeTask?.name ?? ''),
        textCell(taskResult.task.code ?? ''),
        textCell(taskResult.task.name),
        dataCell(comment.parentId ? 'Phản hồi' : 'Bình luận gốc'),
        textCell(parent?.contentText ?? ''),
        dataCell(comment.depth),
        textCell(comment.authorName),
        textCell(comment.createdAtRaw),
        {
          ...textCell(comment.contentText),
          indent: Math.min(comment.depth, 4),
        },
        textCell(comment.mentions.map((mention) => mention.displayName).join(', ')),
        dataCell(comment.attachments.length),
        textCell(comment.attachments.map((attachment) => attachment.name).join('\n')),
      ]);
    }
  }

  return [headers.map(headerCell), ...rows];
};

const createAttachmentRow = (
  result: OneOfficeProjectBatchResult,
  taskResult: OneOfficeProjectTaskCollectionResult,
  comment: OneOfficeComment,
  attachmentIndex: number,
): CellObject[] => {
  const attachment = comment.attachments[attachmentIndex];

  return [
    textCell(result.project.type === 'report' ? taskResult.task.projectInternalId ?? '' : result.project.internalId),
    textCell(result.project.type === 'report' ? taskResult.task.projectName ?? '' : result.project.name),
    textCell(taskResult.task.no),
    textCell(taskResult.task.internalId),
    textCell(taskResult.task.code ?? ''),
    textCell(taskResult.task.name),
    textCell(comment.id),
    textCell(comment.parentId ?? ''),
    textCell(comment.authorName),
    textCell(comment.createdAtRaw),
    dataCell(attachment.index + 1),
    textCell(attachment.name),
    textCell(attachment.sizeText ?? ''),
    dataCell(formatBoolean(attachment.originalDownloadAvailable)),
    textCell(attachment.thumbnailUrl ?? ''),
  ];
};

const createAttachmentsSheet = (result: OneOfficeProjectBatchResult, filter?: DiscussionFilter): SheetData => {
  const headers = [
    'ID dự án',
    'Tên dự án',
    'No.',
    'ID công việc',
    'Mã công việc',
    'Tên công việc',
    'ID bình luận',
    'ID bình luận cha',
    'Người bình luận',
    'Thời gian bình luận',
    'STT tệp',
    'Tên tệp',
    'Dung lượng tệp',
    'Có thể tải file gốc',
    'URL ảnh xem trước',
  ];
  const rows = [...result.tasks]
    .sort((first, second) => first.task.sourceOrder - second.task.sourceOrder)
    .flatMap((taskResult) =>
      (filter ? filterDiscussionComments(taskResult.discussion?.comments ?? [], filter) : taskResult.discussion?.comments ?? []).flatMap((comment) =>
        comment.attachments.map((_, attachmentIndex) =>
          createAttachmentRow(result, taskResult, comment, attachmentIndex),
        ),
      ),
    );

  return [headers.map(headerCell), ...rows];
};

export function createOneOfficeProjectXlsxFilename(
  result: OneOfficeProjectBatchResult,
): string {
  const identity = sanitizeFilenamePart(
    result.project.code || result.project.name || result.project.internalId,
  );
  const date = /^\d{4}-\d{2}-\d{2}/.exec(result.completedAt)?.[0] ?? 'undated';

  return `tera-clip-oneoffice-project-${identity || result.project.internalId}-${date}.xlsx`;
}

export function buildOneOfficeProjectWorkbook(
  result: OneOfficeProjectBatchResult,
  filter?: DiscussionFilter,
): Sheet<Blob>[] {
  return [
    {
      data: createProjectSheet(result, filter),
      sheet: 'Dự án',
      columns: [{ width: 28 }, { width: 80 }],
      stickyRowsCount: 1,
      showGridLines: false,
    },
    {
      data: createTasksSheet(result),
      sheet: 'Công việc',
      columns: [
        { width: 8 },
        { width: 24 },
        { width: 38 },
        { width: 14 },
        { width: 18 },
        { width: 34 },
        { width: 18 },
        { width: 34 },
        { width: 12 },
        { width: 20 },
        { width: 18 },
        { width: 30 },
        { width: 30 },
        { width: 30 },
        { width: 22 },
        { width: 18 },
        { width: 18 },
        { width: 18 },
        { width: 18 },
        { width: 18 },
        { width: 18 },
        { width: 16 },
        { width: 24 },
        { width: 30 },
        { width: 20 },
        { width: 16 },
        { width: 12 },
        { width: 14 },
        { width: 60 },
      ],
      stickyRowsCount: 1,
      stickyColumnsCount: 1,
      showGridLines: false,
    },
    {
      data: createDiscussionSheet(result, filter),
      sheet: 'Thảo luận',
      columns: [
        { width: 10 },
        { width: 14 },
        { width: 18 },
        { width: 34 },
        { width: 18 },
        { width: 34 },
        { width: 24 },
        { width: 38 },
        { width: 16 },
        { width: 60 },
        { width: 12 },
        { width: 24 },
        { width: 24 },
        { width: 72 },
        { width: 32 },
        { width: 18 },
        { width: 44 },
      ],
      stickyRowsCount: 1,
      stickyColumnsCount: 2,
      orientation: 'landscape',
      showGridLines: false,
    },
    {
      data: createAttachmentsSheet(result, filter),
      sheet: 'Tệp đính kèm',
      columns: [
        { width: 18 },
        { width: 30 },
        { width: 14 },
        { width: 18 },
        { width: 24 },
        { width: 38 },
        { width: 16 },
        { width: 20 },
        { width: 24 },
        { width: 24 },
        { width: 18 },
        { width: 44 },
        { width: 18 },
        { width: 28 },
        { width: 60 },
      ],
      stickyRowsCount: 1,
      stickyColumnsCount: 3,
      showGridLines: false,
    },
  ];
}

export async function createOneOfficeProjectXlsxBlob(
  result: OneOfficeProjectBatchResult,
): Promise<Blob> {
  return writeExcelFile(buildOneOfficeProjectWorkbook(result), {
    fontFamily: 'Arial',
    fontSize: 10,
  }).toBlob();
}

export async function downloadOneOfficeProjectXlsx(
  result: OneOfficeProjectBatchResult,
  filter?: DiscussionFilter,
): Promise<string> {
  const filename = createOneOfficeProjectXlsxFilename(result);
  await writeExcelFile(buildOneOfficeProjectWorkbook(result, filter), {
    fontFamily: 'Arial',
    fontSize: 10,
  }).toFile(filename);
  return filename;
}
