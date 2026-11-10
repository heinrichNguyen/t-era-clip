import writeExcelFile, {
  type CellObject,
  type Sheet,
  type SheetData,
} from 'write-excel-file/browser';
import type {
  OneOfficeAttachment,
  OneOfficeComment,
  OneOfficeDiscussionExport,
} from '../types';

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

const sanitizeFilenamePart = (value: string): string =>
  value
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/gi, (character) => (character === 'Đ' ? 'D' : 'd'))
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 60);

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

const countAttachments = (comments: OneOfficeComment[]): number =>
  comments.reduce((total, comment) => total + comment.attachments.length, 0);

const formatMentions = (comment: OneOfficeComment): string =>
  comment.mentions.map((mention) => mention.displayName).join(', ');

const formatAttachmentNames = (comment: OneOfficeComment): string =>
  comment.attachments.map((attachment) => attachment.name).join('\n');

const formatBoolean = (value: boolean): string => (value ? 'Có' : 'Không');
const formatEntityType = (value: OneOfficeDiscussionExport['entity']['type']): string =>
  value === 'project' ? 'Dự án' : 'Công việc';

const createEntitySheet = (result: OneOfficeDiscussionExport): SheetData => {
  const comments = [...result.comments].sort(
    (first, second) => first.sourceOrder - second.sourceOrder,
  );
  const rows: Array<[string, string | number | boolean]> = [
    ['Loại đối tượng', formatEntityType(result.entity.type)],
    ['ID nội bộ', result.entity.internalId],
    ['Mã công việc/dự án', result.entity.code ?? ''],
    ['Tên công việc/dự án', result.entity.name],
    ['Đường dẫn nguồn', result.entity.url],
    ['Thời điểm thu thập', result.capturedAt],
    ['Thu thập đầy đủ', formatBoolean(result.collection.complete)],
    ['Số bình luận gốc dự kiến', result.collection.expectedRootCommentCount ?? 'Không xác định'],
    ['Bình luận gốc đã tải', result.collection.loadedRootCommentCount],
    ['Phản hồi đã tải', result.collection.loadedReplyCount],
    ['Tổng số nội dung thảo luận', comments.length],
    ['Tệp đính kèm', countAttachments(comments)],
    ['Số lần tải thêm', result.collection.loadMoreClicks],
    ['Cảnh báo', result.collection.warnings.join('\n')],
  ];

  return [
    [headerCell('Trường thông tin'), headerCell('Giá trị')],
    ...rows.map(([label, value]) => [labelCell(label), dataCell(value)]),
  ];
};

const createDiscussionSheet = (result: OneOfficeDiscussionExport): SheetData => {
  const headers = [
    'STT',
    'Loại thảo luận',
    'ID bình luận',
    'ID bình luận cha',
    'ID bình luận gốc',
    'Cấp phản hồi',
    'Người bình luận',
    'ID người bình luận',
    'Thời gian bình luận',
    'Thời gian ISO',
    'Nội dung bình luận',
    'Người được nhắc đến',
    'Số tệp đính kèm',
    'Tên tệp đính kèm',
    'Loại đối tượng',
    'ID công việc/dự án',
    'Mã công việc/dự án',
    'Tên công việc/dự án',
  ];
  const comments = [...result.comments].sort(
    (first, second) => first.sourceOrder - second.sourceOrder,
  );

  return [
    headers.map(headerCell),
    ...comments.map((comment) => [
      dataCell(comment.sourceOrder + 1),
      dataCell(comment.parentId ? 'Phản hồi' : 'Bình luận gốc'),
      textCell(comment.id),
      textCell(comment.parentId ?? ''),
      textCell(comment.rootId),
      dataCell(comment.depth),
      dataCell(comment.authorName),
      textCell(comment.authorId ?? ''),
      dataCell(comment.createdAtRaw),
      dataCell(comment.createdAtIso ?? ''),
      {
        ...dataCell(comment.contentText),
        indent: Math.min(comment.depth, 4),
      },
      dataCell(formatMentions(comment)),
      dataCell(comment.attachments.length),
      dataCell(formatAttachmentNames(comment)),
      dataCell(formatEntityType(result.entity.type)),
      textCell(result.entity.internalId),
      textCell(result.entity.code ?? ''),
      dataCell(result.entity.name),
    ]),
  ];
};

const createAttachmentRow = (
  result: OneOfficeDiscussionExport,
  comment: OneOfficeComment,
  attachment: OneOfficeAttachment,
) => [
  dataCell(formatEntityType(result.entity.type)),
  textCell(result.entity.internalId),
  textCell(result.entity.code ?? ''),
  dataCell(result.entity.name),
  textCell(comment.id),
  textCell(comment.parentId ?? ''),
  dataCell(comment.authorName),
  dataCell(comment.createdAtRaw),
  dataCell(attachment.index + 1),
  dataCell(attachment.name),
  dataCell(attachment.sizeText ?? ''),
  dataCell(formatBoolean(attachment.originalDownloadAvailable)),
  dataCell(attachment.thumbnailUrl ?? ''),
];

const createAttachmentsSheet = (result: OneOfficeDiscussionExport): SheetData => {
  const headers = [
    'Loại đối tượng',
    'ID công việc/dự án',
    'Mã công việc/dự án',
    'Tên công việc/dự án',
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
  const rows = [...result.comments]
    .sort((first, second) => first.sourceOrder - second.sourceOrder)
    .flatMap((comment) =>
      comment.attachments.map((attachment) =>
        createAttachmentRow(result, comment, attachment),
      ),
    );

  return [headers.map(headerCell), ...rows];
};

export function createOneOfficeXlsxFilename(result: OneOfficeDiscussionExport): string {
  const identity = sanitizeFilenamePart(
    result.entity.code || result.entity.name || result.entity.internalId,
  );
  const date = /^\d{4}-\d{2}-\d{2}/.exec(result.capturedAt)?.[0] ?? 'undated';

  return `tera-clip-oneoffice-${result.entity.type}-${identity || result.entity.internalId}-${date}.xlsx`;
}

export function buildOneOfficeWorkbook(result: OneOfficeDiscussionExport): Sheet<Blob>[] {
  return [
    {
      data: createEntitySheet(result),
      sheet: 'Đối tượng',
      columns: [{ width: 28 }, { width: 80 }],
      stickyRowsCount: 1,
      showGridLines: false,
    },
    {
      data: createDiscussionSheet(result),
      sheet: 'Thảo luận',
      columns: [
        { width: 12 },
        { width: 16 },
        { width: 16 },
        { width: 20 },
        { width: 18 },
        { width: 8 },
        { width: 24 },
        { width: 14 },
        { width: 24 },
        { width: 28 },
        { width: 72 },
        { width: 30 },
        { width: 16 },
        { width: 40 },
        { width: 14 },
        { width: 20 },
        { width: 22 },
        { width: 34 },
      ],
      stickyRowsCount: 1,
      stickyColumnsCount: 2,
      orientation: 'landscape',
      showGridLines: false,
    },
    {
      data: createAttachmentsSheet(result),
      sheet: 'Tệp đính kèm',
      columns: [
        { width: 14 },
        { width: 20 },
        { width: 22 },
        { width: 34 },
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
      showGridLines: false,
    },
  ];
}

export async function createOneOfficeXlsxBlob(
  result: OneOfficeDiscussionExport,
): Promise<Blob> {
  return writeExcelFile(buildOneOfficeWorkbook(result), {
    fontFamily: 'Arial',
    fontSize: 10,
  }).toBlob();
}

export async function downloadOneOfficeXlsx(
  result: OneOfficeDiscussionExport,
): Promise<string> {
  const filename = createOneOfficeXlsxFilename(result);
  await writeExcelFile(buildOneOfficeWorkbook(result), {
    fontFamily: 'Arial',
    fontSize: 10,
  }).toFile(filename);
  return filename;
}
