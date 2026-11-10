import { describe, expect, it } from 'vitest';
import { strFromU8, unzipSync } from 'fflate';
import {
  buildOneOfficeProjectWorkbook,
  createOneOfficeProjectXlsxBlob,
  createOneOfficeProjectXlsxFilename,
} from '../../src/oneoffice/export/projectXlsx';
import type { OneOfficeProjectBatchResult } from '../../src/oneoffice/types';
import { emptyDiscussionFilter } from '../../src/oneoffice/discussionFilter';

const result: OneOfficeProjectBatchResult = {
  schemaVersion: 1,
  source: '1office-project-batch-dom',
  startedAt: '2026-09-28T08:00:00.000Z',
  completedAt: '2026-09-28T08:00:03.000Z',
  cancelled: false,
  project: {
    type: 'project',
    internalId: '830',
    code: 'TEST',
    name: 'TeraClip Parser Test',
    url: 'https://office.meygroup.vn/apps/work-project-project/view?ID=830',
  },
  inventory: {
    schemaVersion: 1,
    source: '1office-project-dom',
    capturedAt: '2026-09-28T08:00:00.000Z',
    project: {
      type: 'project',
      internalId: '830',
      code: 'TEST',
      name: 'TeraClip Parser Test',
      url: 'https://office.meygroup.vn/apps/work-project-project/view?ID=830',
    },
    collection: {
      complete: true,
      expectedTaskCount: 4,
      loadedTaskCount: 4,
      warnings: [],
    },
    tasks: [
      {
        sourceOrder: 0,
        internalId: '91634',
        no: '1',
        code: 'TEST.TDT',
        name: 'Công việc cấp 1',
        url: 'https://office.meygroup.vn/apps/work-task-task/quickview?ID=91634',
        depth: 1,
        parentNo: null,
        parentInternalId: null,
      },
      {
        sourceOrder: 1,
        internalId: '91635',
        no: '1.1',
        code: 'TEST.TDT.01',
        name: 'Công việc cấp 2',
        url: 'https://office.meygroup.vn/apps/work-task-task/quickview?ID=91635',
        depth: 2,
        parentNo: '1',
        parentInternalId: '91634',
      },
      {
        sourceOrder: 2,
        internalId: '91636',
        no: '1.1.1',
        code: 'TEST.TDT.01.01',
        name: 'Công việc cấp 3',
        url: 'https://office.meygroup.vn/apps/work-task-task/quickview?ID=91636',
        depth: 3,
        parentNo: '1.1',
        parentInternalId: '91635',
      },
      {
        sourceOrder: 3,
        internalId: '91637',
        no: '1.1.1.2',
        code: 'TEST.TDT.01.01.02',
        name: 'TEST_SUB',
        url: 'https://office.meygroup.vn/apps/work-task-task/quickview?ID=91637',
        depth: 4,
        parentNo: '1.1.1',
        parentInternalId: '91636',
      },
    ],
  },
  tasks: [
    {
      task: {
        sourceOrder: 3,
        internalId: '91637',
        no: '1.1.1.2',
        code: 'TEST.TDT.01.01.02',
        name: 'TEST_SUB',
        url: 'https://office.meygroup.vn/apps/work-task-task/quickview?ID=91637',
        depth: 4,
        parentNo: '1.1.1',
        parentInternalId: '91636',
        assigneeNames: ['Nguyễn Văn A', 'Nguyễn Văn B'],
        assignerNames: ['Trần Thị C', 'Trần Thị D'],
        departmentNames: ['Ban A', 'Ban B'],
        taskStatus: 'Đang thực hiện',
        baselineStart: '29/09/2026',
        baselineEnd: '08/10/2026',
        plannedStart: '30/09/2026',
        plannedEnd: '09/10/2026',
        actualStart: '01/10/2026',
        actualEnd: '',
        overdueDays: '2',
        completedOverdueDays: '',
        labels: ['Khẩn cấp', 'Thi công'],
      },
      state: 'complete',
      error: null,
      discussion: {
        schemaVersion: 1,
        source: '1office-comment-api',
        capturedAt: '2026-09-28T08:00:01.000Z',
        entity: {
          type: 'task',
          internalId: '91637',
          code: 'TEST.TDT.01',
          name: 'TEST_SUB',
          url: 'https://office.meygroup.vn/apps/work-task-task/view?ID=91637',
        },
        collection: {
          complete: true,
          loadMoreClicks: 0,
          loadedRootCommentCount: 1,
          loadedReplyCount: 1,
          expectedRootCommentCount: 1,
          warnings: [],
        },
        comments: [
          {
            id: '9398',
            parentId: null,
            rootId: '9398',
            depth: 0,
            sourceOrder: 0,
            authorId: '988',
            authorName: 'Nguyễn Anh Tuân',
            createdAtRaw: '10:00:15 28/09/2026',
            createdAtIso: '2026-09-28T10:00:15+07:00',
            contentText: '=Nội dung comment cha',
            mentions: [{ userId: null, displayName: 'Nguyễn Văn Nam' }],
            attachments: [
              {
                index: 0,
                name: 'Hồ sơ.pdf',
                sizeText: '345.67 KB',
                thumbnailUrl: 'https://office.meygroup.vn/files/hoso.png',
                originalDownloadAvailable: true,
              },
            ],
          },
          {
            id: '9401',
            parentId: '9398',
            rootId: '9398',
            depth: 1,
            sourceOrder: 1,
            authorId: '989',
            authorName: 'Nguyễn Văn Nam',
            createdAtRaw: '10:00:26 28/09/2026',
            createdAtIso: '2026-09-28T10:00:26+07:00',
            contentText: 'Đồng ý',
            mentions: [],
            attachments: [],
          },
        ],
      },
    },
  ],
  warnings: [],
};

const readCellValue = (cell: unknown): unknown =>
  cell && typeof cell === 'object' && 'value' in cell
    ? (cell as { value?: unknown }).value
    : cell;

const readBlobAsArrayBuffer = (blob: Blob): Promise<ArrayBuffer> =>
  new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(reader.error ?? new Error('Unable to read XLSX blob.'));
    reader.onload = () => resolve(reader.result as ArrayBuffer);
    reader.readAsArrayBuffer(blob);
  });

describe('project-wide XLSX export', () => {
  it('retains all existing sheets and task/discussion columns for reports without inventing project hierarchy', () => {
    const reportResult = structuredClone(result);
    reportResult.project = { type: 'report', internalId: '74', code: null, name: 'Báo cáo phòng ban',
      url: 'https://office.meygroup.vn/apps/report-report/view?ID=74&department_ids[]=424' };
    reportResult.inventory.project = reportResult.project;
    reportResult.inventory.source = '1office-report-dom';
    reportResult.tasks.forEach(task => { task.task.projectName = 'Dự án khác'; task.task.projectInternalId = '130'; task.task.parentName = 'Cha theo báo cáo'; });
    const original = buildOneOfficeProjectWorkbook(result);
    const workbook = buildOneOfficeProjectWorkbook(reportResult);
    expect(workbook.map(sheet => sheet.sheet)).toEqual(original.map(sheet => sheet.sheet));
    for (let i = 1; i < 4; i++) expect(workbook[i].data[0]).toEqual(original[i].data[0]);
    expect(workbook[1].data[1].slice(4, 10).map(readCellValue)).toEqual(['', 'Cha theo báo cáo', '', '', 0, '']);
    const overview = new Map(workbook[0].data.slice(1).map(row => [readCellValue(row[0]), readCellValue(row[1])]));
    expect(overview.get('ID dự án')).toBe(''); expect(overview.get('ID báo cáo')).toBe('74');
    expect(overview.get('Đường dẫn báo cáo và bộ lọc')).toContain('department_ids[]=424');
    expect(workbook[3].data[1].slice(0, 2).map(readCellValue)).toEqual(['130', 'Dự án khác']);
  });
  it('filters discussions and attachments but preserves tasks, source data and parent context', () => {
    const original = structuredClone(result);
    const workbook = buildOneOfficeProjectWorkbook(result, { ...emptyDiscussionFilter, kind: 'replies' });
    expect(workbook[1].data).toEqual(buildOneOfficeProjectWorkbook(result)[1].data);
    expect(workbook[2].data).toHaveLength(2);
    expect(workbook[2].data[1][9]).toEqual(buildOneOfficeProjectWorkbook(result)[2].data[2][9]);
    expect(readCellValue(workbook[2].data[1][9])).toBe(result.tasks[0].discussion!.comments[0].contentText);
    expect(workbook[3].data).toHaveLength(1);
    expect(workbook[0].data.some((row) => readCellValue(row[0]) === 'Bộ lọc thảo luận')).toBe(true);
    expect(result).toEqual(original);
  });
  it('builds project, task, discussion, and attachment sheets', () => {
    const workbook = buildOneOfficeProjectWorkbook(result);

    expect(workbook.map((sheet) => sheet.sheet)).toEqual([
      'Dự án',
      'Công việc',
      'Thảo luận',
      'Tệp đính kèm',
    ]);
    expect(workbook[1].data[0].map(readCellValue)).toEqual([
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
    ]);
    expect(workbook[1].data[1].map(readCellValue)).toEqual([
      1,
      'TEST.TDT.01.01.02',
      'TEST_SUB',
      '1.1.1.2',
      '1.1.1',
      'Công việc cấp 3',
      '1.1.1',
      'Công việc cấp 3',
      4,
      '91636',
      '91637',
      'Nguyễn Văn A; Nguyễn Văn B',
      'Trần Thị C; Trần Thị D',
      'Ban A; Ban B',
      'Đang thực hiện',
      '29/09/2026',
      '08/10/2026',
      '30/09/2026',
      '09/10/2026',
      '01/10/2026',
      '',
      '2',
      '',
      'Khẩn cấp; Thi công',
      'Hoàn thành',
      1,
      1,
      1,
      '',
    ]);
    expect(workbook[2].data[0].map(readCellValue)).toEqual([
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
    ]);
    expect(workbook[2].data[2].map(readCellValue)).toMatchObject({
      1: '1.1.1.2',
      2: '1.1.1',
      3: 'Công việc cấp 3',
      4: '1.1.1',
      5: 'Công việc cấp 3',
      8: 'Phản hồi',
      9: '=Nội dung comment cha',
      13: 'Đồng ý',
    });
    expect(workbook[3].data[1].map(readCellValue)).toMatchObject({
      2: '1.1.1.2',
      4: 'TEST.TDT.01.01.02',
      11: 'Hồ sơ.pdf',
      12: '345.67 KB',
    });
  });

  it('creates a stable filename and a valid XLSX without formulas from comment text', async () => {
    expect(createOneOfficeProjectXlsxFilename(result)).toBe(
      'tera-clip-oneoffice-project-test-2026-09-28.xlsx',
    );

    const blob = await createOneOfficeProjectXlsxBlob(result);
    const files = unzipSync(new Uint8Array(await readBlobAsArrayBuffer(blob)));
    const workbookXml = strFromU8(files['xl/workbook.xml']);
    const discussionXml = strFromU8(files['xl/worksheets/sheet3.xml']);
    const sharedStringsXml = strFromU8(files['xl/sharedStrings.xml']);

    expect(workbookXml).toContain('name="Dự án"');
    expect(workbookXml).toContain('name="Công việc"');
    expect(workbookXml).toContain('name="Thảo luận"');
    expect(workbookXml).toContain('name="Tệp đính kèm"');
    expect(discussionXml).not.toContain('<f>');
    expect(sharedStringsXml).toContain('Nội dung comment cha');
  });
});
