import { beforeEach, describe, expect, it } from 'vitest';
import { strFromU8, unzipSync } from 'fflate';
import projectFixture from '../fixtures/oneoffice/project-discussion.html?raw';
import {
  createOneOfficeJsonFilename,
  serializeOneOfficeJson,
} from '../../src/oneoffice/export/json';
import {
  buildOneOfficeWorkbook,
  createOneOfficeXlsxBlob,
  createOneOfficeXlsxFilename,
} from '../../src/oneoffice/export/xlsx';
import { parseOneOfficeDiscussion } from '../../src/oneoffice/parser';
import { isOneOfficeDiscussionExport } from '../../src/oneoffice/validation';
import type { OneOfficeDiscussionExport } from '../../src/oneoffice/types';

let validResult: OneOfficeDiscussionExport;

beforeEach(() => {
  document.body.innerHTML = projectFixture;
  validResult = parseOneOfficeDiscussion(document, {
    capturedAt: '2026-09-28T04:00:00.000Z',
    sourceUrl: 'https://office.example.test/project/500',
  });
});

describe('1Office review validation', () => {
  it('accepts a complete normalized discussion result', () => {
    expect(isOneOfficeDiscussionExport(validResult)).toBe(true);
  });

  it('rejects malformed temporary review data', () => {
    expect(isOneOfficeDiscussionExport({ ...validResult, schemaVersion: 2 })).toBe(false);
    expect(
      isOneOfficeDiscussionExport({
        ...validResult,
        comments: [{ ...validResult.comments[0], attachments: 'invalid' }],
      }),
    ).toBe(false);
  });
});

describe('1Office JSON export', () => {
  it('creates a stable filesystem-safe filename', () => {
    expect(createOneOfficeJsonFilename(validResult)).toBe(
      'tera-clip-oneoffice-project-da-mau-2026-09-28.json',
    );
  });

  it('serializes the normalized schema as formatted JSON', () => {
    const serialized = serializeOneOfficeJson(validResult);
    const reparsed = JSON.parse(serialized) as OneOfficeDiscussionExport;

    expect(serialized.endsWith('\n')).toBe(true);
    expect(reparsed).toEqual(validResult);
    expect(reparsed.comments).toHaveLength(32);
  });
});

const readCellValue = (cell: unknown): unknown => {
  if (cell && typeof cell === 'object' && 'value' in cell) {
    return (cell as { value?: unknown }).value;
  }
  return cell;
};

const readBlobAsArrayBuffer = (blob: Blob): Promise<ArrayBuffer> =>
  new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(reader.error ?? new Error('Unable to read XLSX blob.'));
    reader.onload = () => resolve(reader.result as ArrayBuffer);
    reader.readAsArrayBuffer(blob);
  });

describe('1Office XLSX export', () => {
  it('creates a stable filesystem-safe filename', () => {
    expect(createOneOfficeXlsxFilename(validResult)).toBe(
      'tera-clip-oneoffice-project-da-mau-2026-09-28.xlsx',
    );
  });

  it('builds entity, discussion, and attachment sheets from the normalized result', () => {
    const workbook = buildOneOfficeWorkbook(validResult);

    expect(workbook.map((sheet) => sheet.sheet)).toEqual([
      'Đối tượng',
      'Thảo luận',
      'Tệp đính kèm',
    ]);

    const discussion = workbook[1].data;
    const attachments = workbook[2].data;
    expect(discussion).toHaveLength(33);
    expect(attachments).toHaveLength(2);
    expect(discussion[0].map(readCellValue)).toContain('ID bình luận cha');
    expect(discussion[1].map(readCellValue)).toMatchObject({
      1: 'Bình luận gốc',
      2: '1001',
      3: '',
      10: 'Dòng một\nDòng hai',
      12: 1,
    });
    expect(discussion[2].map(readCellValue)).toMatchObject({
      1: 'Phản hồi',
      2: '2001',
      3: '1001',
      11: 'Người dùng A',
    });
    expect(attachments[1].map(readCellValue)).toMatchObject({
      4: '1001',
      9: 'Tài liệu mẫu.pdf',
      10: '128 KB',
    });
  });

  it('generates an XLSX zip payload', async () => {
    const blob = await createOneOfficeXlsxBlob(validResult);
    const bytes = new Uint8Array(await readBlobAsArrayBuffer(blob));
    const signature = bytes.slice(0, 4);
    const files = unzipSync(bytes);
    const workbookXml = strFromU8(files['xl/workbook.xml']);

    expect(blob.type).toBe(
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );
    expect([...signature]).toEqual([0x50, 0x4b, 0x03, 0x04]);
    expect(workbookXml).toContain('name="Đối tượng"');
    expect(workbookXml).toContain('name="Thảo luận"');
    expect(workbookXml).toContain('name="Tệp đính kèm"');
    expect(files['xl/worksheets/sheet2.xml']).toBeDefined();
  });

  it('keeps formula-like discussion text as text', async () => {
    const formulaLikeResult: OneOfficeDiscussionExport = {
      ...validResult,
      comments: [
        {
          ...validResult.comments[0],
          contentText: '=HYPERLINK("https://example.test","Open")',
        },
      ],
    };
    const blob = await createOneOfficeXlsxBlob(formulaLikeResult);
    const files = unzipSync(new Uint8Array(await readBlobAsArrayBuffer(blob)));
    const discussionXml = strFromU8(files['xl/worksheets/sheet2.xml']);
    const sharedStringsXml = strFromU8(files['xl/sharedStrings.xml']);

    expect(discussionXml).not.toContain('<f>');
    expect(sharedStringsXml).toContain('HYPERLINK');
  });
});
