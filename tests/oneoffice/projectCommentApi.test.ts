import { afterEach, describe, expect, it, vi } from 'vitest';
import { fetchOneOfficeCommentsFromPage, parseOneOfficeCommentApiResponse } from '../../src/oneoffice/projectCommentApi';
import type { OneOfficeProjectTaskInventoryItem } from '../../src/oneoffice/types';

const task: OneOfficeProjectTaskInventoryItem = {
  sourceOrder: 1,
  internalId: '91637',
  no: '1.1',
  code: 'TEST.TDT.01',
  name: 'TEST_SUB',
  url: 'https://office.meygroup.vn/apps/work-task-task/quickview?ID=91637',
  depth: 2,
  parentNo: '1',
  parentInternalId: '91634',
};

describe('parseOneOfficeCommentApiResponse', () => {
  afterEach(() => vi.unstubAllGlobals());
  it('requests the normal-work namespace for normal tasks and preserves their source URL', async () => {
    const fetchMock = vi.fn().mockResolvedValue({ ok: true, json: async () => ({ total: 0, items: [] }) });
    vi.stubGlobal('fetch', fetchMock);
    await fetchOneOfficeCommentsFromPage({ allowedOrigin: window.location.origin, taskId: '89876', offset: 1000, object: 'work-normal-normal' });
    expect(new URL(fetchMock.mock.calls[0][0]).searchParams.get('object')).toBe('work-normal-normal');
    const normalTask = { ...task, internalId: '89876', url: 'https://office.meygroup.vn/apps/work-normal-normal/view?ID=89876' };
    expect(parseOneOfficeCommentApiResponse({ total: 0, items: [] }, normalTask)?.entity.url).toBe(normalTask.url);
  });
  it('normalizes roots, replies, mentions, timestamps, and attachments', () => {
    const result = parseOneOfficeCommentApiResponse(
      {
        total: 1,
        items: [
          {
            ID: 9398,
            parent_id: 0,
            dateTime: '10:00:15 28/09/2026',
            content: '@Nguyễn Văn Nam vui lòng bổ sung hồ sơ',
            user: { ID: 988, name: 'Nguyễn Anh Tuân' },
            files: [
              {
                ID: 29300,
                name: 'Hồ sơ.pdf',
                size: '345.67 KB',
                thumb: '/files/hoso.png',
                url: 'https://office.meygroup.vn/files/hoso.pdf',
              },
            ],
            total_subs: 1,
            subs: [
              {
                ID: 9401,
                parent_id: 9398,
                dateTime: '10:00:26 28/09/2026',
                content: 'Nguyễn Anh Tuân&nbsp;đồng ý<div>28/09/2026</div>',
                user: { ID: 989, name: 'Nguyễn Văn Nam' },
                files: [],
              },
            ],
          },
        ],
      },
      task,
    );

    expect(result).toMatchObject({
      source: '1office-comment-api',
      entity: {
        internalId: '91637',
        code: 'TEST.TDT.01',
        name: 'TEST_SUB',
      },
      collection: {
        complete: true,
        loadedRootCommentCount: 1,
        loadedReplyCount: 1,
        expectedRootCommentCount: 1,
      },
    });
    expect(result?.comments[0]).toMatchObject({
      id: '9398',
      parentId: null,
      rootId: '9398',
      authorId: '988',
      createdAtIso: '2026-09-28T10:00:15+07:00',
      mentions: [{ userId: null, displayName: 'Nguyễn Văn Nam' }],
      attachments: [
        {
          name: 'Hồ sơ.pdf',
          sizeText: '345.67 KB',
          originalDownloadAvailable: true,
        },
      ],
    });
    expect(result?.comments[1]).toMatchObject({
      id: '9401',
      parentId: '9398',
      rootId: '9398',
      depth: 1,
      contentText: 'Nguyễn Anh Tuân đồng ý\n28/09/2026',
      mentions: [{ userId: '988', displayName: 'Nguyễn Anh Tuân' }],
    });
  });

  it('rejects incomplete root or reply totals so the caller can fall back', () => {
    expect(
      parseOneOfficeCommentApiResponse({ total: 2, items: [] }, task),
    ).toBeNull();

    expect(
      parseOneOfficeCommentApiResponse(
        {
          total: 1,
          items: [{ ID: 9398, total_subs: 1, subs: [] }],
        },
        task,
      ),
    ).toBeNull();
  });
});
