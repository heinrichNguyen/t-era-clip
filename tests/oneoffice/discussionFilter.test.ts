import { describe, expect, it } from 'vitest';
import { emptyDiscussionFilter, filterDiscussionComments, validateDiscussionFilter } from '../../src/oneoffice/discussionFilter';
import type { OneOfficeComment } from '../../src/oneoffice/types';

const comment = (id: string, date: string | null, parentId: string | null = null): OneOfficeComment => ({
  id, parentId, rootId: parentId ?? id, depth: parentId ? 1 : 0, sourceOrder: Number(id),
  authorId: id, authorName: parentId ? 'B' : 'A', createdAtRaw: 'source time', createdAtIso: date,
  contentText: id, mentions: [], attachments: [],
});
const comments = [
  comment('1', '2026-09-29T16:59:59Z'),
  comment('2', '2026-09-29T17:00:00Z'),
  comment('3', '2026-09-30T16:59:59Z', '2'),
  comment('4', '2026-09-30T17:00:00Z'),
  comment('5', null),
];

describe('discussion export filter', () => {
  it('keeps all original records by default without mutating them', () => {
    expect(filterDiscussionComments(comments, emptyDiscussionFilter)).toEqual(comments);
    expect(comments).toHaveLength(5);
  });
  it('uses inclusive Vietnamese days and excludes unknown timestamps only when needed', () => {
    expect(filterDiscussionComments(comments, { ...emptyDiscussionFilter, from: '2026-09-30', to: '2026-09-30' }).map((item) => item.id)).toEqual(['2', '3']);
    expect(filterDiscussionComments(comments, { ...emptyDiscussionFilter, to: '2026-09-30' }).map((item) => item.id)).toEqual(['1', '2', '3']);
  });
  it('combines author and reply type; latest means latest within matches', () => {
    expect(filterDiscussionComments(comments, { ...emptyDiscussionFilter, author: 'B', kind: 'replies', latestOnly: true }).map((item) => item.id)).toEqual(['3']);
    expect(filterDiscussionComments(comments, { ...emptyDiscussionFilter, author: 'A', latestOnly: true }).map((item) => item.id)).toEqual(['4']);
  });
  it('keeps timestamp ties and does not invent a latest record with unknown time', () => {
    expect(filterDiscussionComments([comments[0], { ...comments[0], id: 'tie' }], { ...emptyDiscussionFilter, latestOnly: true })).toHaveLength(2);
    expect(filterDiscussionComments([comments[4]], { ...emptyDiscussionFilter, latestOnly: true })).toEqual([]);
  });
  it('rejects reversed or invalid date ranges', () => {
    for (const filter of [
      { ...emptyDiscussionFilter, from: '2026-10-01', to: '2026-09-30' },
      { ...emptyDiscussionFilter, from: '2026-02-30' },
      { ...emptyDiscussionFilter, from: 'invalid' },
    ]) {
      expect(validateDiscussionFilter(filter)).not.toBe('');
      expect(() => filterDiscussionComments(comments, filter)).toThrow();
    }
  });
});
