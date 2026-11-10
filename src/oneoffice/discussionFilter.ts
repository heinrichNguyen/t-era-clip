import type { OneOfficeComment } from './types';

export interface DiscussionFilter {
  from: string;
  to: string;
  author: string;
  kind: 'all' | 'roots' | 'replies';
  latestOnly: boolean;
}

export const emptyDiscussionFilter: DiscussionFilter = {
  from: '', to: '', author: '', kind: 'all', latestOnly: false,
};

export function validateDiscussionFilter(filter: DiscussionFilter): string {
  for (const date of [filter.from, filter.to]) {
    if (date && (!/^\d{4}-\d{2}-\d{2}$/.test(date)
      || !Number.isFinite(Date.parse(`${date}T00:00:00+07:00`))
      || new Date(`${date}T00:00:00Z`).toISOString().slice(0, 10) !== date)) {
      return 'Ngày lọc không hợp lệ.';
    }
  }
  return filter.from && filter.to && filter.from > filter.to
    ? 'Ngày bắt đầu phải trước hoặc bằng ngày kết thúc.' : '';
}

const timestamp = (comment: OneOfficeComment): number => {
  // Only use normalized source timestamps, never guess relative/raw display times.
  return comment.createdAtIso ? Date.parse(comment.createdAtIso) : NaN;
};

export function filterDiscussionComments(
  comments: OneOfficeComment[], filter: DiscussionFilter,
): OneOfficeComment[] {
  const error = validateDiscussionFilter(filter);
  if (error) throw new Error(error);
  const start = filter.from ? Date.parse(`${filter.from}T00:00:00+07:00`) : -Infinity;
  const end = filter.to ? Date.parse(`${filter.to}T00:00:00+07:00`) + 86400000 : Infinity;
  const matches = comments.filter((comment) => {
    const time = timestamp(comment);
    return (!filter.author || comment.authorName === filter.author)
      && (filter.kind === 'all' || (filter.kind === 'roots' ? !comment.parentId : Boolean(comment.parentId)))
      && (!(filter.from || filter.to) || (Number.isFinite(time) && time >= start && time < end));
  });
  if (!filter.latestOnly) return matches;
  const dated = matches.filter((comment) => Number.isFinite(timestamp(comment)));
  const newest = Math.max(...dated.map(timestamp));
  // Keep ties rather than arbitrarily discard equally recent source records.
  return dated.filter((comment) => timestamp(comment) === newest);
}

export function describeDiscussionFilter(filter: DiscussionFilter): string {
  return [
    `Từ ngày: ${filter.from || 'Không giới hạn'}`,
    `Đến ngày: ${filter.to || 'Không giới hạn'} (giờ Việt Nam, UTC+7)`,
    `Người bình luận: ${filter.author || 'Tất cả'}`,
    `Loại: ${filter.kind === 'roots' ? 'Bình luận gốc' : filter.kind === 'replies' ? 'Phản hồi' : 'Tất cả'}`,
    `Mới nhất trong phần khớp bộ lọc của mỗi công việc: ${filter.latestOnly ? 'Có (giữ các bản ghi cùng thời điểm)' : 'Không'}`,
  ].join('\n');
}
