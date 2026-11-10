import { collectOneOfficeProjectInventory } from './collectionClient';
import { collectOneOfficeTaskFromCommentApi } from './projectCommentApi';
import { oneOfficeOrigin } from './origins';
import type { OneOfficeDiscussionExport, OneOfficeProjectInventory, OneOfficeProjectTaskInventoryItem } from './types';
import { collectTestTaskDetails, type TestTaskDetails } from './testTaskDetails';

import { TEST_BASELINE_ORIGIN, TEST_BASELINE_PROJECT } from './testBaselineScope';
export { TEST_BASELINE_ORIGIN, TEST_BASELINE_PROJECT } from './testBaselineScope';

export interface TestBaselineSample {
  task: OneOfficeProjectTaskInventoryItem;
  reason: string;
  discussion: OneOfficeDiscussionExport | null;
  error: string | null;
  details?: TestTaskDetails | null;
  detailError?: string | null;
}

export interface TestBaseline {
  schemaVersion: 1;
  mode: 'read-only-test-baseline';
  origin: string;
  startedAt: string;
  completedAt: string;
  inventory: OneOfficeProjectInventory;
  samples: TestBaselineSample[];
  limitations: string[];
}

export function selectTestBaselineSamples(inventory: OneOfficeProjectInventory): TestBaselineSample[] {
  const projectUrl = new URL(inventory.project.url);
  if (oneOfficeOrigin(projectUrl.href) !== TEST_BASELINE_ORIGIN || inventory.project.internalId !== TEST_BASELINE_PROJECT ||
      projectUrl.pathname !== '/apps/work-project-project/view' || projectUrl.searchParams.get('ID') !== TEST_BASELINE_PROJECT) {
    throw new Error('Chuẩn bị kiểm thử chỉ dành cho dự án 203 trên site test cổng 9443.');
  }
  if (!inventory.collection.complete || inventory.tasks.length < 5 || inventory.collection.loadedTaskCount !== inventory.tasks.length ||
      inventory.collection.expectedTaskCount !== inventory.tasks.length) {
    throw new Error('Cần danh sách đầy đủ và ít nhất 5 công việc trước khi chọn mẫu.');
  }
  const ids = new Set<string>();
  for (const task of inventory.tasks) {
    const url = new URL(task.url);
    if (!/^\d+$/.test(task.internalId) || ids.has(task.internalId) || oneOfficeOrigin(url.href) !== TEST_BASELINE_ORIGIN ||
        !['/apps/work-task-task/view', '/apps/work-task-task/quickview'].includes(url.pathname) || url.searchParams.get('ID') !== task.internalId) {
      throw new Error('Danh sách có ID trùng, định danh URL không khớp hoặc khác môi trường.');
    }
    ids.add(task.internalId);
  }
  const byId = new Map(inventory.tasks.map((task) => [task.internalId, task]));
  for (const task of inventory.tasks) {
    const chain = new Set<string>();
    let current: OneOfficeProjectTaskInventoryItem | undefined = task;
    while (current) {
      if (chain.has(current.internalId)) throw new Error('Quan hệ cha/con có vòng lặp. Chưa chọn mẫu.');
      chain.add(current.internalId);
      if (current.parentInternalId && !byId.has(current.parentInternalId)) throw new Error('Thiếu công việc cha trong bản nguồn.');
      current = current.parentInternalId ? byId.get(current.parentInternalId) : undefined;
    }
  }
  const parents = new Set(inventory.tasks.map((task) => task.parentInternalId).filter(Boolean));
  const selected: TestBaselineSample[] = [];
  const add = (task: OneOfficeProjectTaskInventoryItem | undefined, reason: string) => {
    if (task && !selected.some((sample) => sample.task.internalId === task.internalId) && selected.length < 5) {
      selected.push({ task, reason, discussion: null, error: null });
    }
  };
  const parent = inventory.tasks.find((task) => parents.has(task.internalId));
  add(parent, 'Công việc cha, đối chiếu quan hệ bằng ID');
  add(inventory.tasks.find((task) => task.parentInternalId === parent?.internalId), 'Công việc con của mẫu cha');
  add([...inventory.tasks].sort((a, b) => b.depth - a.depth).find((task) => !parents.has(task.internalId)), 'Công việc lá ở cấp sâu');
  add(inventory.tasks.find((task) => task.sourceNo === ''), 'No. gốc đang trống, không thay bằng số hiển thị');
  add(inventory.tasks.find((task) => task.taskStatus && task.taskStatus !== parent?.taskStatus), 'Trạng thái khác mẫu cha (không suy ra phần trăm hoàn thành)');
  add(inventory.tasks.find((task) => task.assigneeNames?.length), 'Có người thực hiện trong danh sách nguồn');
  for (const task of inventory.tasks) add(task, 'Bổ sung mẫu trong cùng dự án');
  return selected;
}

export async function collectTestBaseline(tabId: number, progress: (message: string) => void, cancelled: () => boolean): Promise<TestBaseline> {
  const startedAt = new Date().toISOString();
  const check = () => { if (cancelled()) throw new Error('Đã hủy chuẩn bị kiểm thử.'); };
  check(); progress('Đang đọc lại danh sách hiện tại, không ghi...');
  const inventory = await collectOneOfficeProjectInventory(tabId, TEST_BASELINE_ORIGIN);
  check();
  const samples = selectTestBaselineSamples(inventory);
  // Only the five selected samples get detail requests. Never navigate or invoke a save fallback.
  for (const [index, sample] of samples.entries()) {
    check(); progress(`Đang đọc thảo luận mẫu ${index + 1}/5, ID ${sample.task.internalId}...`);
    try {
      sample.discussion = await collectOneOfficeTaskFromCommentApi(tabId, sample.task);
      if (!sample.discussion?.collection.complete) sample.error = 'Chưa đọc đủ thảo luận qua API, cần xem thủ công.';
    } catch (error) {
      sample.error = error instanceof Error ? error.message : 'Chưa đọc được thảo luận.';
    }
    check();
    progress(`Đang đọc hai khối thông tin mã/No. mẫu ${index + 1}/5, ID ${sample.task.internalId}...`);
    try {
      sample.details = await collectTestTaskDetails(tabId, sample.task.internalId);
      sample.detailError = sample.details.error;
    } catch (error) {
      sample.details = null;
      sample.detailError = error instanceof Error ? error.message : 'Chưa đọc được chi tiết mẫu.';
    }
    check();
  }
  return {
    schemaVersion: 1, mode: 'read-only-test-baseline', origin: TEST_BASELINE_ORIGIN, startedAt,
    completedAt: new Date().toISOString(), inventory, samples,
    limitations: [
      'Đây là bản hiện trạng đọc trong nhiều request, không phải ảnh chụp nguyên tử hoặc bản sao lưu đầy đủ.',
      'Chỉ đọc thảo luận của 5 mẫu. Không tìm thấy thảo luận ở các mẫu không chứng minh toàn dự án không có thảo luận.',
      'Không suy ra phần trăm hoàn thành từ trạng thái công việc. Chỉ giữ tiến độ khi khối chi tiết trả về số hợp lệ; chưa thu thập nội dung kết quả báo cáo hoặc toàn bộ tệp đính kèm để đối chiếu bảo toàn.',
      'Chưa phê duyệt mẫu để ghi. Phải xác minh các dữ liệu còn thiếu và đọc lại hiện trạng trước mọi cập nhật.',
      'Mã/No. từ hai khối thông tin của 5 mẫu được ghi riêng, không sửa bản nguồn API. Trường trả về trống khác với chưa đọc được; dấu -- chỉ là ký hiệu hiển thị. Kết quả này chưa xác minh dữ liệu biểu mẫu sửa hoặc phê duyệt ghi.',
    ],
  };
}

export function downloadTestBaseline(baseline: TestBaseline): string {
  const name = `tera-clip-test-203-baseline-${baseline.completedAt.replace(/[:.]/g, '-')}.json`;
  const url = URL.createObjectURL(new Blob([JSON.stringify(baseline, null, 2)], { type: 'application/json;charset=utf-8' }));
  const link = document.createElement('a'); link.href = url; link.download = name; link.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  return name;
}
