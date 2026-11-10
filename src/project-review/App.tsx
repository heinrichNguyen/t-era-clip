import React, { useEffect, useMemo, useRef, useState } from 'react';
import {
  collectOneOfficeProjectTaskDiscussions,
  type OneOfficeProjectBatchProgress,
} from '../oneoffice/projectBatchCollector';
import {
  createOneOfficeProjectXlsxFilename,
  downloadOneOfficeProjectXlsx,
} from '../oneoffice/export/projectXlsx';
import { getOneOfficeProjectBatchSession } from '../oneoffice/projectBatchSession';
import type {
  OneOfficeProjectBatchResult,
  OneOfficeProjectTaskCollectionResult,
  OneOfficeProjectTaskCollectionState,
  OneOfficeProjectTaskInventoryItem,
} from '../oneoffice/types';
import { TadtLogo } from '../shared/TadtLogo';
import { useTemporaryData } from '../oneoffice/temporaryData';
import { ClearTemporaryDataButton, TemporaryDataCleared } from '../shared/TemporaryDataControls';
import { emptyDiscussionFilter, filterDiscussionComments, validateDiscussionFilter, type DiscussionFilter } from '../oneoffice/discussionFilter';
import { projectInventoryWarningLabel } from '../shared/projectInventoryWarning';
import { taskCollectionKey } from '../oneoffice/taskCollectionKey';

interface TaskViewState {
  state: OneOfficeProjectTaskCollectionState;
  result: OneOfficeProjectTaskCollectionResult | null;
}

const stateLabel: Record<OneOfficeProjectTaskCollectionState, string> = {
  pending: 'Chờ xử lý',
  collecting: 'Đang thu thập',
  complete: 'Hoàn thành',
  incomplete: 'Chưa đầy đủ',
  failed: 'Thất bại',
  unresolved: 'Chưa lấy được thảo luận',
  cancelled: 'Đã hủy',
};

const stateColor: Record<OneOfficeProjectTaskCollectionState, string> = {
  pending: 'text-slate-500',
  collecting: 'text-blue-700',
  complete: 'text-emerald-700',
  incomplete: 'text-amber-700',
  failed: 'text-red-700',
  unresolved: 'text-amber-700',
  cancelled: 'text-slate-500',
};

const getErrorMessage = (error: unknown): string =>
  error instanceof Error ? error.message : 'Không thể thu thập thảo luận của dự án.';

export function ProjectReviewApp() {
  const [isLoading, setIsLoading] = useState(true);
  const [errorMessage, setErrorMessage] = useState('');
  const [projectName, setProjectName] = useState('');
  const [sourceOrigin, setSourceOrigin] = useState('');
  const [isReport, setIsReport] = useState(false);
  const [inventoryWarnings, setInventoryWarnings] = useState<string[]>([]);
  const [inventoryComplete, setInventoryComplete] = useState(true);
  const [inventoryTasks, setInventoryTasks] = useState<OneOfficeProjectTaskInventoryItem[]>([]);
  const [taskStates, setTaskStates] = useState<Record<string, TaskViewState>>({});
  const [progress, setProgress] = useState<OneOfficeProjectBatchProgress | null>(null);
  const [result, setResult] = useState<OneOfficeProjectBatchResult | null>(null);
  const [isCancelling, setIsCancelling] = useState(false);
  const [isConfirmingExport, setIsConfirmingExport] = useState(false);
  const [isDownloading, setIsDownloading] = useState(false);
  const [incompleteAcknowledged, setIncompleteAcknowledged] = useState(false);
  const [downloadFeedback, setDownloadFeedback] = useState('');
  const [filter, setFilter] = useState<DiscussionFilter>({ ...emptyDiscussionFilter });
  const startedRef = useRef(false);
  const cancelRef = useRef(false);
  const temporary = useTemporaryData(() => {
    cancelRef.current = true;
    setResult(null); setTaskStates({}); setInventoryTasks([]); setProjectName(''); setSourceOrigin('');
    setProgress(null); setIsConfirmingExport(false); setDownloadFeedback(''); setFilter({ ...emptyDiscussionFilter }); setInventoryWarnings([]); setInventoryComplete(true);
  });

  useEffect(() => {
    if (startedRef.current) {
      return;
    }
    startedRef.current = true;

    void getOneOfficeProjectBatchSession()
      .then(async (session) => {
        if (temporary.clearedRef.current) return;
        if (!session) {
          throw new Error('Không có phiên thu thập công việc. Hãy bắt đầu lại từ cửa sổ T-eraClip.');
        }

        setProjectName(session.inventory.project.name || `Dự án ${session.inventory.project.internalId}`);
        setSourceOrigin(new URL(session.inventory.project.url).origin);
        setIsReport(session.inventory.project.type === 'report');
        setInventoryWarnings(session.inventory.collection.warnings);
        setInventoryComplete(session.inventory.collection.complete);
        setInventoryTasks(session.inventory.tasks);
        setTaskStates(
          Object.fromEntries(
            session.inventory.tasks.map((task) => [
              taskCollectionKey(task),
              { state: 'pending' as const, result: null },
            ]),
          ),
        );
        setIsLoading(false);

        const batchResult = await collectOneOfficeProjectTaskDiscussions(
          session.sourceTab.id,
          session.inventory,
          (nextProgress) => {
            if (temporary.clearedRef.current) return;
            setProgress(nextProgress);
            if (nextProgress.currentTask) {
              setTaskStates((current) => ({
                ...current,
                [taskCollectionKey(nextProgress.currentTask!)]: {
                  state: nextProgress.result?.state ?? 'collecting',
                  result: nextProgress.result,
                },
              }));
            }
          },
          () => cancelRef.current,
        );

        if (temporary.clearedRef.current) return;
        setResult(batchResult);
        setTaskStates(
          Object.fromEntries(
            batchResult.tasks.map((taskResult) => [
              taskCollectionKey(taskResult.task),
              { state: taskResult.state, result: taskResult },
            ]),
          ),
        );

      })
      .catch((error) => {
        setErrorMessage(getErrorMessage(error));
        setIsLoading(false);
      });
  }, []);

  const taskRows = useMemo(() => Object.values(taskStates), [taskStates]);
  const successfulCount = taskRows.filter(
    (task) => task.state === 'complete' || task.state === 'incomplete',
  ).length;
  const failedCount = taskRows.filter((task) => task.state === 'failed').length;
  const unresolvedCount = taskRows.filter((task) => task.state === 'unresolved').length;
  const rootCount = taskRows.reduce(
    (total, task) => total + (task.result?.discussion?.collection.loadedRootCommentCount ?? 0),
    0,
  );
  const replyCount = taskRows.reduce(
    (total, task) => total + (task.result?.discussion?.collection.loadedReplyCount ?? 0),
    0,
  );
  const isFinished = Boolean(result);
  const hasIncompleteTasks = Boolean(
    result?.tasks.some((task) => task.state !== 'complete'),
  );
  const filterError = validateDiscussionFilter(filter);
  const filterSummary = useMemo(() => {
    const tasks = result?.tasks ?? [];
    const comments = tasks.flatMap((task) => task.discussion?.comments ?? []);
    const matches = filterError ? [] : tasks.map((task) => filterDiscussionComments(task.discussion?.comments ?? [], filter));
    return {
      authors: [...new Set(comments.map((comment) => comment.authorName).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'vi')),
      count: matches.reduce((sum, items) => sum + items.length, 0),
      tasks: matches.filter((items) => items.length > 0).length,
      unknownTimes: comments.filter((comment) => !comment.createdAtIso || !Number.isFinite(Date.parse(comment.createdAtIso))).length,
    };
  }, [result, filter, filterError]);

  const handleDownload = async () => {
    if (temporary.clearedRef.current || !result || filterError || (hasIncompleteTasks && !incompleteAcknowledged)) {
      return;
    }

    setIsDownloading(true);
    setDownloadFeedback('');
    try {
      const filename = await downloadOneOfficeProjectXlsx(result, filter);
      setDownloadFeedback(`Đã tải xuống ${filename}`);
      setIsConfirmingExport(false);
    } catch (error) {
      setDownloadFeedback(`Tải xuống thất bại. ${getErrorMessage(error)}`);
    } finally {
      setIsDownloading(false);
    }
  };

  if (temporary.cleared) return <TemporaryDataCleared />;
  if (isLoading) {
    return <main className="mx-auto max-w-6xl p-8 text-sm text-slate-600">Đang chuẩn bị danh sách công việc...</main>;
  }

  if (errorMessage) {
    return (
      <main className="mx-auto max-w-3xl p-8">
        <section className="rounded-2xl border border-red-200 bg-red-50 p-6">
          <h1 className="text-xl font-semibold text-red-800">Không thể thu thập dữ liệu công việc</h1>
          <p className="mt-3 text-sm leading-6 text-red-700">{errorMessage}</p>
          <button type="button" className="mt-3 text-sm font-semibold text-blue-700" onClick={() => window.location.reload()}>Thử lại thu thập</button>
          <ClearTemporaryDataButton onClick={() => void temporary.clearTemporaryData()} disabled={temporary.clearing} error={temporary.clearError} />
        </section>
      </main>
    );
  }

  return (
    <main className="mx-auto max-w-7xl p-6 sm:p-8">
      <ClearTemporaryDataButton onClick={() => void temporary.clearTemporaryData()} disabled={temporary.clearing || !isFinished || isDownloading} error={temporary.clearError} />
      {unresolvedCount > 0 && <p className="mb-4 rounded-xl border border-amber-200 bg-amber-50 p-4 text-sm text-amber-900">{unresolvedCount} công việc chưa lấy được thảo luận do chưa xác minh ID. Các dòng này vẫn được xuất đầy đủ trong sheet Công việc, không bị loại và không được coi là không có bình luận.</p>}
      {(!inventoryComplete || inventoryWarnings.length > 0) && <details open={!inventoryComplete} className="mb-4 rounded-xl border border-amber-200 bg-amber-50 p-4 text-sm text-amber-900"><summary className="cursor-pointer font-semibold">{inventoryComplete ? `Cảnh báo danh sách công việc (${inventoryWarnings.length})` : 'Danh sách công việc chưa đầy đủ'}</summary><p className="mt-2">{inventoryComplete ? 'Danh sách đã được kiểm tra trước khi thu thập. Các cảnh báo gốc được giữ bên dưới.' : 'Kết quả chỉ bao gồm các công việc đã lấy được, không đại diện đầy đủ toàn dự án.'}</p><ul className="mt-2 list-inside list-disc">{inventoryWarnings.map((warning, i) => <li key={i}>{projectInventoryWarningLabel(warning)}</li>)}</ul></details>}
      {isFinished && failedCount > 0 && <button type="button" onClick={() => window.location.reload()}
        className="mb-3 rounded-lg border border-blue-300 px-4 py-2 text-sm font-semibold text-blue-700">Thử lại thu thập</button>}
      <header className="overflow-hidden rounded-3xl bg-gradient-to-br from-blue-900 via-indigo-800 to-violet-700 p-6 text-white shadow-xl shadow-blue-900/15">
        <div className="flex flex-col gap-5 sm:flex-row sm:items-start sm:justify-between">
          <div className="min-w-0">
            <TadtLogo className="h-14 w-24 rounded-lg bg-white object-contain p-1" />
            <p className="mt-5 text-xs font-semibold uppercase tracking-[0.2em] text-cyan-200">
              Thu thập thảo luận danh sách công việc
            </p>
            <h1 className="mt-2 text-3xl font-semibold text-white">{projectName}</h1>
            {sourceOrigin && <p className="mt-2 break-all text-sm text-blue-100">Nguồn: {sourceOrigin}</p>}
            <p className="mt-2 text-sm text-blue-100">
              {isFinished
                ? result?.cancelled
                  ? 'Đã dừng thu thập.'
                  : taskRows.length === 0
                    ? 'Không có công việc để thu thập.'
                    : successfulCount === 0
                      ? 'Đã xử lý danh sách, chưa lấy được thảo luận của công việc nào.'
                      : hasIncompleteTasks
                        ? 'Đã xử lý danh sách, kết quả thảo luận chưa đầy đủ. Xem lý do ở từng công việc.'
                        : 'Đã hoàn tất thu thập.'
                : `Đã xử lý ${progress?.completedCount ?? 0}/${progress?.totalCount ?? taskRows.length} công việc.`}
            </p>
          </div>

          {!isFinished ? (
            <button
              className="rounded-md border border-red-800 px-4 py-2 text-sm font-semibold text-red-200 hover:bg-red-950/40 disabled:opacity-50"
              disabled={isCancelling}
              onClick={() => {
                cancelRef.current = true;
                setIsCancelling(true);
              }}
              type="button"
            >
              {isCancelling ? 'Đang dừng sau nhóm hiện tại...' : 'Dừng thu thập'}
            </button>
          ) : (
            <button
              className="rounded-md bg-cyan-500 px-4 py-2 text-sm font-semibold text-slate-950 transition hover:bg-cyan-400"
              disabled={Boolean(filterError)}
              onClick={() => {
                setIncompleteAcknowledged(false);
                setDownloadFeedback('');
                setIsConfirmingExport(true);
              }}
              type="button"
            >
              Xuất Excel
            </button>
          )}
        </div>

        <div className="mt-6 grid gap-3 sm:grid-cols-4">
          <div className="rounded-xl border border-white/15 bg-white/10 p-4 backdrop-blur-sm">
            <p className="text-2xl font-semibold text-white">{progress?.completedCount ?? 0}/{taskRows.length}</p>
            <p className="mt-1 text-xs text-blue-100">Công việc đã xử lý</p>
          </div>
          <div className="rounded-xl border border-white/15 bg-white/10 p-4 backdrop-blur-sm">
            <p className="text-2xl font-semibold text-emerald-200">{successfulCount}</p>
            <p className="mt-1 text-xs text-blue-100">Thu thập thành công</p>
          </div>
          <div className="rounded-xl border border-white/15 bg-white/10 p-4 backdrop-blur-sm">
            <p className="text-2xl font-semibold text-red-200">{failedCount + unresolvedCount}</p>
            <p className="mt-1 text-xs text-blue-100">{unresolvedCount ? 'Chưa lấy được thảo luận' : 'Thất bại'}</p>
          </div>
          <div className="rounded-xl border border-white/15 bg-white/10 p-4 backdrop-blur-sm">
            <p className="text-lg font-semibold text-white">{rootCount} + {replyCount}</p>
            <p className="mt-1 text-xs text-blue-100">Bình luận gốc + phản hồi</p>
          </div>
        </div>

        <div className="mt-5 h-2 overflow-hidden rounded-full bg-white/20">
          <div
            className="h-full bg-cyan-300 transition-all"
            style={{
              width: `${taskRows.length ? ((progress?.completedCount ?? 0) / taskRows.length) * 100 : 0}%`,
            }}
          />
        </div>
        <p className="mt-3 text-xs leading-5 text-blue-100">
          {isReport
            ? 'Giữ trang này và tab báo cáo luôn mở, không đổi bộ lọc. Hệ thống gửi request ngầm theo nhóm nhỏ, không mở lần lượt trang công việc.'
            : 'Giữ trang này và tab nguồn luôn mở. Hệ thống xử lý theo nhóm nhỏ; chỉ các thảo luận cần phương án dự phòng mới được mở trên tab nguồn.'}
        </p>
      </header>

      {downloadFeedback && (
        <p className="mt-4 rounded-xl border border-emerald-200 bg-emerald-50 px-4 py-3 text-sm text-emerald-700">
          {downloadFeedback}
        </p>
      )}

      {result && result.warnings.length > 0 && (
        <section className="mt-5 rounded-2xl border border-amber-200 bg-amber-50 p-5">
          <h2 className="text-sm font-semibold text-amber-800">Cảnh báo khi thu thập</h2>
          <ul className="mt-3 space-y-2 text-sm text-amber-700">
            {result.warnings.map((warning) => (
              <li key={warning}>- {warning}</li>
            ))}
          </ul>
        </section>
      )}

      {result && (
        <section className="mt-5 rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
          <h2 className="text-lg font-semibold text-blue-900">Bộ lọc thảo luận khi xuất Excel</h2>
          <div className="mt-4 grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
            <label className="text-sm text-slate-700">Từ ngày
              <input aria-label="Từ ngày" className="mt-1 block w-full rounded-md border border-slate-300 p-2" type="date" value={filter.from} onChange={(event) => setFilter({ ...filter, from: event.target.value })} />
            </label>
            <label className="text-sm text-slate-700">Đến ngày
              <input aria-label="Đến ngày" className="mt-1 block w-full rounded-md border border-slate-300 p-2" type="date" value={filter.to} onChange={(event) => setFilter({ ...filter, to: event.target.value })} />
            </label>
            <label className="text-sm text-slate-700">Người bình luận
              <select className="mt-1 block w-full rounded-md border border-slate-300 p-2" value={filter.author} onChange={(event) => setFilter({ ...filter, author: event.target.value })}>
                <option value="">Tất cả</option>
                {filterSummary.authors.map((author) => <option key={author} value={author}>{author}</option>)}
              </select>
            </label>
            <label className="text-sm text-slate-700">Loại thảo luận
              <select className="mt-1 block w-full rounded-md border border-slate-300 p-2" value={filter.kind} onChange={(event) => setFilter({ ...filter, kind: event.target.value as DiscussionFilter['kind'] })}>
                <option value="all">Tất cả</option><option value="roots">Bình luận gốc</option><option value="replies">Phản hồi</option>
              </select>
            </label>
          </div>
          <label className="mt-4 flex items-center gap-2 text-sm text-slate-700">
            <input type="checkbox" checked={filter.latestOnly} onChange={(event) => setFilter({ ...filter, latestOnly: event.target.checked })} />
            Chỉ lấy thảo luận mới nhất trong phần khớp bộ lọc của mỗi công việc
          </label>
          <p className="mt-3 text-sm font-semibold text-blue-800">{filterSummary.count} bình luận/phản hồi thuộc {filterSummary.tasks} công việc khớp bộ lọc.</p>
          <p className="mt-2 text-xs leading-5 text-slate-600">Ngày tính theo giờ Việt Nam (UTC+7), gồm cả ngày kết thúc. Bộ lọc chỉ áp dụng sheet Thảo luận và Tệp đính kèm; giữ toàn bộ sheet Công việc và dữ liệu gốc. Nếu trùng tên, lựa chọn người bình luận bao gồm tất cả bản ghi có tên đó. Các bản ghi mới nhất cùng thời điểm đều được giữ.</p>
          {filterSummary.unknownTimes > 0 && <p className="mt-2 text-sm text-amber-700">{filterSummary.unknownTimes} bản ghi không xác định được thời gian, sẽ bị loại khi lọc ngày hoặc chọn mới nhất.</p>}
          {filterError && <p role="alert" className="mt-2 text-sm text-red-700">{filterError}</p>}
          <button className="mt-3 text-sm font-semibold text-blue-700" type="button" onClick={() => setFilter({ ...emptyDiscussionFilter })}>Xóa bộ lọc</button>
        </section>
      )}

      <section className="mt-5 overflow-hidden rounded-2xl border border-slate-200 bg-white shadow-sm">
        <div className="border-b border-slate-200 p-5">
          <h2 className="text-lg font-semibold text-slate-900">Trạng thái thu thập theo công việc</h2>
          <p className="mt-1 text-xs text-slate-500">
            Danh sách chi tiết được hiển thị tại đây để theo dõi đúng phạm vi công việc đã chọn.
          </p>
        </div>

        <div className="overflow-x-auto">
          <table className="min-w-full divide-y divide-slate-200 text-left text-sm">
            <thead className="bg-slate-50 text-xs uppercase tracking-wide text-slate-500">
              <tr>
                <th className="px-4 py-3">No.</th>
                <th className="px-4 py-3">Mã công việc</th>
                <th className="px-4 py-3">Tên công việc</th>
                <th className="px-4 py-3">Trạng thái</th>
                <th className="px-4 py-3">Thảo luận</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {Object.entries(taskStates).map(([taskId, taskView]) => {
                const sourceTask =
                  taskView.result?.task ??
                  inventoryTasks.find((item) => taskCollectionKey(item) === taskId) ??
                  null;
                const discussion = taskView.result?.discussion;

                return (
                  <tr key={taskId}>
                    <td className="whitespace-nowrap px-4 py-3 font-semibold text-blue-700">{sourceTask?.no ?? ''}</td>
                    <td className="whitespace-nowrap px-4 py-3 text-slate-600">{sourceTask?.code ?? ''}</td>
                    <td className="min-w-64 px-4 py-3 text-slate-800">{sourceTask?.name ?? `Công việc ${taskId}`}</td>
                    <td className={`whitespace-nowrap px-4 py-3 font-medium ${stateColor[taskView.state]}`}>
                      {stateLabel[taskView.state]}
                    </td>
                    <td className="px-4 py-3 text-xs text-slate-500">
                      {taskView.result?.error
                        ? taskView.result.error
                        : discussion
                          ? `${discussion.collection.loadedRootCommentCount} bình luận, ${discussion.collection.loadedReplyCount} phản hồi`
                          : ''}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </section>

      {isConfirmingExport && result && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/60 p-4">
          <section
            aria-labelledby="confirm-project-export-title"
            aria-modal="true"
            className="w-full max-w-lg rounded-2xl border border-slate-200 bg-white p-6 shadow-2xl"
            role="dialog"
          >
            <h2
              className="text-lg font-semibold text-slate-900"
              id="confirm-project-export-title"
            >
              Xác nhận xuất Excel danh sách công việc
            </h2>
            <p className="mt-3 break-all text-sm leading-6 text-slate-700">
              {createOneOfficeProjectXlsxFilename(result)}
            </p>
            <p className="mt-2 text-xs leading-5 text-slate-500">
              File gồm các sheet Dự án, Công việc, Thảo luận và Tệp đính kèm. Dữ liệu gồm No., mã và tên công việc, cấu trúc phản hồi, tác giả, thời gian, người được nhắc đến và thông tin tệp. File đính kèm gốc chưa được tải kèm.
            </p>
            <p className="mt-3 text-sm font-semibold text-blue-800">Xuất {filterSummary.count} bình luận/phản hồi thuộc {filterSummary.tasks} công việc theo bộ lọc. Sheet Công việc giữ nguyên toàn bộ danh sách.</p>

            {hasIncompleteTasks && (
              <label className="mt-4 flex items-start gap-3 rounded-md border border-amber-200 bg-amber-50 p-3 text-sm text-amber-800">
                <input
                  checked={incompleteAcknowledged}
                  className="mt-1"
                  onChange={(event) => setIncompleteAcknowledged(event.target.checked)}
                  type="checkbox"
                />
                Tôi hiểu rằng thảo luận của một số công việc chưa lấy được, chưa đầy đủ hoặc đã bị hủy. Tôi vẫn muốn xuất dữ liệu hiện có; sheet Công việc giữ toàn bộ danh sách.
              </label>
            )}

            <div className="mt-6 flex justify-end gap-3">
              <button
                className="rounded-md border border-slate-300 px-4 py-2 text-sm text-slate-700 hover:bg-slate-50"
                disabled={isDownloading}
                onClick={() => setIsConfirmingExport(false)}
                type="button"
              >
                Hủy
              </button>
              <button
                className="rounded-md bg-blue-700 px-4 py-2 text-sm font-semibold text-white hover:bg-blue-800 disabled:cursor-not-allowed disabled:opacity-50"
                disabled={isDownloading || Boolean(filterError) || (hasIncompleteTasks && !incompleteAcknowledged)}
                onClick={() => {
                  void handleDownload();
                }}
                type="button"
              >
                {isDownloading ? 'Đang chuẩn bị...' : 'Tải file Excel'}
              </button>
            </div>
          </section>
        </div>
      )}
    </main>
  );
}
