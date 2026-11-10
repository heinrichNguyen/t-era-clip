import React, { useEffect, useMemo, useState } from 'react';
import {
  createOneOfficeJsonFilename,
  downloadOneOfficeJson,
} from '../oneoffice/export/json';
import {
  createOneOfficeXlsxFilename,
  downloadOneOfficeXlsx,
} from '../oneoffice/export/xlsx';
import { getOneOfficeReview } from '../oneoffice/reviewSession';
import type { OneOfficeComment, OneOfficeDiscussionExport } from '../oneoffice/types';
import { TadtLogo } from '../shared/TadtLogo';
import { useTemporaryData } from '../oneoffice/temporaryData';
import { ClearTemporaryDataButton, TemporaryDataCleared } from '../shared/TemporaryDataControls';

const DUPLICATE_DIAGNOSTIC_PATTERN = /^Ignored \d+ duplicate comment ID\(s\)/;
type ExportFormat = 'json' | 'xlsx';

const formatTimestamp = (comment: OneOfficeComment): string => {
  if (!comment.createdAtIso) {
    return comment.createdAtRaw || 'Không rõ thời gian';
  }

  const parsed = new Date(comment.createdAtIso);
  if (Number.isNaN(parsed.getTime())) {
    return comment.createdAtRaw;
  }

  return parsed.toLocaleString('vi-VN', {
    dateStyle: 'medium',
    timeStyle: 'medium',
  });
};

const getErrorMessage = (error: unknown): string =>
  error instanceof Error ? error.message : 'Không thể tải dữ liệu xem trước tạm thời.';

export function ReviewApp() {
  const [result, setResult] = useState<OneOfficeDiscussionExport | null>(null);
  const [isLoading, setIsLoading] = useState(true);
  const [errorMessage, setErrorMessage] = useState('');
  const [isConfirming, setIsConfirming] = useState(false);
  const [exportFormat, setExportFormat] = useState<ExportFormat>('xlsx');
  const [isDownloading, setIsDownloading] = useState(false);
  const [incompleteAcknowledged, setIncompleteAcknowledged] = useState(false);
  const [downloadFeedback, setDownloadFeedback] = useState('');
  const temporary = useTemporaryData(() => {
    setResult(null); setIsConfirming(false); setDownloadFeedback(''); setErrorMessage('');
  });

  useEffect(() => {
    void getOneOfficeReview()
      .then((reviewResult) => {
        if (temporary.clearedRef.current) return;
        if (!reviewResult) {
          setErrorMessage('Không có dữ liệu xem trước 1Office. Hãy thu thập lại thảo luận.');
          return;
        }
        setResult(reviewResult);
      })
      .catch((error) => setErrorMessage(getErrorMessage(error)))
      .finally(() => setIsLoading(false));
  }, []);

  const comments = useMemo(
    () => [...(result?.comments ?? [])].sort((first, second) => first.sourceOrder - second.sourceOrder),
    [result],
  );
  const attachmentCount = useMemo(
    () => comments.reduce((total, comment) => total + comment.attachments.length, 0),
    [comments],
  );
  const actionableWarnings =
    result?.collection.warnings.filter(
      (warning) => !DUPLICATE_DIAGNOSTIC_PATTERN.test(warning),
    ) ?? [];
  const diagnostics =
    result?.collection.warnings.filter((warning) => DUPLICATE_DIAGNOSTIC_PATTERN.test(warning)) ?? [];

  const openExportConfirmation = (format: ExportFormat) => {
    setExportFormat(format);
    setIncompleteAcknowledged(false);
    setDownloadFeedback('');
    setIsConfirming(true);
  };

  const handleDownload = async () => {
    if (temporary.clearedRef.current || !result || (!result.collection.complete && !incompleteAcknowledged)) {
      return;
    }

    setIsDownloading(true);
    try {
      const filename =
        exportFormat === 'xlsx'
          ? await downloadOneOfficeXlsx(result)
          : downloadOneOfficeJson(result);
      setDownloadFeedback(`Đã tải xuống ${filename}`);
      setIsConfirming(false);
    } catch (error) {
      setDownloadFeedback(`Tải xuống thất bại. ${getErrorMessage(error)}`);
    } finally {
      setIsDownloading(false);
    }
  };

  if (temporary.cleared) return <TemporaryDataCleared />;
  if (isLoading) {
    return <main className="mx-auto max-w-5xl p-8 text-sm text-slate-600">Đang tải bản xem trước...</main>;
  }

  if (errorMessage || !result) {
    return (
      <main className="mx-auto max-w-3xl p-8">
        <section className="rounded-2xl border border-red-200 bg-red-50 p-6">
          <h1 className="text-xl font-semibold text-red-800">Không thể mở bản xem trước</h1>
          <p className="mt-3 text-sm leading-6 text-red-700">{errorMessage}</p>
          <ClearTemporaryDataButton onClick={() => void temporary.clearTemporaryData()} disabled={temporary.clearing} error={temporary.clearError} />
        </section>
      </main>
    );
  }

  const filename =
    exportFormat === 'xlsx'
      ? createOneOfficeXlsxFilename(result)
      : createOneOfficeJsonFilename(result);
  const rootCount = result.collection.loadedRootCommentCount;
  const replyCount = result.collection.loadedReplyCount;

  return (
    <main className="mx-auto max-w-5xl p-6 sm:p-8">
      <ClearTemporaryDataButton onClick={() => void temporary.clearTemporaryData()} disabled={temporary.clearing || isDownloading} error={temporary.clearError} />
      <header className="overflow-hidden rounded-3xl bg-gradient-to-br from-blue-900 via-indigo-800 to-violet-700 p-6 text-white shadow-xl shadow-blue-900/15">
        <div className="flex flex-col gap-5 sm:flex-row sm:items-start sm:justify-between">
          <div className="min-w-0">
            <TadtLogo className="h-14 w-24 rounded-lg bg-white object-contain p-1" />
            <p className="mt-5 text-xs font-semibold uppercase tracking-[0.2em] text-cyan-200">
              Xem trước thảo luận 1Office
            </p>
            <h1 className="mt-2 break-words text-3xl font-semibold text-white">
              {result.entity.name || `${result.entity.type} ${result.entity.internalId}`}
            </h1>
            <p className="mt-1 text-sm text-blue-100">
              {result.entity.code ? `${result.entity.code} · ` : ''}
              {result.entity.type === 'project' ? 'Dự án' : 'Công việc'}
            </p>
          </div>

          <div className="flex flex-wrap gap-3">
            <button
              className="rounded-xl bg-white px-4 py-2.5 text-sm font-semibold text-blue-800 transition hover:bg-blue-50"
              onClick={() => openExportConfirmation('xlsx')}
              type="button"
            >
              Xuất Excel
            </button>
            <button
              className="rounded-xl border border-white/40 bg-white/10 px-4 py-2.5 text-sm font-semibold text-white transition hover:bg-white/20"
              onClick={() => openExportConfirmation('json')}
              type="button"
            >
              Xuất JSON
            </button>
          </div>
        </div>

        <div className="mt-6 grid gap-3 sm:grid-cols-4">
          <div className="rounded-xl border border-white/15 bg-white/10 p-4">
            <p className="text-2xl font-semibold text-white">{rootCount}</p>
            <p className="mt-1 text-xs text-blue-100">Bình luận gốc</p>
          </div>
          <div className="rounded-xl border border-white/15 bg-white/10 p-4">
            <p className="text-2xl font-semibold text-white">{replyCount}</p>
            <p className="mt-1 text-xs text-blue-100">Phản hồi</p>
          </div>
          <div className="rounded-xl border border-white/15 bg-white/10 p-4">
            <p className="text-2xl font-semibold text-white">{attachmentCount}</p>
            <p className="mt-1 text-xs text-blue-100">Tệp đính kèm</p>
          </div>
          <div className="rounded-xl border border-white/15 bg-white/10 p-4">
            <p
              className={`text-sm font-semibold ${
                result.collection.complete ? 'text-emerald-200' : 'text-amber-200'
              }`}
            >
              {result.collection.complete ? 'Đầy đủ' : 'Chưa đầy đủ'}
            </p>
            <p className="mt-2 text-xs text-blue-100">
              {result.collection.loadMoreClicks} lần tải thêm
            </p>
          </div>
        </div>

        <dl className="mt-6 grid gap-3 text-xs sm:grid-cols-2">
          <div>
            <dt className="text-blue-200">Thời điểm thu thập</dt>
            <dd className="mt-1 text-white">{new Date(result.capturedAt).toLocaleString('vi-VN')}</dd>
          </div>
          <div className="min-w-0">
            <dt className="text-blue-200">Nguồn</dt>
            <dd className="mt-1 truncate text-white" title={result.entity.url}>
              {result.entity.url}
            </dd>
          </div>
        </dl>
      </header>

      {actionableWarnings.length > 0 && (
        <section className="mt-5 rounded-2xl border border-amber-200 bg-amber-50 p-5">
          <h2 className="text-sm font-semibold text-amber-800">Cảnh báo cần kiểm tra</h2>
          <ul className="mt-3 space-y-2 text-sm text-amber-700">
            {actionableWarnings.map((warning) => (
              <li key={warning}>- {warning}</li>
            ))}
          </ul>
        </section>
      )}

      {diagnostics.length > 0 && (
        <details className="mt-5 rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
          <summary className="cursor-pointer text-sm font-medium text-slate-700">
            Thông tin kỹ thuật ({diagnostics.length})
          </summary>
          <ul className="mt-3 space-y-2 text-xs text-slate-500">
            {diagnostics.map((diagnostic) => (
              <li key={diagnostic}>- {diagnostic}</li>
            ))}
          </ul>
        </details>
      )}

      <section className="mt-5">
        <div className="flex items-end justify-between gap-3">
          <div>
            <h2 className="text-lg font-semibold text-slate-900">Nội dung thảo luận</h2>
            <p className="mt-1 text-xs text-slate-500">Giữ nguyên thứ tự từ nguồn.</p>
          </div>
          <p className="text-xs text-slate-500">Tổng cộng {comments.length} mục</p>
        </div>

        {comments.length === 0 ? (
          <div className="mt-4 rounded-2xl border border-slate-200 bg-white p-8 text-center shadow-sm">
            <p className="text-sm text-slate-600">Công việc này chưa có thảo luận.</p>
          </div>
        ) : (
          <div className="mt-4 space-y-3">
            {comments.map((comment) => (
              <article
                className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm"
                key={comment.id}
                style={{ marginLeft: `${Math.min(comment.depth, 4) * 24}px` }}
              >
                <div className="flex flex-wrap items-baseline justify-between gap-2">
                  <div className="flex items-center gap-2">
                    <p className="text-sm font-semibold text-blue-800">
                      {comment.authorName || 'Không rõ tác giả'}
                    </p>
                    {comment.depth > 0 && (
                      <span className="rounded bg-blue-50 px-1.5 py-0.5 text-[10px] text-blue-700">
                        Phản hồi {comment.parentId}
                      </span>
                    )}
                  </div>
                  <time className="text-xs text-slate-500">{formatTimestamp(comment)}</time>
                </div>

                {comment.contentText ? (
                  <p className="mt-3 whitespace-pre-wrap text-sm leading-6 text-slate-700">
                    {comment.contentText}
                  </p>
                ) : (
                  <p className="mt-3 text-sm italic text-slate-500">Không có nội dung văn bản</p>
                )}

                {comment.attachments.length > 0 && (
                  <div className="mt-3 border-t border-slate-200 pt-3">
                    <p className="text-[11px] font-medium uppercase tracking-wide text-slate-500">
                      Tệp đính kèm
                    </p>
                    <ul className="mt-2 space-y-1 text-xs text-slate-600">
                      {comment.attachments.map((attachment) => (
                        <li key={`${comment.id}-${attachment.index}`}>
                          {attachment.name || 'Tệp không có tên'}
                          {attachment.sizeText ? ` (${attachment.sizeText})` : ''}
                        </li>
                      ))}
                    </ul>
                  </div>
                )}
              </article>
            ))}
          </div>
        )}
      </section>

      {isConfirming && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/60 p-4">
          <section
            aria-labelledby="confirm-export-title"
            aria-modal="true"
            className="w-full max-w-lg rounded-2xl border border-slate-200 bg-white p-6 shadow-2xl"
            role="dialog"
          >
            <h2 className="text-lg font-semibold text-slate-900" id="confirm-export-title">
              Xác nhận xuất {exportFormat === 'xlsx' ? 'Excel' : 'JSON'}
            </h2>
            <p className="mt-3 break-all text-sm leading-6 text-slate-700">{filename}</p>
            <p className="mt-2 text-xs leading-5 text-slate-500">
              {exportFormat === 'xlsx'
                ? 'File gồm các sheet Đối tượng, Thảo luận và Tệp đính kèm. Dữ liệu gồm nội dung, cấu trúc phản hồi, thời gian, người được nhắc đến và thông tin tệp. File đính kèm gốc chưa được tải kèm.'
                : 'File JSON gồm nội dung, tác giả, thời gian, người được nhắc đến, thông tin tệp, cấu trúc phản hồi, URL nguồn và thông tin thu thập. File đính kèm gốc chưa được tải kèm.'}
            </p>

            {!result.collection.complete && (
              <label className="mt-4 flex items-start gap-3 rounded-md border border-amber-200 bg-amber-50 p-3 text-sm text-amber-800">
                <input
                  checked={incompleteAcknowledged}
                  className="mt-1"
                  onChange={(event) => setIncompleteAcknowledged(event.target.checked)}
                  type="checkbox"
                />
                Tôi hiểu dữ liệu thu thập chưa đầy đủ và vẫn muốn xuất phần dữ liệu hiện có.
              </label>
            )}

            <div className="mt-6 flex justify-end gap-3">
              <button
                className="rounded-md border border-slate-300 px-4 py-2 text-sm text-slate-700 hover:bg-slate-50"
                disabled={isDownloading}
                onClick={() => setIsConfirming(false)}
                type="button"
              >
                Hủy
              </button>
              <button
                className="rounded-md bg-blue-700 px-4 py-2 text-sm font-semibold text-white hover:bg-blue-800 disabled:cursor-not-allowed disabled:opacity-50"
                disabled={
                  isDownloading || (!result.collection.complete && !incompleteAcknowledged)
                }
                onClick={() => {
                  void handleDownload();
                }}
                type="button"
              >
                {isDownloading
                  ? 'Đang chuẩn bị...'
                  : `Tải file ${exportFormat === 'xlsx' ? 'Excel' : 'JSON'}`}
              </button>
            </div>
          </section>
        </div>
      )}

      {downloadFeedback && (
        <p className="fixed bottom-5 right-5 rounded-md border border-emerald-200 bg-emerald-50 px-4 py-3 text-sm text-emerald-700 shadow-xl">
          {downloadFeedback}
        </p>
      )}

    </main>
  );
}
