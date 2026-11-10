import React, { useCallback, useEffect, useState } from 'react';
import {
  collectOneOfficeDiscussion,
  collectOneOfficeFilteredTaskInventory,
  collectOneOfficeProjectInventory,
  getActiveTab,
  isOneOfficeFilteredTaskListUrl,
  isOneOfficeProjectPageUrl,
  isOneOfficeReportPageUrl,
  isSupportedOneOfficeUrl,
  type ActiveTab
} from '../oneoffice/collectionClient';
import type { OneOfficeCollectionProgress } from '../oneoffice/collector';
import { useTemporaryData } from '../oneoffice/temporaryData';
import { openOneOfficeReview } from '../oneoffice/reviewSession';
import { openOneOfficeProjectBatchReview } from '../oneoffice/projectBatchSession';
import type {
  OneOfficeDiscussionExport,
  OneOfficeProjectInventory
} from '../oneoffice/types';
import {
  generateCompletedList,
  generateDailyReport,
  generateFollowUpList,
  generateWaitingList
} from '../shared/reportGenerator';
import { clearItems, getItems, updateItem } from '../shared/storage';
import {
  generateInternalFollowUpTemplate,
  generateLeadershipSummaryTemplate,
  generateMeetingActionListTemplate,
  generateVendorFollowUpTemplate
} from '../shared/templateGenerator';
import type { TeraClipItem, TeraClipPriority, TeraClipStatus } from '../shared/types';
import { TadtLogo } from '../shared/TadtLogo';
import { isAccessHistoryUrl, openAccessReview } from '../oneoffice/accessHistory';
import { ProjectImportLauncher } from '../shared/ProjectImportLauncher';
import { ProjectDiscussionFlow } from './ProjectDiscussionFlow';
import { collectOneOfficeReportInventory } from '../oneoffice/reportTaskInventory';

const STATUS_OPTIONS: TeraClipStatus[] = ['inbox', 'todo', 'doing', 'waiting', 'done', 'archived'];
const PRIORITY_OPTIONS: TeraClipPriority[] = ['low', 'medium', 'high', 'urgent'];

const STATUS_LABELS: Record<TeraClipStatus, string> = {
  inbox: 'Hộp thư',
  todo: 'Cần làm',
  doing: 'Đang làm',
  waiting: 'Đang chờ',
  done: 'Hoàn thành',
  archived: 'Lưu trữ'
};

const PRIORITY_LABELS: Record<TeraClipPriority, string> = {
  low: 'Thấp',
  medium: 'Trung bình',
  high: 'Cao',
  urgent: 'Khẩn cấp'
};

type StatusFilter = TeraClipStatus | 'all';
type PriorityFilter = TeraClipPriority | 'all';
type ReportType = 'daily' | 'followup' | 'waiting' | 'completed';
type TemplateType = 'vendor' | 'internal' | 'leadership' | 'meeting';

type ItemDraft = {
  title: string;
  notes: string;
};

type ItemDrafts = Record<string, ItemDraft>;

const REPORT_OPTIONS: { label: string; value: ReportType }[] = [
  { label: 'Báo cáo hằng ngày', value: 'daily' },
  { label: 'Danh sách theo dõi', value: 'followup' },
  { label: 'Danh sách đang chờ', value: 'waiting' },
  { label: 'Danh sách hoàn thành', value: 'completed' }
];

const TEMPLATE_OPTIONS: { label: string; value: TemplateType }[] = [
  { label: 'Theo dõi nhà cung cấp', value: 'vendor' },
  { label: 'Theo dõi nội bộ', value: 'internal' },
  { label: 'Tóm tắt cho lãnh đạo', value: 'leadership' },
  { label: 'Đầu việc sau cuộc họp', value: 'meeting' }
];

const sortNewestFirst = (items: TeraClipItem[]): TeraClipItem[] =>
  [...items].sort((first, second) => {
    const firstTime = new Date(first.createdAt).getTime();
    const secondTime = new Date(second.createdAt).getTime();

    return secondTime - firstTime;
  });

const formatCreatedAt = (createdAt: string): string => {
  const date = new Date(createdAt);

  if (Number.isNaN(date.getTime())) {
    return createdAt || 'Không rõ thời gian';
  }

  return date.toLocaleString('vi-VN', {
    month: 'short',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit'
  });
};

const getErrorMessage = (error: unknown): string => {
  if (error instanceof Error) {
    return error.message;
  }

  return 'Không thể tải các nội dung đã lưu.';
};

const getOneOfficeErrorMessage = (error: unknown): string => {
  if (error instanceof Error) {
    return error.message;
  }

  return 'Không thể thu thập thảo luận từ 1Office.';
};

const generateReportText = (reportType: ReportType, items: TeraClipItem[]): string => {
  switch (reportType) {
    case 'followup':
      return generateFollowUpList(items);
    case 'waiting':
      return generateWaitingList(items);
    case 'completed':
      return generateCompletedList(items);
    case 'daily':
    default:
      return generateDailyReport(items);
  }
};

const generateTemplateText = (templateType: TemplateType, items: TeraClipItem[]): string => {
  switch (templateType) {
    case 'internal':
      return generateInternalFollowUpTemplate(items);
    case 'leadership':
      return generateLeadershipSummaryTemplate(items);
    case 'meeting':
      return generateMeetingActionListTemplate(items);
    case 'vendor':
    default:
      return generateVendorFollowUpTemplate(items);
  }
};

export function PopupApp() {
  const [activeView, setActiveView] = useState<'main' | 'capture'>('main');
  const [items, setItems] = useState<TeraClipItem[]>([]);
  const [itemDrafts, setItemDrafts] = useState<ItemDrafts>({});
  const [statusFilter, setStatusFilter] = useState<StatusFilter>('all');
  const [priorityFilter, setPriorityFilter] = useState<PriorityFilter>('all');
  const [reportType, setReportType] = useState<ReportType>('daily');
  const [reportText, setReportText] = useState('');
  const [reportFeedback, setReportFeedback] = useState('');
  const [templateType, setTemplateType] = useState<TemplateType>('vendor');
  const [templateText, setTemplateText] = useState('');
  const [templateFeedback, setTemplateFeedback] = useState('');
  const [isLoading, setIsLoading] = useState(true);
  const [isResettingCapture, setIsResettingCapture] = useState(false);
  const [showCaptureResetConfirmation, setShowCaptureResetConfirmation] = useState(false);
  const [captureResetFeedback, setCaptureResetFeedback] = useState('');
  const [captureResetError, setCaptureResetError] = useState('');
  const [updatingItemIds, setUpdatingItemIds] = useState<Set<string>>(new Set());
  const [errorMessage, setErrorMessage] = useState('');
  const [updateErrorMessage, setUpdateErrorMessage] = useState('');
  const [validationErrors, setValidationErrors] = useState<Record<string, string>>({});
  const [oneOfficeTab, setOneOfficeTab] = useState<ActiveTab | null>(null);
  const [isCollectingOneOffice, setIsCollectingOneOffice] = useState(false);
  const [oneOfficeProgress, setOneOfficeProgress] = useState<OneOfficeCollectionProgress | null>(
    null
  );
  const [oneOfficeResult, setOneOfficeResult] = useState<OneOfficeDiscussionExport | null>(null);
  const [oneOfficeError, setOneOfficeError] = useState('');
  const [isOpeningOneOfficeReview, setIsOpeningOneOfficeReview] = useState(false);
  const [isCollectingProjectInventory, setIsCollectingProjectInventory] = useState(false);
  const [projectInventory, setProjectInventory] = useState<OneOfficeProjectInventory | null>(null);
  const [isOpeningProjectReview, setIsOpeningProjectReview] = useState(false);

  const loadItems = useCallback(async () => {
    setIsLoading(true);
    setErrorMessage('');
    setUpdateErrorMessage('');
    setValidationErrors({});
    setReportFeedback('');
    setTemplateFeedback('');

    try {
      const storedItems = await getItems();
      const sortedItems = sortNewestFirst(storedItems);
      setItems(sortedItems);
      setItemDrafts(
        sortedItems.reduce<ItemDrafts>((drafts, item) => {
          drafts[item.id] = {
            title: item.title,
            notes: item.notes
          };
          return drafts;
        }, {})
      );
    } catch (error) {
      setItems([]);
      setItemDrafts({});
      setErrorMessage(getErrorMessage(error));
    } finally {
      setIsLoading(false);
    }
  }, []);

  const saveItem = async (updatedItem: TeraClipItem, errorContext: string): Promise<boolean> => {
    setUpdateErrorMessage('');
    setUpdatingItemIds((currentIds) => new Set(currentIds).add(updatedItem.id));

    try {
      await updateItem(updatedItem);
      setItems((currentItems) =>
        sortNewestFirst(
          currentItems.map((currentItem) => (currentItem.id === updatedItem.id ? updatedItem : currentItem))
        )
      );
      return true;
    } catch (error) {
      setUpdateErrorMessage(`Không thể cập nhật ${errorContext}. ${getErrorMessage(error)}`);
      return false;
    } finally {
      setUpdatingItemIds((currentIds) => {
        const nextIds = new Set(currentIds);
        nextIds.delete(updatedItem.id);
        return nextIds;
      });
    }
  };

  const handleStatusChange = async (item: TeraClipItem, status: TeraClipStatus) => {
    if (item.status === status) {
      return;
    }

    await saveItem(
      {
        ...item,
        status,
        updatedAt: new Date().toISOString()
      },
      'trạng thái'
    );
  };

  const handlePriorityChange = async (item: TeraClipItem, priority: TeraClipPriority) => {
    if (item.priority === priority) {
      return;
    }

    await saveItem(
      {
        ...item,
        priority,
        updatedAt: new Date().toISOString()
      },
      'mức ưu tiên'
    );
  };

  const handleTitleSave = async (item: TeraClipItem, title: string) => {
    const trimmedTitle = title.trim();

    if (!trimmedTitle) {
      setValidationErrors((currentErrors) => ({
        ...currentErrors,
        [item.id]: 'Tiêu đề không được để trống.'
      }));
      return;
    }

    setValidationErrors((currentErrors) => {
      const nextErrors = { ...currentErrors };
      delete nextErrors[item.id];
      return nextErrors;
    });

    if (item.title === trimmedTitle) {
      setItemDrafts((currentDrafts) => ({
        ...currentDrafts,
        [item.id]: {
          title: item.title,
          notes: currentDrafts[item.id]?.notes ?? item.notes
        }
      }));
      return;
    }

    const wasSaved = await saveItem(
      {
        ...item,
        title: trimmedTitle,
        updatedAt: new Date().toISOString()
      },
      'tiêu đề'
    );

    if (wasSaved) {
      setItemDrafts((currentDrafts) => ({
        ...currentDrafts,
        [item.id]: {
          title: trimmedTitle,
          notes: currentDrafts[item.id]?.notes ?? item.notes
        }
      }));
    }
  };

  const handleNotesSave = async (item: TeraClipItem, notes: string) => {
    if (item.notes === notes) {
      return;
    }

    const wasSaved = await saveItem(
      {
        ...item,
        notes,
        updatedAt: new Date().toISOString()
      },
      'ghi chú'
    );

    if (wasSaved) {
      setItemDrafts((currentDrafts) => ({
        ...currentDrafts,
        [item.id]: {
          title: currentDrafts[item.id]?.title ?? item.title,
          notes
        }
      }));
    }
  };

  const handleGenerateReport = () => {
    setReportText(generateReportText(reportType, items));
    setReportFeedback('');
  };

  const handleResetCapture = async () => {
    if (isLoading || isResettingCapture || updatingItemIds.size > 0) {
      return;
    }

    setIsResettingCapture(true);
    setCaptureResetFeedback('');
    setCaptureResetError('');

    try {
      await clearItems();
      setItems([]);
      setItemDrafts({});
      setStatusFilter('all');
      setPriorityFilter('all');
      setValidationErrors({});
      setErrorMessage('');
      setUpdateErrorMessage('');
      setReportText('');
      setReportFeedback('');
      setTemplateText('');
      setTemplateFeedback('');
      setShowCaptureResetConfirmation(false);
      setCaptureResetFeedback('Đã xóa toàn bộ danh sách Capture.');
    } catch (error) {
      setCaptureResetError(`Không thể reset danh sách. ${getErrorMessage(error)}`);
    } finally {
      setIsResettingCapture(false);
    }
  };

  const handleCopyReport = async () => {
    if (!reportText) {
      return;
    }

    try {
      await navigator.clipboard.writeText(reportText);
      setReportFeedback('Đã sao chép vào bộ nhớ tạm.');
    } catch (error) {
      setReportFeedback(`Không thể sao chép báo cáo. ${getErrorMessage(error)}`);
    }
  };

  const handleGenerateTemplate = () => {
    setTemplateText(generateTemplateText(templateType, items));
    setTemplateFeedback('');
  };

  const handleCopyTemplate = async () => {
    if (!templateText) {
      return;
    }

    try {
      await navigator.clipboard.writeText(templateText);
      setTemplateFeedback('Đã sao chép vào bộ nhớ tạm.');
    } catch (error) {
      setTemplateFeedback(`Không thể sao chép mẫu. ${getErrorMessage(error)}`);
    }
  };

  const handleCollectOneOffice = async () => {
    if (!oneOfficeTab || isCollectingOneOffice || isCollectingProjectInventory || isOpeningProjectReview) {
      return;
    }

    setIsCollectingOneOffice(true);
    temporary.reset();
    const revision = temporary.revisionRef.current;
    setOneOfficeProgress(null);
    setOneOfficeResult(null);
    setOneOfficeError('');

    try {
      const result = await collectOneOfficeDiscussion(oneOfficeTab.id, (progress) => {
        if (revision === temporary.revisionRef.current) setOneOfficeProgress(progress);
      }, new URL(oneOfficeTab.url).origin);
      if (revision !== temporary.revisionRef.current) return;
      setOneOfficeResult(result);
    } catch (error) {
      setOneOfficeError(getOneOfficeErrorMessage(error));
    } finally {
      setIsCollectingOneOffice(false);
    }
  };

  const handleOpenOneOfficeReview = async () => {
    if (!oneOfficeResult || isOpeningOneOfficeReview) {
      return;
    }

    setIsOpeningOneOfficeReview(true);
    setOneOfficeError('');

    try {
      await openOneOfficeReview(oneOfficeResult);
    } catch (error) {
      setOneOfficeError(getOneOfficeErrorMessage(error));
      setIsOpeningOneOfficeReview(false);
    }
  };

  const handleCollectProjectInventory = async (openDiscussions = false) => {
    if (!oneOfficeTab || isCollectingProjectInventory || isCollectingOneOffice || isOpeningProjectReview) {
      return;
    }

    setIsCollectingProjectInventory(true);
    temporary.reset();
    const revision = temporary.revisionRef.current;
    setProjectInventory(null);
    setOneOfficeError('');

    try {
      const inventory = isOneOfficeReportPageUrl(oneOfficeTab.url)
        ? await collectOneOfficeReportInventory(oneOfficeTab.id, new URL(oneOfficeTab.url).origin, oneOfficeTab.url)
        : isOneOfficeFilteredTaskListUrl(oneOfficeTab.url)
        ? await collectOneOfficeFilteredTaskInventory(oneOfficeTab.id, new URL(oneOfficeTab.url).origin)
        : await collectOneOfficeProjectInventory(oneOfficeTab.id, new URL(oneOfficeTab.url).origin);
      if (revision !== temporary.revisionRef.current) return;
      setProjectInventory(inventory);
      // Only a complete inventory can proceed without a second, explicit action.
      if (openDiscussions && inventory.collection.complete && inventory.tasks.length > 0) {
        setIsOpeningProjectReview(true);
        await openOneOfficeProjectBatchReview(oneOfficeTab, inventory);
      }
    } catch (error) {
      if (revision === temporary.revisionRef.current) setOneOfficeError(getOneOfficeErrorMessage(error));
      setIsOpeningProjectReview(false);
    } finally {
      setIsCollectingProjectInventory(false);
    }
  };

  const handleOpenProjectReview = async () => {
    if (!oneOfficeTab || !projectInventory?.tasks.length || isOpeningProjectReview || isCollectingProjectInventory || isCollectingOneOffice) {
      return;
    }

    setIsOpeningProjectReview(true);
    setOneOfficeError('');

    try {
      await openOneOfficeProjectBatchReview(oneOfficeTab, projectInventory);
    } catch (error) {
      setOneOfficeError(getOneOfficeErrorMessage(error));
      setIsOpeningProjectReview(false);
    }
  };

  useEffect(() => {
    void loadItems();
  }, [loadItems]);

  useEffect(() => {
    let isCurrent = true;

    void getActiveTab()
      .then((tab) => {
        if (isCurrent && tab && isSupportedOneOfficeUrl(tab.url)) {
          setOneOfficeTab(tab);
        }
      })
      .catch(() => {
        if (isCurrent) {
          setOneOfficeTab(null);
        }
      });

    return () => {
      isCurrent = false;
    };
  }, []);

  const temporary = useTemporaryData(() => {
    setOneOfficeResult(null); setOneOfficeProgress(null); setProjectInventory(null);
  });
  const isFiltering = statusFilter !== 'all' || priorityFilter !== 'all';
  const filteredItems = items.filter((item) => {
    const matchesStatus = statusFilter === 'all' || item.status === statusFilter;
    const matchesPriority = priorityFilter === 'all' || item.priority === priorityFilter;

    return matchesStatus && matchesPriority;
  });
  const itemCountLabel = isFiltering
    ? `${filteredItems.length}/${items.length} nội dung`
    : `${items.length} nội dung đã lưu`;
  const oneOfficeAttachmentCount =
    oneOfficeResult?.comments.reduce((count, comment) => count + comment.attachments.length, 0) ?? 0;
  const oneOfficeRootCount = oneOfficeResult?.collection.loadedRootCommentCount ?? 0;
  const oneOfficeReplyCount = oneOfficeResult?.collection.loadedReplyCount ?? 0;
  const visibleOneOfficeWarnings =
    oneOfficeResult?.collection.warnings.filter(
      (warning) => !/^Ignored \d+ duplicate comment ID\(s\)/.test(warning)
    ) ?? [];
  const isOneOfficeProjectPage = oneOfficeTab
    ? isOneOfficeProjectPageUrl(oneOfficeTab.url)
    : false;
  const isOneOfficeFilteredTaskList = oneOfficeTab
    ? isOneOfficeFilteredTaskListUrl(oneOfficeTab.url)
    : false;
  const isOneOfficeReportPage = oneOfficeTab ? isOneOfficeReportPageUrl(oneOfficeTab.url) : false;
  const inventoryMaxDepth = projectInventory
    ? Math.max(...projectInventory.tasks.map((task) => task.depth), 0)
    : 0;
  const canOpenProjectReview = (projectInventory?.tasks.length ?? 0) > 0;
  const projectInventoryWarnings = projectInventory?.collection.warnings ?? [];
  const visibleProjectInventoryWarnings = projectInventoryWarnings.slice(0, 3);

  if (activeView === 'main') {
    return (
      <div className="flex h-[570px] w-[400px] flex-col overflow-hidden bg-slate-50 text-slate-900">
        <header className="border-b border-slate-200 bg-white px-5 py-4">
          <div className="flex items-center gap-3">
            <TadtLogo className="h-12 w-16 object-contain object-left" />
            <div className="min-w-0 border-l border-slate-200 pl-3">
              <h1 className="text-lg font-bold tracking-tight text-blue-900">T-eraClip</h1>
              <p className="text-xs text-slate-500">Xuất thảo luận 1Office</p>
            </div>
          </div>
        </header>

        <div className="min-h-0 flex-1 overflow-y-auto p-4">
          <section className="overflow-hidden rounded-2xl bg-gradient-to-br from-blue-800 via-indigo-700 to-violet-700 p-5 text-white shadow-lg shadow-blue-900/15">
            <p className="text-[10px] font-semibold uppercase tracking-[0.22em] text-cyan-200">
              Trang 1Office hiện tại
            </p>
            <h2 className="mt-2 truncate text-lg font-semibold" title={oneOfficeTab?.title}>
              {oneOfficeResult?.entity.name || oneOfficeTab?.title || 'Chưa nhận diện trang hỗ trợ'}
            </h2>
            {oneOfficeTab && <p className="mt-2 break-all text-xs text-slate-500">Môi trường: {new URL(oneOfficeTab.url).host}</p>}
            <p className="mt-2 text-xs leading-5 text-blue-100">
              Thu thập trang đang mở, toàn bộ dự án hoặc danh sách công việc trong báo cáo đã lọc.
            </p>

            {oneOfficeTab ? (
              <div className="mt-4 grid gap-2">
                {isAccessHistoryUrl(oneOfficeTab.url) && <button
                  className="rounded-xl bg-white px-4 py-2.5 text-sm font-semibold text-blue-800"
                  onClick={() => void openAccessReview(oneOfficeTab.id, new URL(oneOfficeTab.url).origin).catch(() => setOneOfficeError('Không thể mở trang báo cáo lịch sử truy cập.'))}
                  type="button">Báo cáo lịch sử truy cập</button>}
                {!isOneOfficeProjectPage && !isOneOfficeFilteredTaskList && !isOneOfficeReportPage && !isAccessHistoryUrl(oneOfficeTab.url) && (
                  <button
                    className="rounded-xl bg-white px-4 py-2.5 text-sm font-semibold text-blue-800 shadow-sm transition hover:bg-blue-50 disabled:cursor-not-allowed disabled:opacity-60"
                    disabled={isCollectingOneOffice}
                    onClick={() => void handleCollectOneOffice()}
                    type="button"
                  >
                    {isCollectingOneOffice ? 'Đang thu thập...' : 'Thu thập thảo luận trang hiện tại'}
                  </button>
                )}
                {isOneOfficeProjectPage && <ProjectDiscussionFlow inventory={projectInventory} checking={isCollectingProjectInventory} opening={isOpeningProjectReview} disabled={isCollectingOneOffice}
                  onCollect={() => { if (projectInventory?.tasks.length) void handleOpenProjectReview(); else void handleCollectProjectInventory(true); }}
                  onCheckOnly={() => void handleCollectProjectInventory()} onCollectPage={() => void handleCollectOneOffice()} />}
                {isOneOfficeReportPage && <>
                  <button type="button" disabled={isCollectingProjectInventory || isOpeningProjectReview || isCollectingOneOffice}
                    onClick={() => void handleCollectProjectInventory(true)}
                    className="rounded-xl bg-white px-4 py-3 text-sm font-bold text-blue-800 disabled:opacity-60">
                    {isCollectingProjectInventory ? 'Đang đọc danh sách báo cáo...' : 'Thu thập thảo luận danh sách trong báo cáo'}
                  </button>
                  <p className="text-xs leading-5 text-blue-100">Dùng danh sách theo phòng ban và thời gian đang lọc. Tiến độ mở ở trang riêng, thu thập bằng request ngầm, không chuyển tab qua từng công việc. Ngày báo cáo không tự lọc ngày bình luận.</p>
                </>}
                {isOneOfficeFilteredTaskList && (
                  <button
                    className="rounded-xl border border-white/40 bg-white/10 px-4 py-2.5 text-sm font-semibold text-white transition hover:bg-white/20 disabled:cursor-not-allowed disabled:opacity-60"
                    disabled={isCollectingProjectInventory || isCollectingOneOffice}
                    onClick={() => void handleCollectProjectInventory()}
                    type="button"
                  >
                    {isCollectingProjectInventory
                      ? 'Đang kiểm tra danh sách công việc...'
                      : isOneOfficeFilteredTaskList
                        ? 'Kiểm tra danh sách công việc đang lọc'
                        : 'Kiểm tra danh sách công việc dự án'}
                  </button>
                )}
              </div>
            ) : (
              <div className="mt-4 rounded-xl border border-white/20 bg-white/10 p-3 text-xs leading-5 text-blue-50">
                Hãy mở một trang dự án hoặc công việc trên 1Office, sau đó mở lại T-eraClip.
              </div>
            )}
          </section>

          {isCollectingOneOffice && (
            <section className="mt-3 rounded-xl border border-blue-200 bg-blue-50 p-4 text-xs text-blue-800">
              Đã tải {oneOfficeProgress?.loadedRootCommentCount ?? 0}
              {oneOfficeProgress?.expectedRootCommentCount
                ? `/${oneOfficeProgress.expectedRootCommentCount}`
                : ''}{' '}
              bình luận gốc.
            </section>
          )}

          {oneOfficeError && (
            <section className="mt-3 rounded-xl border border-red-200 bg-red-50 p-4 text-xs leading-5 text-red-700">
              {oneOfficeError}
            </section>
          )}

          {oneOfficeResult && (
            <section className="mt-3 rounded-2xl border border-slate-200 bg-white p-4 shadow-sm">
              <div className="grid grid-cols-3 gap-2 text-center">
                <div className="rounded-lg bg-blue-50 p-2">
                  <p className="text-lg font-bold text-blue-800">{oneOfficeRootCount}</p>
                  <p className="text-[10px] text-slate-500">Bình luận</p>
                </div>
                <div className="rounded-lg bg-indigo-50 p-2">
                  <p className="text-lg font-bold text-indigo-700">{oneOfficeReplyCount}</p>
                  <p className="text-[10px] text-slate-500">Phản hồi</p>
                </div>
                <div className="rounded-lg bg-violet-50 p-2">
                  <p className="text-lg font-bold text-violet-700">{oneOfficeAttachmentCount}</p>
                  <p className="text-[10px] text-slate-500">Tệp đính kèm</p>
                </div>
              </div>
              <p className={`mt-3 text-xs font-medium ${oneOfficeResult.collection.complete ? 'text-emerald-700' : 'text-amber-700'}`}>
                {oneOfficeResult.collection.complete
                  ? 'Đã thu thập đầy đủ.'
                  : 'Dữ liệu chưa đầy đủ, cần kiểm tra cảnh báo trước khi xuất.'}
              </p>
              {visibleOneOfficeWarnings.length > 0 && (
                <ul className="mt-2 space-y-1 text-[11px] leading-4 text-amber-700">
                  {visibleOneOfficeWarnings.map((warning) => <li key={warning}>• {warning}</li>)}
                </ul>
              )}
              <button
                className="mt-3 w-full rounded-xl bg-blue-700 px-4 py-2.5 text-sm font-semibold text-white transition hover:bg-blue-800 disabled:opacity-60"
                disabled={isOpeningOneOfficeReview}
                onClick={() => void handleOpenOneOfficeReview()}
                type="button"
              >
                {isOpeningOneOfficeReview ? 'Đang mở bản xem trước...' : 'Xem trước và xuất dữ liệu'}
              </button>
            </section>
          )}

          {projectInventory && !isOneOfficeProjectPage && (
            <section className="mt-3 rounded-2xl border border-slate-200 bg-white p-4 shadow-sm">
              <div className="flex items-start justify-between gap-3">
                <div>
                  <p className="text-sm font-semibold text-slate-900">
                    {isOneOfficeReportPage ? 'Danh sách công việc trong báo cáo' : isOneOfficeFilteredTaskList
                      ? 'Danh sách công việc đang lọc'
                      : 'Danh sách công việc dự án'}
                  </p>
                  <p className="mt-1 text-xs text-slate-500">
                    {projectInventory.collection.loadedTaskCount} công việc, tối đa {inventoryMaxDepth} cấp
                  </p>
                </div>
                <span className={`rounded-full px-2 py-1 text-[10px] font-semibold ${projectInventory.collection.complete ? 'bg-emerald-50 text-emerald-700' : 'bg-amber-50 text-amber-700'}`}>
                  {projectInventory.collection.complete
                    ? projectInventoryWarnings.length > 0
                      ? 'Đủ, có cảnh báo'
                      : 'Đầy đủ'
                    : 'Thiếu công việc'}
                </span>
              </div>
              <button
                className="mt-3 w-full rounded-xl bg-gradient-to-r from-cyan-500 to-blue-700 px-4 py-2.5 text-sm font-semibold text-white transition hover:from-cyan-600 hover:to-blue-800 disabled:cursor-not-allowed disabled:opacity-60"
                disabled={isOpeningProjectReview || !canOpenProjectReview}
                onClick={() => void handleOpenProjectReview()}
                type="button"
              >
                {isOpeningProjectReview
                  ? 'Đang mở trang thu thập...'
                  : isOneOfficeFilteredTaskList || isOneOfficeReportPage
                    ? 'Thu thập thảo luận danh sách này'
                    : 'Thu thập toàn bộ thảo luận dự án'}
              </button>
              {visibleProjectInventoryWarnings.length > 0 && (
                <ul className="mt-2 space-y-1 text-[10px] leading-4 text-amber-700">
                  {visibleProjectInventoryWarnings.map((warning) => (
                    <li key={warning}>• {warning}</li>
                  ))}
                  {projectInventoryWarnings.length > visibleProjectInventoryWarnings.length && (
                    <li>
                      • Còn {projectInventoryWarnings.length - visibleProjectInventoryWarnings.length} cảnh báo khác.
                    </li>
                  )}
                </ul>
              )}
              <p className="mt-2 text-[10px] leading-4 text-slate-500">
                Tiến độ và danh sách chi tiết sẽ mở ở một trang riêng.
              </p>
            </section>
          )}

          <button
            className="mt-3 flex w-full items-center justify-between rounded-2xl border border-slate-200 bg-white p-4 text-left shadow-sm transition hover:border-blue-300 hover:shadow-md"
            onClick={() => setActiveView('capture')}
            type="button"
          >
            <span>
              <span className="block text-sm font-semibold text-slate-900">Capture</span>
              <span className="mt-1 block text-xs text-slate-500">Mở các công cụ lưu và quản lý nội dung cũ</span>
            </span>
            <span className="text-lg text-blue-700">→</span>
          </button>
        </div>
        <ProjectImportLauncher inventory={!isOneOfficeFilteredTaskList && isOneOfficeProjectPage ? projectInventory : null} sourceTabId={oneOfficeTab?.id} disabled={isCollectingOneOffice || isCollectingProjectInventory} />
      </div>
    );
  }

  return (
    <div className="capture-theme flex h-[570px] w-[400px] flex-col overflow-hidden bg-slate-950 text-slate-100">
      <div className="capture-header border-b border-slate-800 bg-white p-4">
        <div className="flex items-start justify-between gap-3">
          <div className="flex min-w-0 items-start gap-3">
            <TadtLogo className="h-12 w-16 shrink-0 object-contain object-left" />
            <div className="min-w-0 border-l border-slate-200 pl-3">
              <button
                className="text-xs font-medium text-blue-700 hover:text-blue-900"
                onClick={() => setActiveView('main')}
                type="button"
              >
                ← Menu chính
              </button>
              <h1 className="mt-1 text-xl font-bold text-blue-900">Capture</h1>
              <p className="mt-0.5 text-xs text-slate-500">Lưu nội dung và chuyển thành đầu việc.</p>
            </div>
          </div>

          <button
            className="rounded-md border border-slate-700 px-3 py-1.5 text-xs font-medium text-slate-200 transition hover:border-slate-500 hover:bg-slate-900 disabled:cursor-not-allowed disabled:opacity-60"
            disabled={isLoading || isResettingCapture}
            onClick={() => {
              void loadItems();
            }}
            type="button"
          >
            Làm mới
          </button>
        </div>

        <div className="mt-3 flex flex-wrap items-center gap-2">
          <label className="flex items-center gap-2 text-[11px] font-medium text-slate-400">
            Trạng thái
            <select
              className="rounded-md border border-slate-700 bg-slate-950 px-2 py-1 text-[11px] font-medium capitalize text-slate-100 outline-none transition hover:border-slate-500 focus:border-blue-500"
              onChange={(event) => {
                setStatusFilter(event.target.value as StatusFilter);
              }}
              value={statusFilter}
            >
              <option value="all">Tất cả</option>
              {STATUS_OPTIONS.map((status) => (
                <option key={status} value={status}>
                  {STATUS_LABELS[status]}
                </option>
              ))}
            </select>
          </label>

          <label className="flex items-center gap-2 text-[11px] font-medium text-slate-400">
            Ưu tiên
            <select
              className="rounded-md border border-slate-700 bg-slate-950 px-2 py-1 text-[11px] font-medium capitalize text-slate-100 outline-none transition hover:border-slate-500 focus:border-blue-500"
              onChange={(event) => {
                setPriorityFilter(event.target.value as PriorityFilter);
              }}
              value={priorityFilter}
            >
              <option value="all">Tất cả</option>
              {PRIORITY_OPTIONS.map((priority) => (
                <option key={priority} value={priority}>
                  {PRIORITY_LABELS[priority]}
                </option>
              ))}
            </select>
          </label>

          {isFiltering && (
            <button
              className="rounded-md border border-slate-700 px-2 py-1 text-[11px] font-medium text-slate-200 transition hover:border-slate-500 hover:bg-slate-900"
              onClick={() => {
                setStatusFilter('all');
                setPriorityFilter('all');
              }}
              type="button"
            >
              Xóa bộ lọc
            </button>
          )}
        </div>

        <div className="mt-2 flex items-center justify-between gap-2">
          <p className="text-xs text-slate-500">{itemCountLabel}</p>
          <button
            className="text-xs font-medium text-red-700 hover:text-red-900 disabled:cursor-not-allowed disabled:opacity-50"
            disabled={isLoading || isResettingCapture || updatingItemIds.size > 0 || items.length === 0}
            onClick={() => {
              setCaptureResetFeedback('');
              setCaptureResetError('');
              setShowCaptureResetConfirmation(true);
            }}
            type="button"
          >
            Reset danh sách
          </button>
        </div>
        {showCaptureResetConfirmation && (
          <div className="mt-3 rounded-md border border-red-200 bg-red-50 p-3">
            <p className="text-xs leading-5 text-red-800">
              Xóa toàn bộ {items.length} nội dung Capture, kể cả nội dung đang bị bộ lọc ẩn?
              Thao tác này không thể hoàn tác. Cài đặt và dữ liệu xuất thảo luận 1Office được giữ nguyên.
            </p>
            <div className="mt-2 flex justify-end gap-3">
              <button
                className="text-xs font-medium text-slate-600 disabled:opacity-50"
                disabled={isResettingCapture}
                onClick={() => setShowCaptureResetConfirmation(false)}
                type="button"
              >
                Hủy
              </button>
              <button
                className="rounded-md bg-red-700 px-3 py-1.5 text-xs font-medium text-white disabled:opacity-50"
                disabled={isLoading || isResettingCapture || updatingItemIds.size > 0}
                onClick={() => { void handleResetCapture(); }}
                type="button"
              >
                {isResettingCapture ? 'Đang xóa...' : 'Xóa toàn bộ'}
              </button>
            </div>
          </div>
        )}
        {captureResetFeedback && (
          <p className="mt-2 text-xs text-emerald-700" role="status">{captureResetFeedback}</p>
        )}
        {captureResetError && (
          <p className="mt-2 text-xs text-red-700" role="alert">{captureResetError}</p>
        )}
      </div>

      <fieldset disabled={isResettingCapture} className="min-h-0 min-w-0 flex-1 overflow-y-auto p-4">
        {oneOfficeTab && (
          <section className="hidden">
            <div className="flex items-start justify-between gap-3">
              <div className="min-w-0">
                <h2 className="text-sm font-semibold text-cyan-100">1Office discussion</h2>
                <p className="mt-1 truncate text-[11px] text-cyan-200/70" title={oneOfficeTab.title}>
                  {oneOfficeResult?.entity.name || oneOfficeTab.title || 'Current 1Office page'}
                </p>
              </div>

              <button
                className="shrink-0 rounded-md border border-cyan-700 px-3 py-1.5 text-xs font-medium text-cyan-100 transition hover:border-cyan-500 hover:bg-cyan-900/50 disabled:cursor-not-allowed disabled:opacity-60"
                disabled={isCollectingOneOffice}
                onClick={() => {
                  void handleCollectOneOffice();
                }}
                type="button"
              >
                {isCollectingOneOffice ? 'Collecting...' : 'Collect discussion'}
              </button>
            </div>

            {isOneOfficeProjectPage && (
              <div className="mt-3 border-t border-cyan-900/80 pt-3">
                <button
                  className="w-full rounded-md border border-cyan-700 px-3 py-2 text-xs font-semibold text-cyan-100 transition hover:border-cyan-500 hover:bg-cyan-900/40 disabled:cursor-not-allowed disabled:opacity-60"
                  disabled={isCollectingProjectInventory || isCollectingOneOffice}
                  onClick={() => {
                    void handleCollectProjectInventory();
                  }}
                  type="button"
                >
                  {isCollectingProjectInventory ? 'Checking project tasks...' : 'Check project tasks'}
                </button>
                <p className="mt-1.5 text-[10px] leading-4 text-cyan-200/60">
                  Inventory validation only. Open the project&apos;s Công việc tab first.
                </p>
              </div>
            )}

            {isCollectingOneOffice && (
              <p className="mt-3 text-xs text-cyan-100">
                Loaded {oneOfficeProgress?.loadedRootCommentCount ?? 0}
                {oneOfficeProgress?.expectedRootCommentCount
                  ? `/${oneOfficeProgress.expectedRootCommentCount}`
                  : ''}{' '}
                root comments
                {oneOfficeProgress ? `, ${oneOfficeProgress.loadMoreClicks} load-more clicks` : ''}.
              </p>
            )}

            {oneOfficeError && (
              <div className="mt-3 rounded-md border border-red-900 bg-red-950/40 p-3 text-xs text-red-100">
                {oneOfficeError}
              </div>
            )}

            {oneOfficeResult && (
              <div className="mt-3 rounded-md border border-cyan-900 bg-slate-950/50 p-3">
                <p className="text-xs font-medium text-cyan-100">
                  {oneOfficeRootCount} root {oneOfficeRootCount === 1 ? 'comment' : 'comments'},{' '}
                  {oneOfficeReplyCount} {oneOfficeReplyCount === 1 ? 'reply' : 'replies'},{' '}
                  {oneOfficeAttachmentCount}{' '}
                  {oneOfficeAttachmentCount === 1 ? 'attachment' : 'attachments'}
                </p>
                <p
                  className={`mt-1 text-[11px] ${
                    oneOfficeResult.collection.complete ? 'text-emerald-300' : 'text-amber-300'
                  }`}
                >
                  {oneOfficeResult.collection.complete
                    ? 'Collection complete.'
                    : 'Collection incomplete. Review warnings before export.'}
                </p>
                {visibleOneOfficeWarnings.length > 0 && (
                  <ul className="mt-2 space-y-1 text-[11px] leading-4 text-amber-200">
                    {visibleOneOfficeWarnings.map((warning) => (
                      <li key={warning}>- {warning}</li>
                    ))}
                  </ul>
                )}
                <p className="mt-2 text-[11px] text-slate-500">
                  Temporary validation only. Nothing was saved or downloaded.
                </p>
                <button
                  className="mt-3 w-full rounded-md border border-cyan-700 px-3 py-2 text-xs font-semibold text-cyan-100 transition hover:border-cyan-500 hover:bg-cyan-900/40 disabled:cursor-not-allowed disabled:opacity-60"
                  disabled={isOpeningOneOfficeReview}
                  onClick={() => {
                    void handleOpenOneOfficeReview();
                  }}
                  type="button"
                >
                  {isOpeningOneOfficeReview ? 'Opening preview...' : 'Review and export JSON'}
                </button>
              </div>
            )}

            {projectInventory && (
              <div className="mt-3 rounded-md border border-cyan-900 bg-slate-950/50 p-3">
                <p className="text-xs font-medium text-cyan-100">
                  {projectInventory.collection.loadedTaskCount} project tasks, {inventoryMaxDepth}{' '}
                  {inventoryMaxDepth === 1 ? 'level' : 'levels'}
                </p>
                <p
                  className={`mt-1 text-[11px] ${
                    projectInventory.collection.complete ? 'text-emerald-300' : 'text-amber-300'
                  }`}
                >
                  {projectInventory.collection.complete
                    ? 'Task inventory complete.'
                    : 'Task inventory needs review.'}
                </p>

                {projectInventory.collection.warnings.length > 0 && (
                  <ul className="mt-2 space-y-1 text-[11px] leading-4 text-amber-200">
                    {projectInventory.collection.warnings.map((warning) => (
                      <li key={warning}>- {warning}</li>
                    ))}
                  </ul>
                )}

                <button
                  className="mt-3 w-full rounded-md bg-cyan-500 px-3 py-2 text-xs font-semibold text-slate-950 transition hover:bg-cyan-400 disabled:cursor-not-allowed disabled:opacity-60"
                  disabled={isOpeningProjectReview || !canOpenProjectReview}
                  onClick={() => {
                    void handleOpenProjectReview();
                  }}
                  type="button"
                >
                  {isOpeningProjectReview
                    ? 'Opening project collection...'
                    : 'Collect all task discussions'}
                </button>
                <p className="mt-2 text-[10px] text-slate-500">
                  Progress and task details will open in a separate full-page view.
                </p>
              </div>
            )}
          </section>
        )}

        <section className="mb-4 rounded-md border border-slate-800 bg-slate-900 p-4">
          <div className="flex items-start justify-between gap-3">
            <div>
              <h2 className="text-sm font-semibold text-slate-100">Báo cáo</h2>
              <p className="mt-1 text-[11px] leading-4 text-slate-500">
                Báo cáo sử dụng toàn bộ nội dung đã lưu, không chỉ kết quả đang lọc.
              </p>
            </div>
          </div>

          <div className="mt-3 flex flex-wrap items-end gap-2">
            <label className="min-w-0 flex-1 text-[11px] font-medium text-slate-400">
              Loại báo cáo
              <select
                className="mt-1 w-full rounded-md border border-slate-700 bg-slate-950 px-2 py-1.5 text-xs font-medium text-slate-100 outline-none transition hover:border-slate-500 focus:border-blue-500"
                onChange={(event) => {
                  setReportType(event.target.value as ReportType);
                  setReportFeedback('');
                }}
                value={reportType}
              >
                {REPORT_OPTIONS.map((option) => (
                  <option key={option.value} value={option.value}>
                    {option.label}
                  </option>
                ))}
              </select>
            </label>

            <button
              className="rounded-md border border-slate-700 px-3 py-1.5 text-xs font-medium text-slate-200 transition hover:border-slate-500 hover:bg-slate-800 disabled:cursor-not-allowed disabled:opacity-60"
              disabled={isLoading}
              onClick={handleGenerateReport}
              type="button"
            >
              Tạo báo cáo
            </button>
          </div>

          {reportText && (
            <>
              <textarea
                className="mt-3 h-32 w-full resize-none rounded-md border border-slate-700 bg-slate-950 px-3 py-2 text-xs leading-5 text-slate-100 outline-none"
                readOnly
                value={reportText}
              />

              <div className="mt-2 flex items-center justify-between gap-3">
                <p
                  className={`min-h-4 text-[11px] ${
                    reportFeedback.startsWith('Không thể') ? 'text-red-200' : 'text-emerald-300'
                  }`}
                >
                  {reportFeedback}
                </p>
                <button
                  className="rounded-md border border-slate-700 px-3 py-1.5 text-xs font-medium text-slate-200 transition hover:border-slate-500 hover:bg-slate-800 disabled:cursor-not-allowed disabled:opacity-60"
                  disabled={!reportText}
                  onClick={() => {
                    void handleCopyReport();
                  }}
                  type="button"
                >
                  Sao chép
                </button>
              </div>
            </>
          )}
        </section>

        <section className="mb-4 rounded-md border border-slate-800 bg-slate-900 p-4">
          <h2 className="text-sm font-semibold text-slate-100">Mẫu tin nhắn</h2>
          <p className="mt-1 text-[11px] leading-4 text-slate-500">
            Mẫu tin nhắn sử dụng toàn bộ nội dung đã lưu, không chỉ kết quả đang lọc.
          </p>

          <div className="mt-3 flex flex-wrap items-end gap-2">
            <label className="min-w-0 flex-1 text-[11px] font-medium text-slate-400">
              Loại mẫu
              <select
                className="mt-1 w-full rounded-md border border-slate-700 bg-slate-950 px-2 py-1.5 text-xs font-medium text-slate-100 outline-none transition hover:border-slate-500 focus:border-blue-500"
                onChange={(event) => {
                  setTemplateType(event.target.value as TemplateType);
                  setTemplateFeedback('');
                }}
                value={templateType}
              >
                {TEMPLATE_OPTIONS.map((option) => (
                  <option key={option.value} value={option.value}>
                    {option.label}
                  </option>
                ))}
              </select>
            </label>

            <button
              className="rounded-md border border-slate-700 px-3 py-1.5 text-xs font-medium text-slate-200 transition hover:border-slate-500 hover:bg-slate-800 disabled:cursor-not-allowed disabled:opacity-60"
              disabled={isLoading}
              onClick={handleGenerateTemplate}
              type="button"
            >
              Tạo mẫu
            </button>
          </div>

          {templateText && (
            <>
              <textarea
                className="mt-3 h-32 w-full resize-none rounded-md border border-slate-700 bg-slate-950 px-3 py-2 text-xs leading-5 text-slate-100 outline-none"
                readOnly
                value={templateText}
              />

              <div className="mt-2 flex items-center justify-between gap-3">
                <p
                  className={`min-h-4 text-[11px] ${
                    templateFeedback.startsWith('Không thể') ? 'text-red-200' : 'text-emerald-300'
                  }`}
                >
                  {templateFeedback}
                </p>
                <button
                  className="rounded-md border border-slate-700 px-3 py-1.5 text-xs font-medium text-slate-200 transition hover:border-slate-500 hover:bg-slate-800 disabled:cursor-not-allowed disabled:opacity-60"
                  disabled={!templateText}
                  onClick={() => {
                    void handleCopyTemplate();
                  }}
                  type="button"
                >
                  Sao chép
                </button>
              </div>
            </>
          )}
        </section>

        {isLoading ? (
          <div className="rounded-md border border-slate-800 bg-slate-900 p-4 text-sm text-slate-300">
            Đang tải nội dung đã lưu...
          </div>
        ) : errorMessage ? (
          <div className="rounded-md border border-red-900 bg-red-950/40 p-4">
            <p className="text-sm font-medium text-red-100">Không thể tải nội dung đã lưu.</p>
            <p className="mt-2 text-xs text-red-200">{errorMessage}</p>
          </div>
        ) : items.length === 0 ? (
          <div className="rounded-md border border-slate-800 bg-slate-900 p-4">
            <p className="text-sm font-medium text-slate-100">Chưa có nội dung nào được lưu.</p>
            <p className="mt-2 text-xs leading-5 text-slate-400">
              Bôi đen nội dung trên trang web, nhấp chuột phải và chọn Lưu vào T-eraClip.
            </p>
          </div>
        ) : filteredItems.length === 0 ? (
          <div className="rounded-md border border-slate-800 bg-slate-900 p-4">
            <p className="text-sm font-medium text-slate-100">Không có nội dung phù hợp với bộ lọc.</p>
            <p className="mt-2 text-xs leading-5 text-slate-400">Xóa bộ lọc để hiển thị toàn bộ nội dung.</p>
          </div>
        ) : (
          <div className="space-y-3">
            {updateErrorMessage && (
              <div className="rounded-md border border-red-900 bg-red-950/40 p-3 text-xs text-red-100">
                {updateErrorMessage}
              </div>
            )}

            {filteredItems.map((item) => {
              const draft = itemDrafts[item.id] ?? {
                title: item.title,
                notes: item.notes
              };
              const isUpdating = updatingItemIds.has(item.id);
              const isTitleDirty = draft.title !== item.title;
              const isNotesDirty = draft.notes !== item.notes;

              return (
                <article className="rounded-md border border-slate-800 bg-slate-900 p-4" key={item.id}>
                  <div className="flex items-start justify-between gap-3">
                    <label className="min-w-0 flex-1 text-[11px] font-medium text-slate-400">
                      Tiêu đề
                      <div className="mt-1 flex gap-2">
                        <input
                          className="min-w-0 flex-1 rounded-md border border-slate-700 bg-slate-950 px-2 py-1.5 text-sm font-semibold leading-5 text-slate-100 outline-none transition placeholder:text-slate-600 hover:border-slate-500 focus:border-blue-500 disabled:cursor-not-allowed disabled:opacity-60"
                          disabled={isUpdating}
                          onChange={(event) => {
                            setItemDrafts((currentDrafts) => ({
                              ...currentDrafts,
                              [item.id]: {
                                title: event.target.value,
                                notes: currentDrafts[item.id]?.notes ?? item.notes
                              }
                            }));
                          }}
                          value={draft.title}
                        />
                        <button
                          className="rounded-md border border-slate-700 px-2 py-1 text-[11px] font-medium text-slate-200 transition hover:border-slate-500 hover:bg-slate-800 disabled:cursor-not-allowed disabled:opacity-60"
                          disabled={isUpdating || !isTitleDirty}
                          onClick={() => {
                            void handleTitleSave(item, draft.title);
                          }}
                          type="button"
                        >
                          Lưu
                        </button>
                      </div>
                    </label>
                    <time className="shrink-0 pt-5 text-[11px] text-slate-500">{formatCreatedAt(item.createdAt)}</time>
                  </div>

                  {validationErrors[item.id] && (
                    <p className="mt-2 text-[11px] font-medium text-red-200">{validationErrors[item.id]}</p>
                  )}

                  <p className="mt-3 line-clamp-3 text-xs leading-5 text-slate-300">{item.originalText}</p>

                  {(item.sourceTitle || item.sourceUrl) && (
                    <div className="mt-3 border-t border-slate-800 pt-3">
                      {item.sourceTitle && <p className="truncate text-xs text-slate-400">{item.sourceTitle}</p>}
                      {item.sourceUrl && <p className="mt-1 truncate text-[11px] text-slate-500">{item.sourceUrl}</p>}
                    </div>
                  )}

                  <div className="mt-3 flex flex-wrap items-center gap-2">
                    <label className="flex items-center gap-2 text-[11px] font-medium text-slate-400">
                      Trạng thái
                      <select
                        className="rounded-md border border-slate-700 bg-slate-950 px-2 py-1 text-[11px] font-medium capitalize text-slate-100 outline-none transition hover:border-slate-500 focus:border-blue-500 disabled:cursor-not-allowed disabled:opacity-60"
                        disabled={isUpdating}
                        onChange={(event) => {
                          void handleStatusChange(item, event.target.value as TeraClipStatus);
                        }}
                        value={item.status}
                      >
                        {STATUS_OPTIONS.map((status) => (
                          <option key={status} value={status}>
                            {STATUS_LABELS[status]}
                          </option>
                        ))}
                      </select>
                    </label>

                    <label className="flex items-center gap-2 text-[11px] font-medium text-slate-400">
                      Ưu tiên
                      <select
                        className="rounded-md border border-slate-700 bg-slate-950 px-2 py-1 text-[11px] font-medium capitalize text-slate-100 outline-none transition hover:border-slate-500 focus:border-blue-500 disabled:cursor-not-allowed disabled:opacity-60"
                        disabled={isUpdating}
                        onChange={(event) => {
                          void handlePriorityChange(item, event.target.value as TeraClipPriority);
                        }}
                        value={item.priority}
                      >
                        {PRIORITY_OPTIONS.map((priority) => (
                          <option key={priority} value={priority}>
                            {PRIORITY_LABELS[priority]}
                          </option>
                        ))}
                      </select>
                    </label>
                  </div>

                  <label className="mt-3 block text-[11px] font-medium text-slate-400">
                    Ghi chú
                    <textarea
                      className="mt-1 h-16 w-full resize-none rounded-md border border-slate-700 bg-slate-950 px-2 py-1.5 text-xs leading-5 text-slate-100 outline-none transition placeholder:text-slate-600 hover:border-slate-500 focus:border-blue-500 disabled:cursor-not-allowed disabled:opacity-60"
                      disabled={isUpdating}
                      onChange={(event) => {
                        setItemDrafts((currentDrafts) => ({
                          ...currentDrafts,
                          [item.id]: {
                            title: currentDrafts[item.id]?.title ?? item.title,
                            notes: event.target.value
                          }
                        }));
                      }}
                      placeholder="Thêm ghi chú..."
                      value={draft.notes}
                    />
                  </label>

                  <div className="mt-2 flex justify-end">
                    <button
                      className="rounded-md border border-slate-700 px-2 py-1 text-[11px] font-medium text-slate-200 transition hover:border-slate-500 hover:bg-slate-800 disabled:cursor-not-allowed disabled:opacity-60"
                      disabled={isUpdating || !isNotesDirty}
                      onClick={() => {
                        void handleNotesSave(item, draft.notes);
                      }}
                      type="button"
                    >
                      Lưu ghi chú
                    </button>
                  </div>
                </article>
              );
            })}
          </div>
        )}
      </fieldset>
    </div>
  );
}
