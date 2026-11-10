import { useEffect, useRef, useState } from 'react';
import { DISCUSSIONS_ONLY } from '../shared/buildProfile';

const REVIEW_KEYS = DISCUSSIONS_ONLY ? ['oneOfficeDiscussionReview', 'oneOfficeProjectBatchSession'] : ['oneOfficeDiscussionReview', 'oneOfficeProjectBatchSession', 'oneOfficeNoReview'];
const CLEAR_REVISION_KEY = 'oneOfficeTemporaryClearRevision';

export function useTemporaryData(onClear: () => void) {
  const callback = useRef(onClear);
  callback.current = onClear;
  const clearedRef = useRef(false);
  const revisionRef = useRef(0);
  const [cleared, setCleared] = useState(false);
  const [clearing, setClearing] = useState(false);
  const [clearError, setClearError] = useState('');
  const wipe = () => {
    clearedRef.current = true;
    revisionRef.current++;
    callback.current();
    setCleared(true);
  };

  useEffect(() => {
    const api = (globalThis as any).chrome;
    const listener = (changes: Record<string, unknown>, area: string) => {
      if (area === 'session' && CLEAR_REVISION_KEY in changes) wipe();
    };
    api?.storage?.onChanged?.addListener(listener);
    return () => api?.storage?.onChanged?.removeListener(listener);
  }, []);

  const clearTemporaryData = async () => {
    if (clearing || !window.confirm(DISCUSSIONS_ONLY ? 'Xóa dữ liệu báo cáo tạm trong các trang T-eraClip đang mở? Không xóa file đã tải hoặc dữ liệu trên 1Office.' : 'Xóa dữ liệu báo cáo tạm trong các trang T-eraClip đang mở? Không xóa Capture, file đã tải hoặc dữ liệu trên 1Office.')) return;
    setClearing(true); setClearError('');
    const api = (globalThis as any).chrome;
    try {
      const storage = api?.storage?.session;
      if (!storage) throw new Error('Không thể xóa dữ liệu tạm. Hãy tải lại extension và thử lại.');
      const operation = (run: (done: () => void) => void) => new Promise<void>((resolve, reject) => {
        run(() => api.runtime?.lastError ? reject(new Error(api.runtime.lastError.message)) : resolve());
      });
      await operation((done) => storage.remove(REVIEW_KEYS, done));
      // A session-storage change invalidates memory-only reports in other open extension pages.
      await operation((done) => storage.set({ [CLEAR_REVISION_KEY]: crypto.randomUUID() }, done));
      wipe();
    } catch (error) {
      setClearError(error instanceof Error ? error.message : 'Không thể xóa dữ liệu tạm. Hãy thử lại.');
    } finally { setClearing(false); }
  };
  return { clearTemporaryData, cleared, clearedRef, revisionRef, clearing, clearError,
    reset: () => { clearedRef.current = false; setCleared(false); setClearError(''); } };
}
