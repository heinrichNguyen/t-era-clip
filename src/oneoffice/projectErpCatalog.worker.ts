import { readSheet } from 'read-excel-file/browser';
import { parseProjectErpCatalog } from './projectErpCatalog';

self.onmessage = async (event: MessageEvent<{ buffer: ArrayBuffer; sheet: string }>) => {
  try {
    const sheet = event.data.sheet.trim();
    if (!sheet || sheet.length > 100) throw new Error('Nhập tên sheet danh mục hợp lệ.');
    const data = await readSheet(event.data.buffer, sheet);
    self.postMessage({ catalog: parseProjectErpCatalog({ sheet, data }) });
  } catch (error) {
    self.postMessage({ error: error instanceof Error && !(error.name === 'SheetNotFoundError') ? error.message : 'Không tìm thấy sheet. Kiểm tra tên sheet trong file danh mục.' });
  }
};
