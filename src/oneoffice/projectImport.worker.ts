import readExcelFile from 'read-excel-file/browser';

self.onmessage = async (event: MessageEvent<ArrayBuffer>) => {
  try {
    const sheets = await readExcelFile(event.data);
    if (sheets.length > 20 || sheets.some((s) => s.data.length > 10001 || s.data.some((r) => r.length > 100))) {
      throw new Error('File vượt giới hạn 20 sheet, 10.000 dòng hoặc 100 cột mỗi sheet.');
    }
    self.postMessage({ sheets });
  } catch {
    self.postMessage({ error: 'Không đọc được file Excel hoặc file vượt giới hạn. Chỉ nhận .xlsx, tối đa 20 sheet, 10.000 dòng và 100 cột mỗi sheet.' });
  }
};
