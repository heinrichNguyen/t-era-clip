export const ONEOFFICE_ALLOWED_ORIGIN = 'https://office.meygroup.vn';
export const ONEOFFICE_ORIGINS = [
  ONEOFFICE_ALLOWED_ORIGIN,
  'https://testoffice.meygroup.vn',
  'https://testoffice.meygroup.vn:9443',
] as const;

export function oneOfficeOrigin(value: string): string {
  const url = new URL(value);
  if (url.username || url.password || !ONEOFFICE_ORIGINS.some((origin) => origin === url.origin)) {
    throw new Error('Địa chỉ không thuộc môi trường 1Office được hỗ trợ.');
  }
  return url.origin;
}

export async function getOneOfficeTabOrigin(tabId: number, expectedOrigin?: string): Promise<string> {
  const api = (globalThis as any).chrome;
  if (!api?.tabs?.get) throw new Error('Không thể xác minh địa chỉ tab nguồn. Hãy mở lại extension từ tab 1Office.');
  const tab = await api.tabs.get(tabId);
  const origin = oneOfficeOrigin(tab?.url ?? '');
  if (expectedOrigin !== undefined && origin !== oneOfficeOrigin(expectedOrigin)) {
    throw new Error('Tab nguồn đã chuyển môi trường 1Office. Hãy mở lại từ đúng dự án.');
  }
  return origin;
}
