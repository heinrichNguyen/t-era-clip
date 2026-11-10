export interface ErpCatalogEntry { code: string; name: string; row: number }
export interface ProjectErpCatalog {
  sheet: string;
  projects: ErpCatalogEntry[];
  packages: ErpCatalogEntry[];
  warnings: string[];
}
const text = (value: unknown) => typeof value === 'string' || typeof value === 'number' ? String(value).trim() : '';
export const erpSearchText = (value: string) => value.normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
export const validErpCode = (value: string) => !!value && value.length <= 160 && !/[\s\x00-\x1f\x7f]/.test(value) && !/(^\.|\.$|\.\.)/.test(value);

// Only the two explicitly named catalogs are retained. Other sections can contain personal data.
export function parseProjectErpCatalog(sheet: { sheet: string; data: readonly (readonly unknown[])[] }): ProjectErpCatalog {
  if (sheet.data.length > 100000 || sheet.data.some((r) => r.length > 20)) throw new Error('Sheet danh mục vượt giới hạn 100.000 dòng hoặc 20 cột.');
  const result: ProjectErpCatalog = { sheet: sheet.sheet, projects: [], packages: [], warnings: [] };
  let section: 'projects' | 'packages' | null = null;
  let header = false;
  let invalid = 0;
  const seen = new Set<string>();
  for (let i = 0; i < sheet.data.length; i++) {
    const row = sheet.data[i];
    const a = text(row[0]); const b = text(row[1]);
    const title = erpSearchText(a).replace(/^\d+\s*/, '');
    if (/^\d+\s*[.)]?\s*danh\s+mục/i.test(a) || /^danh\s+mục/i.test(a)) {
      section = title === 'danh muc goi thau' ? 'packages' : title === 'danh muc hang muc' ? 'projects' : null;
      header = false;
      continue;
    }
    if (!section) continue;
    if (erpSearchText(a) === 'ma' && erpSearchText(b) === 'ten') { header = true; continue; }
    if (!a && !b) continue;
    if (!header || !validErpCode(a) || !b || b.length > 2000) { invalid++; continue; }
    const key = JSON.stringify([section, a, b]);
    if (!seen.has(key)) { result[section].push({ code: a, name: b, row: i + 1 }); seen.add(key); }
  }
  if (!result.packages.length) throw new Error('Không tìm thấy Danh mục gói thầu có hai cột Mã, Tên trong sheet đã chọn.');
  if (result.packages.length + result.projects.length > 20000) throw new Error('Danh mục vượt giới hạn 20.000 mã.');
  if (!result.projects.length) result.warnings.push('Sheet này không có Danh mục hạng mục dùng cho mã dự án; chưa xác minh mã dự án.');
  if (invalid) result.warnings.push(`${invalid} dòng trong vùng danh mục không hợp lệ đã bị bỏ qua; danh mục có thể chưa đầy đủ.`);
  const codes = new Set<string>();
  for (const entry of result.packages) {
    if (codes.has(entry.code)) { result.warnings.push('Có mã gói thầu ứng với nhiều tên. Các mã đó cần kiểm tra, không tự chọn.'); break; }
    codes.add(entry.code);
  }
  return result;
}

export function indexErpPackages(catalog: ProjectErpCatalog) {
  const index = new Map<string, ErpCatalogEntry[]>();
  for (const entry of catalog.packages) { const entries = index.get(entry.code) || []; entries.push(entry); index.set(entry.code, entries); }
  return index;
}

export function indexErpPackageNames(catalog: ProjectErpCatalog) {
  const index = new Map<string, Set<string>>();
  for (const entry of catalog.packages) {
    const name = erpPackageNameKey(entry);
    const codes = index.get(name) || new Set<string>(); codes.add(entry.code); index.set(name, codes);
  }
  return index;
}

export function erpPackageNameKey(entry: ErpCatalogEntry) {
  const name = entry.name.startsWith(entry.code + ':') ? entry.name.slice(entry.code.length + 1) : entry.name;
  return name.normalize('NFC').trim().replace(/\s+/g, ' ').toLowerCase();
}

const searchName = (value: string) => erpSearchText(value).replace(/\bqh\b/g, 'quy hoach');
const COMMON_SEARCH_WORDS = new Set('tu van thiet ke goi thau cong viec dich vu thuc hien lap phan va cua cho ve theo du an noi dung phuc'.split(' '));

export function suggestErpPackages(catalog: ProjectErpCatalog, query: string) {
  const bounded = query.slice(0, 300);
  const search = searchName(bounded);
  const codeSearch = erpSearchText(bounded);
  if (!search) return { entries: [] as ErpCatalogEntry[], total: 0 };
  const allTokens = [...new Set(search.split(' '))];
  const specificTokens = allTokens.filter((t) => t.length > 1 && !COMMON_SEARCH_WORDS.has(t) && !/^\d+$/.test(t));
  const tokens = specificTokens.length ? specificTokens : allTokens;
  // Ratios/numeric pairs are explicit query constraints, never inferred ERP identity.
  const qualifiers = [...bounded.matchAll(/\b\d+\s*[/ :]\s*\d+\b/g)].map((m) => erpSearchText(m[0]));
  const entries = catalog.packages.map((entry) => {
    const code = erpSearchText(entry.code); const name = searchName(entry.name);
    if (code === codeSearch) return { entry, score: 20000 };
    if (qualifiers.some((q) => !` ${name} `.includes(` ${q} `))) return { entry, score: 0 };
    if (code.startsWith(codeSearch + ' ')) return { entry, score: 15000 };
    if (name === search) return { entry, score: 10000 };
    const words = new Set(name.split(' '));
    const matched = tokens.filter((t) => words.has(t)).length;
    const minimum = Math.ceil(tokens.length * (specificTokens.length ? 0.6 : 1));
    const score = matched < minimum ? 0 : (name.includes(search) ? 1000 : 0) + 100 * matched / tokens.length;
    return { entry, score };
  }).filter((r) => r.score > 0).sort((a, b) => b.score - a.score || a.entry.row - b.entry.row);
  return { entries: entries.slice(0, 20).map((r) => r.entry), total: entries.length };
}
