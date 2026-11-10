export const companyScopes = [
  { code: '', title: 'Toàn bộ', description: 'Tất cả tài khoản active', members: [] },
  { code: 'TDR', title: 'TDR / Meyland', description: 'Bất động sản', members: [] },
  { code: 'TDG', title: 'TDG', description: 'Đầu tư Tập đoàn', members: [] },
  { code: 'TDT', title: 'TDT / T ERA', description: 'Công nghệ cao', members: [] },
  { code: 'TDK', title: 'TDK', description: 'Công nghiệp và Công nghệ', members: [
    { code: 'TDH', title: 'Kim khí và Gia dụng' }, { code: 'TDP', title: 'Nhựa' },
    { code: 'TDC', title: 'Sơn & Hóa chất' }, { code: 'TDW', title: 'Công nghệ và Giải pháp nước' },
  ] },
  { code: 'TDU', title: 'TDU', description: 'Đầu tư xây dựng HUD Kiên Giang', members: [
    { code: 'TDU.CKKG', title: 'Cơ khí Kiên Giang' }, { code: 'TDU.GNKG', title: 'Gạch ngói Kiên Giang' },
    { code: 'TDU.VLXD', title: 'Sản xuất Vật liệu xây dựng Kiên Giang' }, { code: 'TDU.XMHT', title: 'Xi măng Hà Tiên Kiên Giang' },
  ] },
] as const;

export function companyScopeLabel(code: string): string {
  for (const scope of companyScopes) {
    if (scope.code === code) return scope.members.length ? `${scope.title}: toàn nhóm (${scope.members.map((member) => member.code).join(', ')})` : scope.title;
    const member = scope.members.find((item) => item.code === code);
    if (member) return `${member.code}: ${member.title}`;
  }
  return code;
}

export function matchesCompanyScope(code: string, companyCode: string, groupCode?: string): boolean | undefined {
  const selected = code.toUpperCase();
  const scope = companyScopes.find((item) => item.code === selected);
  if (scope && selected) return companyCode.toUpperCase() === selected ||
    scope.members.some((member) => member.code === companyCode.toUpperCase()) ||
    (scope.members.length > 0 && groupCode === selected);
  if (companyScopes.some((item) => item.members.some((member) => member.code === selected))) return companyCode.toUpperCase() === selected;
  return undefined;
}
