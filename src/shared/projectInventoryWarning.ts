export function projectInventoryWarningLabel(warning: string) {
  const duplicateNo = /^(\d+) duplicate task No\. value\(s\) were found\.$/.exec(warning);
  return duplicateNo ? `Có ${duplicateNo[1]} giá trị No. bị trùng.` : warning;
}
