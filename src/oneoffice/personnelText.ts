// Inert parsing also handles icon markup encoded as HTML entities.
export function personnelText(value: string): string {
  let result = value;
  for (let pass = 0; pass < 4; pass++) {
    const doc = new DOMParser().parseFromString(result, 'text/html');
    doc.querySelectorAll('script, style, iframe, object').forEach((node) => node.remove());
    const next = (doc.body.textContent ?? '').replace(/\s+/g, ' ').trim();
    if (next === result) break;
    result = next;
  }
  return result === '--' ? '' : result;
}
