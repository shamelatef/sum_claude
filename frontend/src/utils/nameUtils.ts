export function cleanPMName(raw: string | null | undefined): string {
  if (!raw) return '';
  const str = String(raw).trim();
  const commaIdx = str.indexOf(',');
  return commaIdx === -1 ? str : str.slice(0, commaIdx).trim();
}

export function normalizeGate(raw: string | null | undefined): string {
  if (!raw) return 'Unspecified';
  const str = String(raw).trim();
  return str.length > 0 ? str : 'Unspecified';
}

export function truncate(text: string, maxLen: number): string {
  if (text.length <= maxLen) return text;
  return text.slice(0, maxLen - 1) + '…';
}
