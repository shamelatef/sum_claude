export function cleanPMName(raw: string): string {
  if (!raw) return '';
  const commaIdx = raw.indexOf(',');
  return commaIdx === -1 ? raw.trim() : raw.slice(0, commaIdx).trim();
}

export function downloadBlob(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

export function formatDate(): string {
  return new Date().toISOString().slice(0, 10);
}
