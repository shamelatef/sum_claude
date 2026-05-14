import { LAYOUT } from '../config/constants';

export function contentY(): number   { return LAYOUT.contentStartY; }
export function contentBottom(): number { return LAYOUT.contentEndY; }
export function slideW(): number     { return LAYOUT.slideW; }
export function slideH(): number     { return LAYOUT.slideH; }
export function marginLeft(): number { return LAYOUT.marginL; }
export function usableWidth(): number { return LAYOUT.usableW; }

export function clamp(val: number, min: number, max: number): number {
  return Math.max(min, Math.min(max, val));
}

/** Grid distribution for N cards */
export function distributeCards(
  count: number,
  maxPerRow = 3
): { cols: number; rows: number } {
  if (count <= 1) return { cols: 1, rows: 1 };
  if (count <= 2) return { cols: 2, rows: 1 };
  if (count <= 4) return { cols: 2, rows: Math.ceil(count / 2) };
  if (count <= 6) return { cols: 3, rows: Math.ceil(count / 3) };
  return { cols: maxPerRow, rows: Math.ceil(count / maxPerRow) };
}
