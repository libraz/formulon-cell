/** A1 print-range grammar: print areas, print-title rows/columns, and page-order sorting of print regions. */
import { colFromLetters, parseA1Atom } from '../engine/address.js';
import type { PageSetup } from '../store/store.js';

export interface PrintAreaBounds {
  row0: number;
  col0: number;
  row1: number;
  col1: number;
}

/** Parse an A1-style row range like "1:3" / "$1:$3" / "2" → `[r0, r1]`
 *  inclusive, 0-indexed. Returns null on bad input. */
export function parsePrintTitleRows(raw?: string): [number, number] | null {
  if (!raw) return null;
  const trimmed = raw.trim().replace(/\$/g, '');
  if (!trimmed) return null;
  const parts = trimmed.split(':');
  const a = Number.parseInt(parts[0] ?? '', 10);
  if (!Number.isFinite(a) || a < 1) return null;
  if (parts.length === 1) return [a - 1, a - 1];
  const b = Number.parseInt(parts[1] ?? '', 10);
  if (!Number.isFinite(b) || b < 1) return null;
  return [Math.min(a, b) - 1, Math.max(a, b) - 1];
}

/** Parse an A1-style col range like "A:B" / "$A:$B" / "C" → `[c0, c1]`
 *  inclusive, 0-indexed. Returns null on bad input. */
export function parsePrintTitleCols(raw?: string): [number, number] | null {
  if (!raw) return null;
  const trimmed = raw.trim().replace(/\$/g, '');
  if (!trimmed) return null;
  const parts = trimmed.split(':');
  const a = colFromLetters(parts[0] ?? '');
  if (a < 0) return null;
  if (parts.length === 1) return [a, a];
  const b = colFromLetters(parts[1] ?? '');
  if (b < 0) return null;
  return [Math.min(a, b), Math.max(a, b)];
}

/** Parse an A1-style rectangular print area like "A1:D20" / "$A$1:$D$20"
 *  / "B2" → zero-indexed bounds. Returns null on bad input. */
export function parsePrintArea(raw?: string): PrintAreaBounds | null {
  if (!raw) return null;
  const trimmed = raw.trim().replace(/\$/g, '');
  if (!trimmed || trimmed.includes(',')) return null;
  const parts = trimmed.split(':');
  if (parts.length > 2) return null;
  const a = parseA1Atom(parts[0] ?? '');
  const b = parseA1Atom(parts[1] ?? parts[0] ?? '');
  if (!a || !b) return null;
  return {
    row0: Math.min(a.row, b.row),
    col0: Math.min(a.col, b.col),
    row1: Math.max(a.row, b.row),
    col1: Math.max(a.col, b.col),
  };
}

export function parsePrintAreas(raw?: string): PrintAreaBounds[] | null {
  if (!raw) return null;
  const parts = raw
    .split(',')
    .map((part) => part.trim())
    .filter(Boolean);
  if (parts.length === 0) return null;
  const areas = parts.map((part) => parsePrintArea(part));
  if (areas.some((area) => area === null)) return null;
  return areas as PrintAreaBounds[];
}

export function orderPrintRegionsForPageOrder(
  regions: readonly PrintAreaBounds[],
  pageOrder: PageSetup['pageOrder'] = 'downThenOver',
): PrintAreaBounds[] {
  const ordered = [...regions];
  ordered.sort((a, b) =>
    pageOrder === 'overThenDown'
      ? a.row0 - b.row0 || a.col0 - b.col0 || a.row1 - b.row1 || a.col1 - b.col1
      : a.col0 - b.col0 || a.row0 - b.row0 || a.col1 - b.col1 || a.row1 - b.row1,
  );
  return ordered;
}
