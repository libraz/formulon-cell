import { addrKey, MAX_COL, MAX_ROW, parseAddrKey } from '../engine/address.js';
import type { Range } from '../engine/types.js';
import type { CellFormat, ValueFilterCriteria } from '../store/store.js';

export interface AxisEdit {
  at: number;
  count: number;
}

export function normalizeAxisEdit(
  at: number,
  count: number,
  max: number,
  kind: 'insert' | 'delete',
): AxisEdit | null {
  if (!Number.isInteger(at) || !Number.isInteger(count)) return null;
  if (!Number.isFinite(at) || !Number.isFinite(count)) return null;
  if (at < 0 || at > max || count <= 0) return null;
  // A row/column insertion must leave room for every cell that is moved
  // right/down. Deletions may consume the final row/column.
  const remaining = kind === 'insert' ? max - at : max + 1 - at;
  const normalizedCount = Math.min(count, remaining);
  return normalizedCount > 0 ? { at, count: normalizedCount } : null;
}

export function cloneInsertedFormat(format: CellFormat): CellFormat {
  const next: CellFormat = { ...format };
  delete next.hyperlink;
  delete next.hyperlinkDisplay;
  delete next.hyperlinkTooltip;
  delete next.comment;
  delete next.commentAuthor;
  delete next.validation;
  if (format.borders) next.borders = { ...format.borders };
  if (format.numFmt) next.numFmt = { ...format.numFmt };
  if (format.phonetic) next.phonetic = format.phonetic.map((run) => ({ ...run }));
  return next;
}

/** Shift indices in a sparse Map keyed by integer index. Indices >= split move
 *  by `delta` (delta>0 = shift right/down, delta<0 = shift left/up). For
 *  delete (delta<0), keys in [split, split+|delta|) are removed. */
export function shiftIndexedMap(
  src: Map<number, number>,
  split: number,
  delta: number,
  max = Number.POSITIVE_INFINITY,
): Map<number, number> {
  const out = new Map<number, number>();
  for (const [k, v] of src) {
    if (k < split) {
      out.set(k, v);
      continue;
    }
    if (delta < 0 && k < split - delta) continue; // dropped
    const next = k + delta;
    if (next < 0 || next > max) continue;
    out.set(next, v);
  }
  return out;
}

export function shiftIndexedSet(
  src: Set<number>,
  split: number,
  delta: number,
  max = Number.POSITIVE_INFINITY,
): Set<number> {
  const out = new Set<number>();
  for (const k of src) {
    if (k < split) {
      out.add(k);
      continue;
    }
    if (delta < 0 && k < split - delta) continue;
    const next = k + delta;
    if (next < 0 || next > max) continue;
    out.add(next);
  }
  return out;
}

/** Shift addrKey-keyed formats so any key with row >= splitRow moves by deltaRow.
 *  When deltaRow < 0, formats in the deleted band are dropped. */
export function shiftFormatsByRow(
  src: Map<string, CellFormat>,
  sheet: number,
  splitRow: number,
  deltaRow: number,
): Map<string, CellFormat> {
  const out = new Map<string, CellFormat>();
  for (const [key, fmt] of src) {
    const addr = parseAddrKey(key);
    if (!addr) {
      out.set(key, fmt);
      continue;
    }
    const { sheet: s, row: r, col: c } = addr;
    if (s !== sheet || r < splitRow) {
      out.set(key, fmt);
      continue;
    }
    if (deltaRow < 0 && r < splitRow - deltaRow) continue; // in deleted band
    const nextRow = r + deltaRow;
    if (nextRow < 0 || nextRow > MAX_ROW) continue;
    out.set(addrKey({ sheet: s, row: nextRow, col: c }), fmt);
  }
  return out;
}

export function inheritFormatsByRow(
  src: Map<string, CellFormat>,
  shifted: Map<string, CellFormat>,
  sheet: number,
  splitRow: number,
  count: number,
): Map<string, CellFormat> {
  if (splitRow <= 0 || count <= 0) return shifted;
  const out = new Map(shifted);
  for (const [key, fmt] of src) {
    const addr = parseAddrKey(key);
    if (!addr || addr.sheet !== sheet || addr.row !== splitRow - 1) continue;
    for (let row = splitRow; row < splitRow + count && row <= MAX_ROW; row += 1) {
      out.set(addrKey({ sheet, row, col: addr.col }), cloneInsertedFormat(fmt));
    }
  }
  return out;
}

export function shiftFormatsByCol(
  src: Map<string, CellFormat>,
  sheet: number,
  splitCol: number,
  deltaCol: number,
): Map<string, CellFormat> {
  const out = new Map<string, CellFormat>();
  for (const [key, fmt] of src) {
    const addr = parseAddrKey(key);
    if (!addr) {
      out.set(key, fmt);
      continue;
    }
    const { sheet: s, row: r, col: c } = addr;
    if (s !== sheet || c < splitCol) {
      out.set(key, fmt);
      continue;
    }
    if (deltaCol < 0 && c < splitCol - deltaCol) continue;
    const nextCol = c + deltaCol;
    if (nextCol < 0 || nextCol > MAX_COL) continue;
    out.set(addrKey({ sheet: s, row: r, col: nextCol }), fmt);
  }
  return out;
}

export function inheritFormatsByCol(
  src: Map<string, CellFormat>,
  shifted: Map<string, CellFormat>,
  sheet: number,
  splitCol: number,
  count: number,
): Map<string, CellFormat> {
  if (splitCol <= 0 || count <= 0) return shifted;
  const out = new Map(shifted);
  for (const [key, fmt] of src) {
    const addr = parseAddrKey(key);
    if (!addr || addr.sheet !== sheet || addr.col !== splitCol - 1) continue;
    for (let col = splitCol; col < splitCol + count && col <= MAX_COL; col += 1) {
      out.set(addrKey({ sheet, row: addr.row, col }), cloneInsertedFormat(fmt));
    }
  }
  return out;
}

export function shiftIndexedMapWithInheritance(
  src: Map<number, number>,
  split: number,
  delta: number,
  count: number,
  max = Number.POSITIVE_INFINITY,
): Map<number, number> {
  const out = shiftIndexedMap(src, split, delta, max);
  if (split <= 0 || count <= 0) return out;
  const inherited = src.get(split - 1);
  if (inherited === undefined) return out;
  for (let index = split; index < split + count && index <= max; index += 1) {
    out.set(index, inherited);
  }
  return out;
}

/** Map a 1-D interval [lo,hi] through a row/col insert (delta>0) or delete
 *  (delta<0) at `split`. Returns null when a deletion consumes the whole
 *  interval. Inserting inside a span widens it — matching how spreadsheets grow
 *  a merge / conditional-format / filter region when rows or cols are added
 *  within it. */
export function adjustInterval(
  lo: number,
  hi: number,
  split: number,
  delta: number,
): [number, number] | null {
  if (delta > 0) {
    return [lo >= split ? lo + delta : lo, hi >= split ? hi + delta : hi];
  }
  const count = -delta;
  const bandHi = split + count - 1;
  const nlo = lo < split ? lo : lo > bandHi ? lo - count : split;
  const nhi = hi < split ? hi : hi > bandHi ? hi - count : split - 1;
  return nlo > nhi ? null : [nlo, nhi];
}

/** Shift a range's row or col span for a structure edit. Returns null when a
 *  deletion removes the whole span. */
export function shiftRangeAxis(
  range: Range,
  axis: 'row' | 'col',
  split: number,
  delta: number,
): Range | null {
  const lo = axis === 'row' ? range.r0 : range.c0;
  const hi = axis === 'row' ? range.r1 : range.c1;
  const res = adjustInterval(lo, hi, split, delta);
  if (!res) return null;
  const [a, b] = res;
  const max = axis === 'row' ? MAX_ROW : MAX_COL;
  if (a < 0 || b > max) return null;
  return axis === 'row' ? { ...range, r0: a, r1: b } : { ...range, c0: a, c1: b };
}

export function shiftFilterCriteria(
  criteria: readonly ValueFilterCriteria[],
  sheet: number,
  axis: 'row' | 'col',
  split: number,
  delta: number,
): ValueFilterCriteria[] {
  const out: ValueFilterCriteria[] = [];
  for (const c of criteria) {
    if (c.range.sheet !== sheet) {
      out.push(c);
      continue;
    }
    const shifted = shiftRangeAxis(c.range, axis, split, delta);
    if (!shifted) continue; // filtered column removed
    let byCol = c.byCol;
    if (axis === 'col') {
      const mapped = adjustInterval(byCol, byCol, split, delta);
      if (!mapped) continue; // this column was deleted
      byCol = mapped[0];
    }
    out.push({ ...c, range: shifted, byCol });
  }
  return out;
}
