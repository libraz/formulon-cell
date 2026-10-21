import { addrKey, MAX_COL, MAX_ROW } from '../engine/address.js';
import type { Addr } from '../engine/types.js';
import type { State } from '../store/store.js';

export const formatHasContent = (format: object | undefined): boolean =>
  format !== undefined && Object.keys(format).length > 0;

export const isPopulated = (s: State, sheet: number, row: number, col: number): boolean => {
  const key = addrKey({ sheet, row, col });
  return s.data.cells.has(key) || formatHasContent(s.format.formats.get(key));
};

/** Excel's Current Region: expand from the active cell through adjacent
 * non-empty rows and columns. It is intentionally based on model content
 * (including format-only cells), matching the same definition Ctrl+Arrow
 * uses. */
export function currentRegion(
  s: State,
  a: Addr,
): { sheet: number; r0: number; c0: number; r1: number; c1: number } | null {
  if (!isPopulated(s, a.sheet, a.row, a.col)) return null;
  let r0 = a.row;
  let r1 = a.row;
  let c0 = a.col;
  let c1 = a.col;
  let changed = true;
  const rowHasContent = (row: number): boolean => {
    for (let col = c0; col <= c1; col += 1) {
      if (isPopulated(s, a.sheet, row, col)) return true;
    }
    return false;
  };
  const colHasContent = (col: number): boolean => {
    for (let row = r0; row <= r1; row += 1) {
      if (isPopulated(s, a.sheet, row, col)) return true;
    }
    return false;
  };
  while (changed) {
    changed = false;
    while (r0 > 0 && rowHasContent(r0 - 1)) {
      r0 -= 1;
      changed = true;
    }
    while (r1 < MAX_ROW && rowHasContent(r1 + 1)) {
      r1 += 1;
      changed = true;
    }
    while (c0 > 0 && colHasContent(c0 - 1)) {
      c0 -= 1;
      changed = true;
    }
    while (c1 < MAX_COL && colHasContent(c1 + 1)) {
      c1 += 1;
      changed = true;
    }
  }
  return { sheet: a.sheet, r0, c0, r1, c1 };
}
