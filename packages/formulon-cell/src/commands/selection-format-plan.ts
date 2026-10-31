/** Selection geometry planning and the cell budget shared by format commands. */

import { addrKey, MAX_COL, MAX_ROW, parseAddrKey } from '../engine/address.js';
import type { AxisFormatBand } from '../engine/axis-format-sync.js';
import type { Addr, Range } from '../engine/types.js';
import {
  rangeContainsAddr,
  rangesIntersect,
  selectionCoversRange,
  subtractRange,
} from '../store/selection-geometry.js';
import type { CellFormat, State } from '../store/store.js';
import { planAxisSelection } from './axis-selection-plan.js';

export const MAX_MATERIALIZED_FORMAT_CELLS = 100_000;

const validSelectionRange = (range: Range, sheet: number): boolean =>
  range.sheet === sheet &&
  Number.isInteger(range.sheet) &&
  Number.isInteger(range.r0) &&
  Number.isInteger(range.c0) &&
  Number.isInteger(range.r1) &&
  Number.isInteger(range.c1) &&
  range.sheet >= 0 &&
  range.r0 >= 0 &&
  range.c0 >= 0 &&
  range.r0 <= range.r1 &&
  range.c0 <= range.c1 &&
  range.r1 <= MAX_ROW &&
  range.c1 <= MAX_COL;

export const selectionRanges = (state: State): Range[] | null => {
  const ranges = [state.selection.range, ...(state.selection.extraRanges ?? [])];
  const primary = ranges[0];
  if (!primary || !validSelectionRange(primary, primary.sheet)) return null;
  if (!ranges.every((range) => validSelectionRange(range, primary.sheet))) return null;
  return ranges.map((range) => ({ ...range }));
};

const disjointRanges = (ranges: readonly Range[]): Range[] => {
  const result: Range[] = [];
  for (const range of ranges) {
    let pieces = [{ ...range }];
    for (const cover of result) {
      pieces = pieces.flatMap((piece) => subtractRange(piece, cover));
      if (pieces.length === 0) break;
    }
    result.push(...pieces);
  }
  return result;
};

const mergeCoverageValid = (state: State, ranges: readonly Range[]): boolean => {
  const selection = { range: ranges[0] as Range, extraRanges: ranges.slice(1) as Range[] };
  for (const merge of state.merges.byAnchor.values()) {
    if (ranges.some((range) => rangesIntersect(range, merge))) {
      if (!selectionCoversRange(selection, merge)) return false;
    }
  }
  return true;
};

const materializeSelection = (ranges: readonly Range[]): Addr[] | null => {
  const cells = new Map<string, Addr>();
  for (const range of disjointRanges(ranges)) {
    for (let row = range.r0; row <= range.r1; row += 1) {
      for (let col = range.c0; col <= range.c1; col += 1) {
        const addr = { sheet: range.sheet, row, col };
        cells.set(addrKey(addr), addr);
        if (cells.size > MAX_MATERIALIZED_FORMAT_CELLS) return null;
      }
    }
  }
  return [...cells.values()];
};

export const selectionPlan = (
  state: State,
): {
  ranges: Range[];
  cells: Addr[];
  exceptionCells: Addr[];
  axisBands: AxisFormatBand[];
} | null => {
  const ranges = selectionRanges(state);
  if (!ranges || !mergeCoverageValid(state, ranges)) return null;
  const normalized = disjointRanges(ranges);
  const geometry = planAxisSelection(ranges, MAX_MATERIALIZED_FORMAT_CELLS);
  if (!geometry) return null;
  const cells = materializeSelection(geometry.ordinaryRanges);
  const exceptionCells = materializeSelection(geometry.exceptionRanges);
  if (
    !cells ||
    !exceptionCells ||
    cells.length + exceptionCells.length > MAX_MATERIALIZED_FORMAT_CELLS
  )
    return null;
  const sparse = new Map<string, Addr>();
  for (const addr of cells) sparse.set(addrKey(addr), addr);
  const exceptions = new Set(exceptionCells.map(addrKey));
  // Axis selections modify the compact native style and only materialise
  // already-authored cells in the store. Blank/future cells remain implicit.
  if (geometry.axisBands.length > 0) {
    for (const key of state.format.formats.keys()) {
      const addr = parseAddrKey(key);
      if (!addr || addr.sheet !== normalized[0]?.sheet) continue;
      if (
        !exceptions.has(key) &&
        geometry.axisBands.some(
          (band) =>
            (band.axis === 'column' && addr.col >= band.first && addr.col <= band.last) ||
            (band.axis === 'row' && addr.row >= band.first && addr.row <= band.last),
        )
      )
        sparse.set(key, addr);
    }
  }
  if (sparse.size + exceptions.size > MAX_MATERIALIZED_FORMAT_CELLS) return null;
  return {
    ranges: normalized,
    cells: [...sparse.values()],
    exceptionCells,
    axisBands: geometry.axisBands.map((band) => ({ ...band })),
  };
};

export interface SelectionFormatPlan {
  readonly ranges: readonly Range[];
  readonly cells: readonly Addr[];
  readonly exceptionCells: readonly Addr[];
  readonly axisBands: readonly AxisFormatBand[];
}

/** Plan the complete primary-plus-extra selection for dialog-style actions. */
export function planSelectionFormat(state: State): SelectionFormatPlan | null {
  const plan = selectionPlan(state);
  if (!plan) return null;
  return {
    ranges: plan.ranges.map((range) => ({ ...range })),
    cells: plan.cells.map((addr) => ({ ...addr })),
    exceptionCells: plan.exceptionCells.map((addr) => ({ ...addr })),
    axisBands: plan.axisBands.map((band) => ({ ...band })),
  };
}

export const selectionIncludes = (ranges: readonly Range[], addr: Addr): boolean =>
  ranges.some((range) => rangeContainsAddr(range, addr));

export interface SelectionFormatEntry {
  addr: Addr;
  patch: Partial<CellFormat> | null;
}

/** Disjoint selection ranges plus the already-authored cells (stored formats
 *  and the pending cell) they cover; blank cells stay implicit. */
export const sparseSelectionGeometry = (
  state: State,
): { ranges: Range[]; cells: Addr[] } | null => {
  const ranges = selectionRanges(state);
  if (!ranges || !mergeCoverageValid(state, ranges)) return null;
  const normalized = disjointRanges(ranges);
  const cells: Addr[] = [];
  for (const key of state.format.formats.keys()) {
    const addr = parseAddrKey(key);
    if (addr && selectionIncludes(normalized, addr)) cells.push(addr);
  }
  const pending = state.ui.pendingFormat?.addr;
  if (pending && selectionIncludes(normalized, pending)) cells.push(pending);
  return { ranges: normalized, cells };
};
