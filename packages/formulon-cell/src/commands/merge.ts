import { addrKey, MAX_COL, MAX_ROW, parseAddrKey } from '../engine/address.js';
import type { Addr, CellValue, Range } from '../engine/types.js';
import { writeCell } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { rangeArea, rangeContainsAddr, rangesIntersect } from '../store/selection-geometry.js';
import {
  type CellBorderSide,
  type CellBorders,
  type CellFormat,
  mutators,
  type SpreadsheetStore,
  type State,
} from '../store/store.js';
import type { History } from './history.js';
import { isCellWritable, isSheetProtected, warnProtected } from './protection.js';
import { recordFormatChange, recordMergesChangeWithEngine } from './slice-history.js';

/** Look up the merge that covers `addr`, if any. Returns the full merge range
 *  (anchor on top-left, opposite corner on bottom-right). */
export function mergeAt(state: State, addr: Addr): Range | null {
  const ak = state.merges.byCell.get(addrKey(addr));
  if (ak) {
    const r = state.merges.byAnchor.get(ak);
    return r ?? null;
  }
  // The cell may itself be the anchor (anchors don't appear in `byCell`).
  const direct = state.merges.byAnchor.get(addrKey(addr));
  return direct ?? null;
}

/** If `addr` is inside a merge, return the merge anchor; otherwise return `addr`
 *  unchanged. Used for click-to-select and keyboard-into-merge: desktop spreadsheets always
 *  reports the anchor as the active cell when a merge is selected. */
export function mergeAnchorOf(state: State, addr: Addr): Addr {
  const m = mergeAt(state, addr);
  if (!m) return addr;
  return { sheet: m.sheet, row: m.r0, col: m.c0 };
}

const MAX_MERGE_CELLS = 100_000;

const isValidRangeCoordinates = (range: Range): boolean =>
  Number.isInteger(range.sheet) &&
  range.sheet >= 0 &&
  Number.isInteger(range.r0) &&
  Number.isInteger(range.r1) &&
  Number.isInteger(range.c0) &&
  Number.isInteger(range.c1) &&
  Number.isFinite(range.r0) &&
  Number.isFinite(range.r1) &&
  Number.isFinite(range.c0) &&
  Number.isFinite(range.c1) &&
  range.r0 >= 0 &&
  range.r1 >= 0 &&
  range.c0 >= 0 &&
  range.c1 >= 0 &&
  range.r0 <= MAX_ROW &&
  range.r1 <= MAX_ROW &&
  range.c0 <= MAX_COL &&
  range.c1 <= MAX_COL &&
  range.r0 <= range.r1 &&
  range.c0 <= range.c1;

/** Expand `range` so it fully covers any merges that intersect it. Repeats
 *  until convergence (a newly-included merge can pull more cells into the
 *  range, which can pull more merges, etc.). */
export function expandRangeWithMerges(state: State, range: Range): Range {
  if (!isValidRangeCoordinates(range)) return { ...range };
  let r0 = range.r0;
  let r1 = range.r1;
  let c0 = range.c0;
  let c1 = range.c1;
  let changed = true;
  while (changed) {
    changed = false;
    for (const m of state.merges.byAnchor.values()) {
      if (m.sheet !== range.sheet) continue;
      // Intersects?
      if (m.r1 < r0 || m.r0 > r1 || m.c1 < c0 || m.c0 > c1) continue;
      if (m.r0 < r0) {
        r0 = m.r0;
        changed = true;
      }
      if (m.r1 > r1) {
        r1 = m.r1;
        changed = true;
      }
      if (m.c0 < c0) {
        c0 = m.c0;
        changed = true;
      }
      if (m.c1 > c1) {
        c1 = m.c1;
        changed = true;
      }
    }
  }
  return { sheet: range.sheet, r0, c0, r1, c1 };
}

/** Compute the next address when stepping from `from` by (dRow, dCol). When the
 *  cursor sits inside a merge, exits in the move direction past the merge edge.
 *  When the step lands inside a merge, snaps to that merge's anchor. */
export function stepWithMerge(
  state: State,
  from: Addr,
  dRow: number,
  dCol: number,
  maxRow: number,
  maxCol: number,
): Addr {
  const here = mergeAt(state, from);
  let row = from.row;
  let col = from.col;
  if (here) {
    if (dRow > 0) row = here.r1;
    else if (dRow < 0) row = here.r0;
    if (dCol > 0) col = here.c1;
    else if (dCol < 0) col = here.c0;
  }
  let target: Addr = {
    sheet: from.sheet,
    row: Math.max(0, Math.min(maxRow, row + dRow)),
    col: Math.max(0, Math.min(maxCol, col + dCol)),
  };
  // Snap to anchor when target lands on a merge body.
  target = mergeAnchorOf(state, target);
  return target;
}

const hasContent = (
  cell: { value: { kind: string }; formula: string | null } | undefined,
): boolean => {
  if (!cell) return false;
  if (cell.formula) return true;
  return cell.value.kind !== 'blank';
};

const rangeWidth = (range: Range): number => range.c1 - range.c0 + 1;

const rangeHeight = (range: Range): number => range.r1 - range.r0 + 1;

const canMaterializeMergeRange = (range: Range): boolean =>
  isValidRangeCoordinates(range) && rangeArea(range) <= MAX_MERGE_CELLS;

const intersectingMerges = (state: State, range: Range): Range[] =>
  [...state.merges.byAnchor.values()].filter((merge) => rangesIntersect(merge, range));

const hasIntersectingTable = (state: State, range: Range): boolean =>
  state.tables.tables.some((table) => rangesIntersect(table.range, range));

type StoredCell = { value: CellValue; formula: string | null };

const contentCellsInRange = (
  state: State,
  range: Range,
): Array<{ addr: Addr; cell: StoredCell }> => {
  const cells: Array<{ addr: Addr; cell: StoredCell }> = [];
  for (const [key, cell] of state.data.cells) {
    const addr = parseAddrKey(key);
    if (!addr || !rangeContainsAddr(range, addr) || !hasContent(cell)) continue;
    cells.push({ addr, cell });
  }
  cells.sort((a, b) => a.addr.row - b.addr.row || a.addr.col - b.addr.col);
  return cells;
};

const sameAddr = (a: Addr, b: Addr): boolean =>
  a.sheet === b.sheet && a.row === b.row && a.col === b.col;

/** Visual fields that Excel carries from the upper-left cell to every cell
 *  covered by a merge. Perimeter and uniform diagonal borders are normalized
 *  separately below.
 *  Metadata such as comments, hyperlinks, validation and protection flags
 *  remains attached to each original cell. */
const mergeVisualKeys: readonly (keyof CellFormat)[] = [
  'cellStyle',
  'numFmt',
  'bold',
  'italic',
  'underline',
  'strike',
  'fontVertAlign',
  'align',
  'vAlign',
  'wrap',
  'justifyLastLine',
  'shrinkToFit',
  'indent',
  'rotation',
  'textDirection',
  'color',
  'fill',
  'fillPattern',
  'fillPatternColor',
  'fontFamily',
  'fontSize',
];

type MergeBorderSide = 'top' | 'right' | 'bottom' | 'left' | 'diagonalDown' | 'diagonalUp';

const borderSideSignature = (side: CellBorderSide | undefined): string => {
  if (side === undefined) return 'absent';
  if (typeof side === 'boolean') return `boolean:${side}`;
  return `style:${side.style};color:${side.color ?? ''}`;
};

const cloneBorderSide = (side: CellBorderSide): CellBorderSide =>
  typeof side === 'object' ? { ...side } : side;

const uniformBorderSide = (
  sides: readonly (CellBorderSide | undefined)[],
): CellBorderSide | undefined => {
  const first = sides[0];
  if (first === undefined || first === false) return undefined;
  const signature = borderSideSignature(first);
  if (sides.some((side) => borderSideSignature(side) !== signature)) return undefined;
  return cloneBorderSide(first);
};

const mergePerimeterSides = (
  state: State,
  range: Range,
): Record<MergeBorderSide, CellBorderSide | undefined> => {
  const at = (row: number, col: number): CellFormat | undefined =>
    state.format.formats.get(addrKey({ sheet: range.sheet, row, col }));
  const top: Array<CellBorderSide | undefined> = [];
  const bottom: Array<CellBorderSide | undefined> = [];
  const left: Array<CellBorderSide | undefined> = [];
  const right: Array<CellBorderSide | undefined> = [];
  const diagonalDown: Array<CellBorderSide | undefined> = [];
  const diagonalUp: Array<CellBorderSide | undefined> = [];
  for (let col = range.c0; col <= range.c1; col += 1) {
    top.push(at(range.r0, col)?.borders?.top);
    bottom.push(at(range.r1, col)?.borders?.bottom);
  }
  for (let row = range.r0; row <= range.r1; row += 1) {
    left.push(at(row, range.c0)?.borders?.left);
    right.push(at(row, range.c1)?.borders?.right);
  }
  for (let row = range.r0; row <= range.r1; row += 1) {
    for (let col = range.c0; col <= range.c1; col += 1) {
      const borders = at(row, col)?.borders;
      diagonalDown.push(borders?.diagonalDown);
      diagonalUp.push(borders?.diagonalUp);
    }
  }
  return {
    top: uniformBorderSide(top),
    right: uniformBorderSide(right),
    bottom: uniformBorderSide(bottom),
    left: uniformBorderSide(left),
    diagonalDown: uniformBorderSide(diagonalDown),
    diagonalUp: uniformBorderSide(diagonalUp),
  };
};

const normalizedMergeBorders = (
  current: CellBorders | undefined,
  range: Range,
  row: number,
  col: number,
  perimeter: Record<MergeBorderSide, CellBorderSide | undefined>,
): CellBorders | undefined => {
  const next: CellBorders = { ...(current ?? {}) };
  const sides: MergeBorderSide[] = ['top', 'right', 'bottom', 'left'];
  for (const side of sides) {
    const isOuter =
      (side === 'top' && row === range.r0) ||
      (side === 'right' && col === range.c1) ||
      (side === 'bottom' && row === range.r1) ||
      (side === 'left' && col === range.c0);
    const value = isOuter ? perimeter[side] : undefined;
    if (value === undefined) delete next[side];
    else next[side] = cloneBorderSide(value);
  }
  for (const side of ['diagonalDown', 'diagonalUp'] as const) {
    const value = perimeter[side];
    if (value === undefined) delete next[side];
    else next[side] = cloneBorderSide(value);
  }
  return Object.keys(next).length > 0 ? next : undefined;
};

const mergeVisualFormat = (
  target: CellFormat | undefined,
  anchor: CellFormat | undefined,
): CellFormat | undefined => {
  const next: CellFormat = { ...(target ?? {}) };
  for (const key of mergeVisualKeys) {
    const value = anchor?.[key];
    if (value === undefined) delete next[key];
    else Object.assign(next, { [key]: value });
  }
  return Object.keys(next).length > 0 ? next : undefined;
};

const sameFormat = (a: CellFormat | undefined, b: CellFormat | undefined): boolean =>
  JSON.stringify(a) === JSON.stringify(b);

const copyMergeVisualFormats = (store: SpreadsheetStore, range: Range): void => {
  const state = store.getState();
  const anchorFormat = state.format.formats.get(
    addrKey({ sheet: range.sheet, row: range.r0, col: range.c0 }),
  );
  const perimeter = mergePerimeterSides(state, range);
  const formats = new Map(state.format.formats);
  let changed = false;
  for (let row = range.r0; row <= range.r1; row += 1) {
    for (let col = range.c0; col <= range.c1; col += 1) {
      const addr = { sheet: range.sheet, row, col };
      const key = addrKey(addr);
      const current = state.format.formats.get(key);
      const nextVisual = mergeVisualFormat(current, anchorFormat);
      const next = { ...(nextVisual ?? {}) };
      const nextBorders = normalizedMergeBorders(current?.borders, range, row, col, perimeter);
      if (nextBorders === undefined) delete next.borders;
      else next.borders = nextBorders;
      const final = Object.keys(next).length > 0 ? next : undefined;
      if (sameFormat(current, final)) continue;
      changed = true;
      if (final === undefined) formats.delete(key);
      else formats.set(key, final);
    }
  }
  if (!changed) return;
  store.setState((s) => ({ ...s, format: { ...s.format, formats } }));
};

/**
 * Whether merging `range` would discard data. Excel promotes the first
 * nonblank cell in row-major order to the top-left anchor, so a merge warns
 * only when there are at least two populated cells in its effective range.
 * Existing merges are expanded before counting, matching the range that the
 * eventual merge command will actually replace.
 */
export function mergeWillLoseData(state: State, range: Range): boolean {
  if (!isValidRangeCoordinates(range)) return false;
  const effective = expandRangeWithMerges(state, range);
  return contentCellsInRange(state, effective).length > 1;
}

/** Whether Merge Across would discard data. Each row gets its own anchor, so
 *  the warning threshold is applied independently per row. */
export function mergeAcrossWillLoseData(state: State, range: Range): boolean {
  if (!isValidRangeCoordinates(range)) return false;
  const effective = expandRangeWithMerges(state, range);
  const counts = new Map<number, number>();
  for (const { addr } of contentCellsInRange(state, effective)) {
    const count = (counts.get(addr.row) ?? 0) + 1;
    if (count > 1) return true;
    counts.set(addr.row, count);
  }
  return false;
}

/**
 * Whether sheet protection should block a merge/unmerge over `range`. Merging
 * writes to (clears) the covered cells, so on a protected sheet the operation
 * is refused unless every cell in the range is writable — matching Excel, which
 * disables the merge controls on a protected sheet.
 */
function mergeBlockedByProtection(state: State, range: Range): boolean {
  const sheet = range.sheet;
  if (!isSheetProtected(state, sheet)) return false;
  if (!canMaterializeMergeRange(range)) return true;
  for (let r = range.r0; r <= range.r1; r += 1) {
    for (let c = range.c0; c <= range.c1; c += 1) {
      if (!isCellWritable(state, { sheet, row: r, col: c })) return true;
    }
  }
  return false;
}

/** A merge/unmerge selection may touch only one cell of an existing merge,
 *  but Excel's protection check covers the complete merged area. */
function unmergeBlockedByProtection(state: State, range: Range): boolean {
  if (!isSheetProtected(state, range.sheet)) return false;
  for (const merge of intersectingMerges(state, range)) {
    if (mergeBlockedByProtection(state, merge)) return true;
  }
  return false;
}

const mergeBlockedByTable = (state: State, range: Range): boolean =>
  hasIntersectingTable(state, range);

/**
 * Merge `range` into a single visual cell. Excel keeps the first nonblank cell
 * in row-major order, promoting it to the top-left anchor when necessary;
 * non-anchor cells are then cleared. The writes go through `wb` (so they get
 * individual undo entries via WorkbookHandle), and the merges-state mutation
 * gets one undo entry via `recordMergesChange`. Both are wrapped in a single
 * `history` transaction so Cmd+Z reverts the whole merge in one step.
 *
 * Returns false on a 1×1 no-op range.
 */
export function applyMerge(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  range: Range,
): boolean {
  const stateBefore = store.getState();
  if (!canMaterializeMergeRange(range)) return false;
  const effective = expandRangeWithMerges(stateBefore, range);
  if (effective.r0 === effective.r1 && effective.c0 === effective.c1) return false;
  if (!canMaterializeMergeRange(effective)) return false;
  if (mergeBlockedByTable(stateBefore, effective)) return false;
  const sheet = effective.sheet;

  // Sheet protection: refuse to clear locked cells (no silent backdoor).
  if (mergeBlockedByProtection(stateBefore, effective)) {
    warnProtected({ sheet, row: range.r0, col: range.c0 });
    return false;
  }

  if (history) history.begin();
  try {
    const content = contentCellsInRange(stateBefore, effective);
    const anchor = { sheet, row: effective.r0, col: effective.c0 };
    const first = content[0];
    // Formula text is copied verbatim: promoting a cell to the merge anchor does
    // not re-anchor its relative references.
    if (first && !sameAddr(first.addr, anchor)) {
      writeCell(wb, anchor, first.cell.value, first.cell.formula);
    }

    recordFormatChange(history, store, () => copyMergeVisualFormats(store, effective));

    for (let r = effective.r0; r <= effective.r1; r += 1) {
      for (let c = effective.c0; c <= effective.c1; c += 1) {
        if (r === effective.r0 && c === effective.c0) continue;
        const cell = stateBefore.data.cells.get(addrKey({ sheet, row: r, col: c }));
        if (hasContent(cell)) wb.setBlank({ sheet, row: r, col: c });
      }
    }
    recordMergesChangeWithEngine(history, store, wb, sheet, () => {
      mutators.mergeRange(store, effective);
    });
  } finally {
    if (history) history.end();
  }
  return true;
}

/** Merge every row of `range` independently. Existing merges intersecting the
 *  selection are removed once before any row merge is applied. The entire
 *  command is one history transaction, including typed promotion, format
 *  propagation, cell clearing, and merge state. */
export function applyMergeAcross(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  range: Range,
): boolean {
  const stateBefore = store.getState();
  if (!canMaterializeMergeRange(range)) return false;
  const effective = expandRangeWithMerges(stateBefore, range);
  if (rangeWidth(effective) <= 1 || rangeHeight(effective) <= 0) return false;
  if (!canMaterializeMergeRange(effective)) return false;
  if (mergeBlockedByTable(stateBefore, effective)) return false;
  if (mergeBlockedByProtection(stateBefore, effective)) {
    warnProtected({ sheet: effective.sheet, row: effective.r0, col: effective.c0 });
    return false;
  }

  if (history) history.begin();
  try {
    if (intersectingMerges(stateBefore, effective).length > 0) {
      if (!applyUnmerge(store, wb, history, effective)) return false;
    }
    let applied = true;
    for (let row = effective.r0; row <= effective.r1; row += 1) {
      const rowRange: Range = {
        sheet: effective.sheet,
        r0: row,
        c0: effective.c0,
        r1: row,
        c1: effective.c1,
      };
      if (!applyMerge(store, wb, history, rowRange)) applied = false;
    }
    return applied;
  } finally {
    if (history) history.end();
  }
}

/**
 * Remove every merge that intersects `range`. The cells stay as they are —
 * Spreadsheets keep the (single) anchor value visible in the top-left after split.
 * `wb` may be null in entry points that don't have an engine handle (e.g. the
 * paste path runs before the engine has been attached) — engine sync is
 * skipped in that case.
 */
export function applyUnmerge(
  store: SpreadsheetStore,
  wb: WorkbookHandle | null,
  history: History | null,
  range: Range,
): boolean {
  const state = store.getState();
  const before = state.merges.byAnchor;
  let touched = false;
  for (const r of before.values()) {
    if (r.sheet !== range.sheet) continue;
    if (r.r1 < range.r0 || r.r0 > range.r1 || r.c1 < range.c0 || r.c0 > range.c1) continue;
    touched = true;
    break;
  }
  if (!touched) return false;
  // Sheet protection: unmerging is a structural edit — refuse on locked cells.
  if (unmergeBlockedByProtection(state, range)) {
    warnProtected({ sheet: range.sheet, row: range.r0, col: range.c0 });
    return false;
  }
  recordMergesChangeWithEngine(history, store, wb, range.sheet, () => {
    mutators.unmergeRange(store, range);
  });
  return true;
}
