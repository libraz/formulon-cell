import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { addrKey } from '../engine/workbook-handle.js';
import { mutators, type SpreadsheetStore, type State } from '../store/store.js';
import { type History, recordMergesChangeWithEngine } from './history.js';

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

/** Expand `range` so it fully covers any merges that intersect it. Repeats
 *  until convergence (a newly-included merge can pull more cells into the
 *  range, which can pull more merges, etc.). */
export function expandRangeWithMerges(state: State, range: Range): Range {
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

/**
 * Merge `range` into a single visual cell. spreadsheet parity: the top-left value is
 * preserved; non-anchor cells are cleared. The clearing writes go through `wb`
 * (so they get individual undo entries via WorkbookHandle), and the
 * merges-state mutation gets one undo entry via `recordMergesChange`. Both are
 * wrapped in a single `history` transaction so Cmd+Z reverts the whole merge
 * in one step.
 *
 * Returns false on a 1×1 no-op range.
 */
export function applyMerge(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  range: Range,
): boolean {
  if (range.r0 === range.r1 && range.c0 === range.c1) return false;
  const sheet = range.sheet;

  if (history) history.begin();
  try {
    const state = store.getState();
    for (let r = range.r0; r <= range.r1; r += 1) {
      for (let c = range.c0; c <= range.c1; c += 1) {
        if (r === range.r0 && c === range.c0) continue;
        const cell = state.data.cells.get(addrKey({ sheet, row: r, col: c }));
        if (hasContent(cell)) wb.setBlank({ sheet, row: r, col: c });
      }
    }
    recordMergesChangeWithEngine(history, store, wb, sheet, () => {
      mutators.mergeRange(store, range);
    });
  } finally {
    if (history) history.end();
  }
  return true;
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
  const before = store.getState().merges.byAnchor;
  let touched = false;
  for (const r of before.values()) {
    if (r.sheet !== range.sheet) continue;
    if (r.r1 < range.r0 || r.r0 > range.r1 || r.c1 < range.c0 || r.c0 > range.c1) continue;
    touched = true;
    break;
  }
  if (!touched) return false;
  recordMergesChangeWithEngine(history, store, wb, range.sheet, () => {
    mutators.unmergeRange(store, range);
  });
  return true;
}
