import { expandRangeWithMerges, mergeAnchorOf, stepWithMerge } from '../commands/merge.js';
import { groupCols, groupRows, ungroupCols, ungroupRows } from '../commands/outline.js';
import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { addrKey } from '../engine/workbook-handle.js';
import { mutators, type SpreadsheetStore, type State } from '../store/store.js';

const MAX_ROW = 1_048_575; // Excel limit; clamp navigation.
const MAX_COL = 16_383;

const move = (a: Addr, dRow: number, dCol: number): Addr => ({
  sheet: a.sheet,
  row: Math.max(0, Math.min(MAX_ROW, a.row + dRow)),
  col: Math.max(0, Math.min(MAX_COL, a.col + dCol)),
});

const clamp = (a: Addr, row: number, col: number): Addr => ({
  sheet: a.sheet,
  row: Math.max(0, Math.min(MAX_ROW, row)),
  col: Math.max(0, Math.min(MAX_COL, col)),
});

const isPopulated = (s: State, sheet: number, row: number, col: number): boolean =>
  s.data.cells.has(addrKey({ sheet, row, col }));

/**
 * Excel-style Ctrl+Arrow jump. Three behaviors based on neighbor state:
 *  - empty origin            → jump to next populated cell (or sheet edge)
 *  - origin populated, adjacent populated → jump to last consecutive populated cell
 *  - origin populated, adjacent empty     → jump past the gap to next populated cell
 */
function jumpEdge(s: State, a: Addr, dRow: number, dCol: number): Addr {
  const sheet = a.sheet;
  const limitRow = dRow > 0 ? MAX_ROW : 0;
  const limitCol = dCol > 0 ? MAX_COL : 0;
  // Step 1: probe one step ahead.
  let r = a.row + dRow;
  let c = a.col + dCol;
  // Already at the edge — clamp.
  if (r < 0 || r > MAX_ROW || c < 0 || c > MAX_COL) return a;

  const originPop = isPopulated(s, sheet, a.row, a.col);
  const nextPop = isPopulated(s, sheet, r, c);

  if (originPop && nextPop) {
    // Walk until next cell is empty.
    while (
      r + dRow >= 0 &&
      r + dRow <= MAX_ROW &&
      c + dCol >= 0 &&
      c + dCol <= MAX_COL &&
      isPopulated(s, sheet, r + dRow, c + dCol)
    ) {
      r += dRow;
      c += dCol;
    }
    return clamp(a, r, c);
  }
  // Either origin is blank, or origin is populated but adjacent is blank.
  // Skip blanks to next populated cell.
  while (r >= 0 && r <= MAX_ROW && c >= 0 && c <= MAX_COL) {
    if (isPopulated(s, sheet, r, c)) return clamp(a, r, c);
    if ((dRow !== 0 && r === limitRow) || (dCol !== 0 && c === limitCol)) break;
    r += dRow;
    c += dCol;
  }
  // No populated cell found — go to sheet edge in that direction.
  return clamp(a, dRow !== 0 ? limitRow : a.row, dCol !== 0 ? limitCol : a.col);
}

/** Last populated cell on the active sheet — Ctrl+End target. */
function lastUsedCell(s: State, sheet: number): { row: number; col: number } {
  let maxRow = 0;
  let maxCol = 0;
  for (const key of s.data.cells.keys()) {
    const parts = key.split(':');
    if (parts.length !== 3) continue;
    if (Number(parts[0]) !== sheet) continue;
    const r = Number(parts[1]);
    const c = Number(parts[2]);
    if (r > maxRow) maxRow = r;
    if (c > maxCol) maxCol = c;
  }
  return { row: maxRow, col: maxCol };
}

export interface KeyboardDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  wb: WorkbookHandle;
  /** Optional shared history. When provided, Cmd/Ctrl+Z/Y route through
   *  this stack instead of the workbook's local one — so undo also reverts
   *  format and layout changes. */
  history?: import('../commands/history.js').History | null;
  /** Called when the user enters edit mode — playground / editor element
   *  consumes this to position the editor input. */
  onBeginEdit: (seed: string) => void;
  onClearActive: () => void;
  /** Called after Cmd/Ctrl+Z or Cmd/Ctrl+Y reached the workbook. The host
   *  needs to refresh its cached cell map. */
  onAfterHistory?: () => void;
  /** Called for F5 / Ctrl+G — Excel's "Go To" shortcut. The chrome layer
   *  decides whether to focus the Name Box, open a dialog, etc. */
  onGoTo?: () => void;
  /** Called for Shift+F2 — Excel's comment shortcut. The chrome layer is
   *  expected to wire this to the comment dialog feature. */
  onEditComment?: (addr: Addr) => void;
  /** Called for Ctrl/Cmd+PageUp/PageDown — Excel sheet-tab navigation. */
  onSwitchSheet?: (delta: 1 | -1) => void;
}

export function attachKeyboard(deps: KeyboardDeps): () => void {
  const { host, store } = deps;

  const onKey = (e: KeyboardEvent): void => {
    const s = store.getState();
    if (s.ui.editor.kind !== 'idle') return; // editor handles its own keys.

    const k = e.key;
    const meta = e.ctrlKey || e.metaKey;
    const shift = e.shiftKey;
    const a = s.selection.active;

    // Special-case F5 / Ctrl+G — Go To.
    if (k === 'F5' || (meta && (k === 'g' || k === 'G'))) {
      e.preventDefault();
      deps.onGoTo?.();
      return;
    }

    if (meta && (k === 'PageUp' || k === 'PageDown')) {
      e.preventDefault();
      deps.onSwitchSheet?.(k === 'PageDown' ? 1 : -1);
      return;
    }

    if (meta && (k === 'a' || k === 'A')) {
      e.preventDefault();
      mutators.selectAll(store);
      return;
    }

    if (meta && k === ' ') {
      e.preventDefault();
      mutators.selectCol(store, a.col);
      return;
    }

    if (shift && k === ' ' && !meta && !e.altKey) {
      e.preventDefault();
      mutators.selectRow(store, a.row);
      return;
    }

    // Alt+Shift+Right — group; Alt+Shift+Left — ungroup. Excel parity.
    // Axis pick: if the selection spans more rows than columns, group rows;
    // otherwise group columns. Excel prompts for this; we infer from shape.
    if (e.altKey && shift && (k === 'ArrowRight' || k === 'ArrowLeft')) {
      e.preventDefault();
      const range = s.selection.range;
      const rowSpan = range.r1 - range.r0;
      const colSpan = range.c1 - range.c0;
      const useRows = rowSpan >= colSpan;
      if (k === 'ArrowRight') {
        if (useRows) groupRows(store, deps.history ?? null, range.r0, range.r1);
        else groupCols(store, deps.history ?? null, range.c0, range.c1);
      } else {
        if (useRows) ungroupRows(store, deps.history ?? null, range.r0, range.r1);
        else ungroupCols(store, deps.history ?? null, range.c0, range.c1);
      }
      return;
    }

    // Shift+F2 — edit comment on active cell (Excel parity).
    if (k === 'F2' && shift) {
      e.preventDefault();
      deps.onEditComment?.(a);
      return;
    }

    // Compute the target address, then commit either as set-active or
    // extend-range based on Shift state.
    let target: Addr | null = null;

    if (k === 'ArrowUp')
      target = meta ? jumpEdge(s, a, -1, 0) : stepWithMerge(s, a, -1, 0, MAX_ROW, MAX_COL);
    else if (k === 'ArrowDown')
      target = meta ? jumpEdge(s, a, 1, 0) : stepWithMerge(s, a, 1, 0, MAX_ROW, MAX_COL);
    else if (k === 'ArrowLeft')
      target = meta ? jumpEdge(s, a, 0, -1) : stepWithMerge(s, a, 0, -1, MAX_ROW, MAX_COL);
    else if (k === 'ArrowRight')
      target = meta ? jumpEdge(s, a, 0, 1) : stepWithMerge(s, a, 0, 1, MAX_ROW, MAX_COL);
    else if (k === 'Home') target = meta ? clamp(a, 0, 0) : clamp(a, a.row, 0);
    else if (k === 'End' && meta) {
      const { row, col } = lastUsedCell(s, a.sheet);
      target = clamp(a, row, col);
    } else if (k === 'PageDown') target = move(a, Math.max(1, s.viewport.rowCount - 1), 0);
    else if (k === 'PageUp') target = move(a, -Math.max(1, s.viewport.rowCount - 1), 0);
    else if (k === 'Tab') target = stepWithMerge(s, a, 0, shift ? -1 : 1, MAX_ROW, MAX_COL);
    else if (k === 'Enter' && !meta) {
      target = stepWithMerge(s, a, shift ? -1 : 1, 0, MAX_ROW, MAX_COL);
    } else if (meta && k === 'Enter') {
      deps.onBeginEdit('');
      e.preventDefault();
      return;
    } else if (k === 'F2') {
      const f = deps.wb.cellFormula(a);
      const seed = f ?? formatExisting(s, a);
      deps.onBeginEdit(seed);
      e.preventDefault();
      return;
    } else if (k === 'Backspace' || k === 'Delete') {
      // Delete clears the entire selection range — Excel parity.
      const range = s.selection.range;
      const sheet = range.sheet;
      // Iterate populated cells only; full-sheet selection would otherwise loop 17B times.
      for (const key of s.data.cells.keys()) {
        const parts = key.split(':');
        if (parts.length !== 3) continue;
        if (Number(parts[0]) !== sheet) continue;
        const row = Number(parts[1]);
        const col = Number(parts[2]);
        if (row < range.r0 || row > range.r1) continue;
        if (col < range.c0 || col > range.c1) continue;
        deps.wb.setBlank({ sheet, row, col });
      }
      deps.onClearActive();
      e.preventDefault();
      return;
    } else if (k === 'Escape') {
      // No editor active; nothing to do.
      return;
    } else if (meta && (k === 'z' || k === 'Z')) {
      const h = deps.history;
      const ok = h ? (shift ? h.redo() : h.undo()) : shift ? deps.wb.redo() : deps.wb.undo();
      if (ok) deps.onAfterHistory?.();
      e.preventDefault();
      return;
    } else if (meta && (k === 'y' || k === 'Y')) {
      const h = deps.history;
      const ok = h ? h.redo() : deps.wb.redo();
      if (ok) deps.onAfterHistory?.();
      e.preventDefault();
      return;
    } else if (k.length === 1 && !meta && !e.altKey) {
      // Printable -> begin edit, seed with the typed char.
      deps.onBeginEdit(k);
      e.preventDefault();
      return;
    } else {
      return;
    }

    if (target == null) return;
    e.preventDefault();
    // Merge-aware: snap the active cell to the anchor and grow shift-extends so
    //  the selection always covers full merge rectangles.
    target = mergeAnchorOf(s, target);
    if (shift && k !== 'Tab') {
      mutators.extendRangeTo(store, target);
      const after = store.getState();
      const grown = expandRangeWithMerges(after, after.selection.range);
      if (
        grown.r0 !== after.selection.range.r0 ||
        grown.r1 !== after.selection.range.r1 ||
        grown.c0 !== after.selection.range.c0 ||
        grown.c1 !== after.selection.range.c1
      ) {
        mutators.setRange(store, grown);
      }
    } else {
      mutators.setActive(store, target);
    }
  };

  host.addEventListener('keydown', onKey);
  return () => host.removeEventListener('keydown', onKey);
}

function formatExisting(s: ReturnType<SpreadsheetStore['getState']>, a: Addr): string {
  const cell = s.data.cells.get(`${a.sheet}:${a.row}:${a.col}`);
  if (!cell) return '';
  if (cell.formula) return cell.formula;
  switch (cell.value.kind) {
    case 'number':
      return String(cell.value.value);
    case 'bool':
      return cell.value.value ? 'TRUE' : 'FALSE';
    case 'text':
      return cell.value.value;
    case 'error':
      return cell.value.text;
    default:
      return '';
  }
}
