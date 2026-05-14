import { addrKey } from '../engine/address.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SpreadsheetStore, State } from '../store/store.js';

/** Spreadsheet parity: refuse to sort when the range intersects any merge —
 *  rearranging rows would tear the merged rectangle apart. */
const rangeIntersectsMerges = (state: State, range: Range): boolean => {
  for (const m of state.merges.byAnchor.values()) {
    if (m.sheet !== range.sheet) continue;
    if (m.r1 < range.r0 || m.r0 > range.r1) continue;
    if (m.c1 < range.c0 || m.c0 > range.c1) continue;
    return true;
  }
  return false;
};

export type SortDirection = 'asc' | 'desc';

export interface SortOptions {
  /** Column to sort by, in absolute sheet coords. Must lie within `range`. */
  byCol: number;
  direction: SortDirection;
  /** When true, the first row is treated as a header and not moved. */
  hasHeader?: boolean;
}

/** Sort the rows of `range` in place by the values in `byCol`. Writes the
 *  resulting cells back through `wb`. Numbers come first (ascending) then
 *  text, then blanks — same as the spreadsheet's default. */
export function sortRange(
  state: State,
  _store: SpreadsheetStore,
  wb: WorkbookHandle,
  range: Range,
  opts: SortOptions,
): boolean {
  const start = opts.hasHeader ? range.r0 + 1 : range.r0;
  if (start > range.r1) return false;
  if (opts.byCol < range.c0 || opts.byCol > range.c1) return false;
  if (rangeIntersectsMerges(state, range)) return false;

  // Snapshot the rows we'll move, including formula text so we can write it back.
  interface RowSnap {
    cells: Array<{ value: unknown; formula: string | null; col: number } | null>;
    sortKey: { kind: 'number' | 'text' | 'blank'; n?: number; s?: string };
  }
  const snaps: { row: number; snap: RowSnap }[] = [];
  for (let r = start; r <= range.r1; r += 1) {
    const cells: RowSnap['cells'] = [];
    for (let c = range.c0; c <= range.c1; c += 1) {
      const key = addrKey({ sheet: range.sheet, row: r, col: c });
      const cell = state.data.cells.get(key);
      cells.push(cell ? { value: cell.value, formula: cell.formula, col: c } : null);
    }
    const keyCell = state.data.cells.get(addrKey({ sheet: range.sheet, row: r, col: opts.byCol }));
    let sortKey: RowSnap['sortKey'] = { kind: 'blank' };
    if (keyCell) {
      const v = keyCell.value;
      if (v.kind === 'number') sortKey = { kind: 'number', n: v.value };
      else if (v.kind === 'text') sortKey = { kind: 'text', s: v.value };
      else if (v.kind === 'bool') sortKey = { kind: 'number', n: v.value ? 1 : 0 };
    }
    snaps.push({ row: r, snap: { cells, sortKey } });
  }

  const dir = opts.direction === 'desc' ? -1 : 1;
  snaps.sort((a, b) => {
    const ka = a.snap.sortKey;
    const kb = b.snap.sortKey;
    // blanks sink to bottom regardless of dir (spreadsheet)
    if (ka.kind === 'blank' && kb.kind === 'blank') return 0;
    if (ka.kind === 'blank') return 1;
    if (kb.kind === 'blank') return -1;
    if (ka.kind === 'number' && kb.kind === 'number') return dir * ((ka.n ?? 0) - (kb.n ?? 0));
    if (ka.kind === 'text' && kb.kind === 'text')
      return dir * (ka.s ?? '').localeCompare(kb.s ?? '');
    // mixed types — numbers before text
    return ka.kind === 'number' ? -1 * dir : 1 * dir;
  });

  // Write back into wb in sorted order.
  for (let i = 0; i < snaps.length; i += 1) {
    const dstRow = start + i;
    const snap = snaps[i]?.snap;
    if (!snap) continue;
    for (const cell of snap.cells) {
      if (!cell) {
        wb.setBlank({ sheet: range.sheet, row: dstRow, col: 0 });
        continue;
      }
      const addr = { sheet: range.sheet, row: dstRow, col: cell.col };
      if (cell.formula) wb.setFormula(addr, cell.formula);
      else {
        const v = cell.value as { kind: string; value?: unknown };
        if (v.kind === 'number') wb.setNumber(addr, v.value as number);
        else if (v.kind === 'text') wb.setText(addr, v.value as string);
        else if (v.kind === 'bool') wb.setBool(addr, v.value as boolean);
        else wb.setBlank(addr);
      }
    }
  }
  wb.recalc();
  return true;
}

/** Remove duplicate rows from a range (keeping the first occurrence). */
export function removeDuplicates(
  state: State,
  _store: SpreadsheetStore,
  wb: WorkbookHandle,
  range: Range,
): number {
  if (rangeIntersectsMerges(state, range)) return 0;
  const seen = new Set<string>();
  const keep: number[] = [];
  for (let r = range.r0; r <= range.r1; r += 1) {
    const sig: string[] = [];
    for (let c = range.c0; c <= range.c1; c += 1) {
      const cell = state.data.cells.get(addrKey({ sheet: range.sheet, row: r, col: c }));
      if (!cell) sig.push('');
      else {
        const v = cell.value;
        if (v.kind === 'number') sig.push(`n:${v.value}`);
        else if (v.kind === 'text') sig.push(`t:${v.value}`);
        else if (v.kind === 'bool') sig.push(`b:${v.value ? 1 : 0}`);
        else sig.push('');
      }
    }
    const key = sig.join('');
    if (seen.has(key)) continue;
    seen.add(key);
    keep.push(r);
  }
  if (keep.length === range.r1 - range.r0 + 1) return 0;

  // Snapshot kept rows then rewrite from r0.
  interface RowSnap {
    cells: Array<{ value: unknown; formula: string | null; col: number } | null>;
  }
  const snaps: RowSnap[] = [];
  for (const r of keep) {
    const cells: RowSnap['cells'] = [];
    for (let c = range.c0; c <= range.c1; c += 1) {
      const cell = state.data.cells.get(addrKey({ sheet: range.sheet, row: r, col: c }));
      cells.push(cell ? { value: cell.value, formula: cell.formula, col: c } : null);
    }
    snaps.push({ cells });
  }

  for (let i = 0; i < snaps.length; i += 1) {
    const dstRow = range.r0 + i;
    const snap = snaps[i];
    if (!snap) continue;
    for (const cell of snap.cells) {
      if (!cell) continue;
      const addr = { sheet: range.sheet, row: dstRow, col: cell.col };
      if (cell.formula) wb.setFormula(addr, cell.formula);
      else {
        const v = cell.value as { kind: string; value?: unknown };
        if (v.kind === 'number') wb.setNumber(addr, v.value as number);
        else if (v.kind === 'text') wb.setText(addr, v.value as string);
        else if (v.kind === 'bool') wb.setBool(addr, v.value as boolean);
        else wb.setBlank(addr);
      }
    }
  }
  // Clear tail rows that were dropped.
  for (let r = range.r0 + snaps.length; r <= range.r1; r += 1) {
    for (let c = range.c0; c <= range.c1; c += 1) {
      wb.setBlank({ sheet: range.sheet, row: r, col: c });
    }
  }
  wb.recalc();
  return range.r1 - range.r0 + 1 - snaps.length;
}
