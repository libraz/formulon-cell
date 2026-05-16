import { addrKey } from '../engine/address.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { type CellFormat, mutators, type SpreadsheetStore, type State } from '../store/store.js';
import { inferAutoFilterRange } from './filter.js';
import type { History } from './history.js';
import { isCellWritable, warnProtected } from './protection.js';

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

const ensureWritableRange = (state: State, range: Range): boolean => {
  for (let r = range.r0; r <= range.r1; r += 1) {
    for (let c = range.c0; c <= range.c1; c += 1) {
      const addr = { sheet: range.sheet, row: r, col: c };
      if (!isCellWritable(state, addr)) {
        warnProtected(addr);
        return false;
      }
    }
  }
  return true;
};

const writeCellSnapshot = (
  wb: WorkbookHandle,
  addr: { sheet: number; row: number; col: number },
  cell: { value: unknown; formula: string | null } | null,
): void => {
  if (!cell) {
    wb.setBlank(addr);
    return;
  }
  if (cell.formula) {
    wb.setFormula(addr, cell.formula);
    return;
  }
  const v = cell.value as { kind: string; value?: unknown };
  if (v.kind === 'number') wb.setNumber(addr, v.value as number);
  else if (v.kind === 'text') wb.setText(addr, v.value as string);
  else if (v.kind === 'bool') wb.setBool(addr, v.value as boolean);
  else wb.setBlank(addr);
};

export type SortDirection = 'asc' | 'desc';

export interface SortKey {
  /** Column to sort by, in absolute sheet coords. Must lie within `range`. */
  byCol: number;
  direction: SortDirection;
}

export interface SortOptions {
  /** Column to sort by, in absolute sheet coords. Must lie within `range`. */
  byCol: number;
  direction: SortDirection;
  /** Additional Excel-style sort levels. When present, these keys are applied
   *  in order and `byCol`/`direction` are used only as a fallback. */
  keys?: readonly SortKey[];
  /** When true, the first row is treated as a header and not moved. */
  hasHeader?: boolean;
}

export interface RemoveDuplicatesOptions {
  /** Absolute sheet columns to compare. Defaults to every column in `range`. */
  columns?: readonly number[];
  /** When true, the first row is preserved and not compared as data. */
  hasHeader?: boolean;
}

const cloneFormat = (fmt: CellFormat | undefined): CellFormat | undefined =>
  fmt ? { ...fmt } : undefined;

const normalizedSortKeys = (opts: SortOptions): SortKey[] => {
  const keys = opts.keys?.length ? opts.keys : [{ byCol: opts.byCol, direction: opts.direction }];
  const out: SortKey[] = [];
  for (const key of keys) {
    if (!Number.isInteger(key.byCol)) continue;
    const direction = key.direction === 'desc' ? 'desc' : 'asc';
    out.push({ byCol: key.byCol, direction });
  }
  return out;
};

const sortKeyForCell = (
  state: State,
  sheet: number,
  row: number,
  col: number,
): { kind: 'number' | 'text' | 'blank'; n?: number; s?: string } => {
  const keyCell = state.data.cells.get(addrKey({ sheet, row, col }));
  if (!keyCell) return { kind: 'blank' };
  const v = keyCell.value;
  if (v.kind === 'number') return { kind: 'number', n: v.value };
  if (v.kind === 'text') return { kind: 'text', s: v.value };
  if (v.kind === 'bool') return { kind: 'number', n: v.value ? 1 : 0 };
  return { kind: 'blank' };
};

const compareSortKeys = (
  a: { kind: 'number' | 'text' | 'blank'; n?: number; s?: string },
  b: { kind: 'number' | 'text' | 'blank'; n?: number; s?: string },
  direction: SortDirection,
): number => {
  const dir = direction === 'desc' ? -1 : 1;
  // Blanks sink to bottom regardless of direction, matching spreadsheet sort.
  if (a.kind === 'blank' && b.kind === 'blank') return 0;
  if (a.kind === 'blank') return 1;
  if (b.kind === 'blank') return -1;
  if (a.kind === 'number' && b.kind === 'number') return dir * ((a.n ?? 0) - (b.n ?? 0));
  if (a.kind === 'text' && b.kind === 'text') return dir * (a.s ?? '').localeCompare(b.s ?? '');
  // Mixed types: numbers before text in ascending order.
  return a.kind === 'number' ? -1 * dir : 1 * dir;
};

const cellKindAt = (state: State, sheet: number, row: number, col: number): string => {
  const cell = state.data.cells.get(addrKey({ sheet, row, col }));
  return cell?.value.kind ?? 'blank';
};

const hasDistinctHeaderFormat = (
  state: State,
  sheet: number,
  headerRow: number,
  dataRow: number,
  col: number,
): boolean => {
  const header = state.format.formats.get(addrKey({ sheet, row: headerRow, col }));
  const data = state.format.formats.get(addrKey({ sheet, row: dataRow, col }));
  if (!header) return false;
  return (
    (header.bold === true && data?.bold !== true) ||
    (header.italic === true && data?.italic !== true) ||
    (header.underline === true && data?.underline !== true) ||
    (header.fill != null && header.fill !== data?.fill)
  );
};

/**
 * Conservative Excel-style header inference for one-click Sort A-Z / Z-A.
 * The old toolbar path treated every multi-row range as headered, which meant
 * plain numeric selections never moved their first row. We only infer a header
 * when the first row looks label-like and differs from the data row below.
 */
export function inferSortHasHeader(state: State, range: Range): boolean {
  if (range.r0 >= range.r1) return false;
  let firstRowText = 0;
  let firstRowNonBlank = 0;
  let comparableColumns = 0;
  let typeMismatch = 0;
  let distinctHeaderFormats = 0;

  for (let col = range.c0; col <= range.c1; col += 1) {
    const headerKind = cellKindAt(state, range.sheet, range.r0, col);
    if (headerKind !== 'blank') firstRowNonBlank += 1;
    if (headerKind === 'text') firstRowText += 1;

    const dataKind = cellKindAt(state, range.sheet, range.r0 + 1, col);
    if (headerKind !== 'blank' && dataKind !== 'blank') comparableColumns += 1;
    if (headerKind === 'text' && dataKind !== 'blank' && dataKind !== 'text') typeMismatch += 1;
    if (hasDistinctHeaderFormat(state, range.sheet, range.r0, range.r0 + 1, col)) {
      distinctHeaderFormats += 1;
    }
  }

  if (firstRowNonBlank === 0 || firstRowText === 0) return false;
  if (typeMismatch > 0) return true;
  return comparableColumns > 0 && distinctHeaderFormats > 0 && firstRowText >= comparableColumns;
}

/** Sort the rows of `range` in place by the values in `byCol`. Writes the
 *  resulting cells back through `wb`. Numbers come first (ascending) then
 *  text, then blanks — same as the spreadsheet's default. */
export function sortRange(
  state: State,
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  range: Range,
  opts: SortOptions,
): boolean {
  const start = opts.hasHeader ? range.r0 + 1 : range.r0;
  if (start > range.r1) return false;
  const sortKeys = normalizedSortKeys(opts);
  if (sortKeys.length === 0) return false;
  if (sortKeys.some((key) => key.byCol < range.c0 || key.byCol > range.c1)) return false;
  if (rangeIntersectsMerges(state, range)) return false;
  if (!ensureWritableRange(state, { ...range, r0: start })) return false;

  // Snapshot the rows we'll move, including formula text so we can write it back.
  interface RowSnap {
    cells: Array<{
      value: unknown;
      formula: string | null;
      col: number;
      format?: CellFormat;
    }>;
    sortKeys: Array<{ kind: 'number' | 'text' | 'blank'; n?: number; s?: string }>;
  }
  const snaps: { row: number; snap: RowSnap }[] = [];
  for (let r = start; r <= range.r1; r += 1) {
    const cells: RowSnap['cells'] = [];
    for (let c = range.c0; c <= range.c1; c += 1) {
      const key = addrKey({ sheet: range.sheet, row: r, col: c });
      const cell = state.data.cells.get(key);
      cells.push({
        value: cell?.value ?? { kind: 'blank' },
        formula: cell?.formula ?? null,
        col: c,
        format: cloneFormat(state.format.formats.get(key)),
      });
    }
    snaps.push({
      row: r,
      snap: {
        cells,
        sortKeys: sortKeys.map((key) => sortKeyForCell(state, range.sheet, r, key.byCol)),
      },
    });
  }

  snaps.sort((a, b) => {
    for (let i = 0; i < sortKeys.length; i += 1) {
      const result = compareSortKeys(
        a.snap.sortKeys[i] ?? { kind: 'blank' },
        b.snap.sortKeys[i] ?? { kind: 'blank' },
        sortKeys[i]?.direction ?? 'asc',
      );
      if (result !== 0) return result;
    }
    return a.row - b.row;
  });

  // Write back into wb in sorted order.
  for (let i = 0; i < snaps.length; i += 1) {
    const dstRow = start + i;
    const snap = snaps[i]?.snap;
    if (!snap) continue;
    for (const cell of snap.cells) {
      const addr = { sheet: range.sheet, row: dstRow, col: cell.col };
      writeCellSnapshot(wb, addr, cell);
    }
  }
  store.setState((s) => {
    const formats = new Map(s.format.formats);
    for (let r = start; r <= range.r1; r += 1) {
      for (let c = range.c0; c <= range.c1; c += 1) {
        formats.delete(addrKey({ sheet: range.sheet, row: r, col: c }));
      }
    }
    for (let i = 0; i < snaps.length; i += 1) {
      const dstRow = start + i;
      const snap = snaps[i]?.snap;
      if (!snap) continue;
      for (const cell of snap.cells) {
        if (!cell.format) continue;
        formats.set(addrKey({ sheet: range.sheet, row: dstRow, col: cell.col }), cell.format);
      }
    }
    return { ...s, format: { formats } };
  });
  wb.recalc();
  return true;
}

/** Remove duplicate rows from a range (keeping the first occurrence). */
export function removeDuplicates(
  state: State,
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  range: Range,
  options: RemoveDuplicatesOptions = {},
): number {
  if (rangeIntersectsMerges(state, range)) return 0;
  if (!ensureWritableRange(state, range)) return 0;
  const columns = (options.columns?.length ? options.columns : undefined)?.filter(
    (col, index, arr) =>
      col >= range.c0 && col <= range.c1 && Number.isInteger(col) && arr.indexOf(col) === index,
  );
  const keyColumns = columns?.length
    ? columns
    : Array.from({ length: range.c1 - range.c0 + 1 }, (_, i) => range.c0 + i);
  const firstDataRow = options.hasHeader ? range.r0 + 1 : range.r0;
  if (firstDataRow > range.r1) return 0;
  const seen = new Set<string>();
  const keep: number[] = options.hasHeader ? [range.r0] : [];
  for (let r = firstDataRow; r <= range.r1; r += 1) {
    const sig: string[] = [];
    for (const c of keyColumns) {
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
    cells: Array<{
      value: unknown;
      formula: string | null;
      col: number;
      format?: CellFormat;
    } | null>;
  }
  const snaps: RowSnap[] = [];
  for (const r of keep) {
    const cells: RowSnap['cells'] = [];
    for (let c = range.c0; c <= range.c1; c += 1) {
      const key = addrKey({ sheet: range.sheet, row: r, col: c });
      const cell = state.data.cells.get(key);
      cells.push(
        cell
          ? {
              value: cell.value,
              formula: cell.formula,
              col: c,
              format: cloneFormat(state.format.formats.get(key)),
            }
          : null,
      );
    }
    snaps.push({ cells });
  }

  for (let i = 0; i < snaps.length; i += 1) {
    const dstRow = range.r0 + i;
    const snap = snaps[i];
    if (!snap) continue;
    for (let offset = 0; offset < snap.cells.length; offset += 1) {
      const addr = { sheet: range.sheet, row: dstRow, col: range.c0 + offset };
      writeCellSnapshot(wb, addr, snap.cells[offset] ?? null);
    }
  }
  store.setState((s) => {
    const formats = new Map(s.format.formats);
    for (let r = range.r0; r <= range.r1; r += 1) {
      for (let c = range.c0; c <= range.c1; c += 1) {
        formats.delete(addrKey({ sheet: range.sheet, row: r, col: c }));
      }
    }
    for (let i = 0; i < snaps.length; i += 1) {
      const dstRow = range.r0 + i;
      const snap = snaps[i];
      if (!snap) continue;
      for (let offset = 0; offset < snap.cells.length; offset += 1) {
        const cell = snap.cells[offset];
        if (!cell?.format) continue;
        formats.set(
          addrKey({ sheet: range.sheet, row: dstRow, col: range.c0 + offset }),
          cell.format,
        );
      }
    }
    return { ...s, format: { formats } };
  });
  // Clear tail rows that were dropped.
  for (let r = range.r0 + snaps.length; r <= range.r1; r += 1) {
    for (let c = range.c0; c <= range.c1; c += 1) {
      wb.setBlank({ sheet: range.sheet, row: r, col: c });
    }
  }
  wb.recalc();
  return range.r1 - range.r0 + 1 - snaps.length;
}

export interface SortRangeWithHistoryDeps {
  store: SpreadsheetStore;
  workbook: WorkbookHandle;
  history: History;
  range: Range;
  options: SortOptions;
}

/** Wrap [[sortRange]] in a history transaction and refresh the affected
 *  sheet's cells when anything actually moved. Returns whether the sort
 *  changed any data, so callers can short-circuit follow-up side effects. */
export const sortRangeWithHistory = (deps: SortRangeWithHistoryDeps): boolean => {
  const { store, workbook, history, range, options } = deps;
  const state = store.getState();
  history.begin();
  let ok = false;
  try {
    ok = sortRange(state, store, workbook, range, options);
  } finally {
    history.end();
  }
  if (ok) mutators.replaceCells(store, workbook.cells(state.data.sheetIndex));
  return ok;
};

export interface SortActiveColumnAutoOptions {
  store: SpreadsheetStore;
  workbook: WorkbookHandle;
  history: History;
  direction: SortDirection;
}

/** "Sort A→Z / Z→A" toolbar action. Auto-detects the contiguous range around
 *  the selection, picks header behavior automatically, sorts by the active
 *  cell's column, and refreshes the sheet cells. Returns whether anything
 *  actually changed so the host can skip side effects on a no-op. */
export const sortActiveColumnAuto = (deps: SortActiveColumnAutoOptions): boolean => {
  const { store, workbook, history, direction } = deps;
  const state = store.getState();
  const range = inferAutoFilterRange(state);
  return sortRangeWithHistory({
    store,
    workbook,
    history,
    range,
    options: {
      byCol: state.selection.active.col,
      direction,
      hasHeader: inferSortHasHeader(state, range),
    },
  });
};
