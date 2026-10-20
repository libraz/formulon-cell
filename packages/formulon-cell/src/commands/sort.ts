import { addrKey, MAX_COL } from '../engine/address.js';
import { createProfileTextSemantics, localeForWorkbook } from '../engine/profile-text-semantics.js';
import type { Addr, CellValue, Range } from '../engine/types.js';
import { fromEngineValue, writeCell } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { rangeArea } from '../store/selection-geometry.js';
import { type CellFormat, mutators, type SpreadsheetStore, type State } from '../store/store.js';
import { customListsFor } from './autofill-series.js';
import { recordCommentChange } from './comment.js';
import { inferAutoFilterRange } from './filter.js';
import { shiftFormulaRefs } from './formula-refs.js';
import type { History } from './history.js';
import { isCellWritable, warnProtected } from './protection.js';
import { recordFormatChange } from './slice-history.js';

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
  addr: Addr,
  cell: { value: CellValue; formula: string | null } | null,
  sourceRow = addr.row,
): void => {
  if (!cell) {
    wb.setBlank(addr);
    return;
  }
  const formula = cell.formula ? shiftFormulaRefs(cell.formula, addr.row - sourceRow, 0) : null;
  writeCell(wb, addr, cell.value, formula);
};

export type SortDirection = 'asc' | 'desc';

export interface SortKey {
  /** Column to sort by, in absolute sheet coords. Must lie within `range`. */
  byCol: number;
  direction: SortDirection;
  /** Sort by values by default, or by cell/font color when specified. */
  sortOn?: 'value' | 'cellColor' | 'fontColor' | 'customList';
  /** Target color for color sorts. Matching rows sort first for asc, last for desc. */
  color?: string;
  /** Custom list order for `sortOn: "customList"`. Built-in lists are used when omitted. */
  customList?: readonly string[];
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

/** Bring native-only notes into the format snapshot before sorting. A workbook
 * can be edited through the engine API before the UI store is hydrated; those
 * notes still have to move with their rows and participate in history. */
const hydrateSortComments = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  range: Range,
  startRow: number,
): State => {
  if (!wb.capabilities.comments) return store.getState();
  const comments = wb.capabilities.commentsEnumerable
    ? wb.getComments(range.sheet)
    : Array.from(wb.cells(range.sheet)).flatMap((entry) => {
        const comment = wb.getComment(range.sheet, entry.addr.row, entry.addr.col);
        return comment
          ? [
              {
                row: entry.addr.row,
                col: entry.addr.col,
                author: comment.author,
                text: comment.text,
              },
            ]
          : [];
      });
  const inRange = comments.filter(
    (comment) =>
      comment.row >= startRow &&
      comment.row <= range.r1 &&
      comment.col >= range.c0 &&
      comment.col <= range.c1,
  );
  if (inRange.length === 0) return store.getState();
  const current = store.getState();
  const formats = new Map(current.format.formats);
  let changed = false;
  for (const comment of inRange) {
    const key = addrKey({ sheet: range.sheet, row: comment.row, col: comment.col });
    const previous = formats.get(key) ?? {};
    if (previous.comment === comment.text && previous.commentAuthor === comment.author) continue;
    formats.set(key, { ...previous, comment: comment.text, commentAuthor: comment.author });
    changed = true;
  }
  if (!changed) return current;
  store.setState((s) => ({ ...s, format: { ...s.format, formats } }));
  return store.getState();
};

const MAX_SORT_CELLS = 100_000;
// Native SORTBY is n log n and remained sub-100ms at 75k rows in the
// bundled-engine probe. Keep a separate row budget so a future engine
// regression cannot synchronously evaluate an unbounded dynamic array.
const MAX_NATIVE_SORT_ROWS = 50_000;
// UNIQUE in the bundled engine grows much faster than SORTBY: 10k rows took
// ~0.3s, 20k ~1.5s, 50k ~7s, and 75k ~23s in the local probe. Mutating
// commands reject the native route above this budget before evaluating.
const MAX_NATIVE_UNIQUE_ROWS = 20_000;

const sortArea = (range: Range, startRow: number): number =>
  Math.max(0, range.r1 - startRow + 1) * (range.c1 - range.c0 + 1);

const canRewriteExactRange = (range: Range, startRow = range.r0): boolean =>
  sortArea(range, startRow) <= MAX_SORT_CELLS;

const columnName = (column: number): string => {
  let n = column + 1;
  let out = '';
  while (n > 0) {
    const digit = (n - 1) % 26;
    out = String.fromCharCode(65 + digit) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
};

const absoluteColumnRange = (column: number, firstRow: number, lastRow: number): string => {
  const name = columnName(column);
  return `$${name}$${firstRow + 1}:$${name}$${lastRow + 1}`;
};

const verticalNumberArray = (values: readonly number[]): string => `{${values.join(';')}}`;

const nativeEvaluationAnchor = (state: State, range: Range): Addr => {
  // evaluateFormulaArray is read-only, but the engine still rejects a spill
  // whose anchor cell already carries a value. Pick an empty cell on the same
  // sheet; the formula itself uses only absolute references, so the anchor
  // does not change its meaning. State hydration makes this scan cheap in the
  // normal case and keeps it bounded for sparse large sheets.
  const columns = Math.min(MAX_COL, Math.max(range.c1 + 8, 64));
  for (let row = 0; row <= Math.min(range.r0 + 1, 64); row += 1) {
    for (let col = 0; col <= columns; col += 1) {
      if (!state.data.cells.has(addrKey({ sheet: range.sheet, row, col }))) {
        return { sheet: range.sheet, row, col };
      }
    }
  }
  return { sheet: range.sheet, row: 0, col: 0 };
};

/**
 * Ask the real formulon bundle for the row permutation. Dynamic-array
 * evaluation is read-only, so the temporary SORTBY formula never creates a
 * scratch cell or changes recalc state. The caller must treat a malformed
 * result as a hard failure: a native workbook should not silently fall back
 * to a browser comparator after a native operation was attempted.
 */
const nativeSortPermutation = (
  state: State,
  wb: WorkbookHandle,
  range: Range,
  startRow: number,
  sortKeys: readonly SortKey[],
): number[] | null => {
  const count = range.r1 - startRow + 1;
  if (count <= 0) return [];
  if (count > MAX_NATIVE_SORT_ROWS) return null;

  const builtInLists = customListsFor(localeForWorkbook(wb));
  const byKey = sortKeys.map((key) => {
    const direction = key.direction === 'desc' ? -1 : 1;
    const sortOn = key.sortOn ?? 'value';
    if (sortOn === 'value') {
      return `${absoluteColumnRange(key.byCol, startRow, range.r1)},${direction}`;
    }
    const ranks: number[] = [];
    for (let row = startRow; row <= range.r1; row += 1) {
      const sortValue = sortKeyForCell(state, range.sheet, row, key, builtInLists);
      if (sortOn === 'cellColor' || sortOn === 'fontColor') {
        ranks.push(sortValue.match === true ? 0 : 1);
      } else {
        // Unknown custom-list entries remain after all listed entries in both
        // directions, matching the existing JS path's explicit sentinel.
        ranks.push(sortValue.customIndex ?? sortKeys.length + count);
      }
    }
    return `${verticalNumberArray(ranks)},${direction}`;
  });
  const formula = `=SORTBY(SEQUENCE(${count}),${byKey.join(',')})`;
  try {
    const result = wb.evaluateFormulaArray(nativeEvaluationAnchor(state, range), formula);
    if (!result.status.ok || result.rows !== count || result.cols !== 1) return null;
    if (result.cells.length !== count) return null;
    const permutation: number[] = [];
    const seen = new Set<number>();
    for (const row of result.cells) {
      if (row?.length !== 1) return null;
      const raw = row?.[0];
      if (!raw) return null;
      const value = fromEngineValue(raw);
      if (value.kind !== 'number' || !Number.isInteger(value.value)) return null;
      const index = value.value;
      if (index < 1 || index > count || seen.has(index)) return null;
      seen.add(index);
      permutation.push(index);
    }
    return permutation.length === count ? permutation : null;
  } catch {
    return null;
  }
};

/** Key used only while recovering representatives from a native UNIQUE
 * result. The engine's dynamic-array ABI exposes both a source blank and a
 * numeric zero as numeric zero; keep them in one queue so a repeated native
 * zero can still recover the source rows in first-occurrence order. */
const nativeUniqueValueKey = (value: CellValue): string => {
  switch (value.kind) {
    case 'blank':
      return 'zero-or-blank';
    case 'number':
      return value.value === 0 ? 'zero-or-blank' : `number:${String(value.value)}`;
    case 'text':
      return `text:${value.value}`;
    case 'bool':
      return `bool:${value.value ? '1' : '0'}`;
    case 'error':
      return `error:${value.code}:${value.text}`;
  }
};

const nativeUniqueRowKey = (values: readonly CellValue[]): string =>
  JSON.stringify(values.map((value) => nativeUniqueValueKey(value)));

const nativeUniqueRows = (
  state: State,
  wb: WorkbookHandle,
  range: Range,
  firstDataRow: number,
  keyColumns: readonly number[],
): number[] | null => {
  const count = range.r1 - firstDataRow + 1;
  if (count <= 0 || keyColumns.length === 0) return [];
  if (count > MAX_NATIVE_UNIQUE_ROWS) return null;
  const startRef = absoluteColumnRange(range.c0, firstDataRow, range.r1);
  const endRef = absoluteColumnRange(range.c1, firstDataRow, range.r1).split(':')[1];
  const sourceRange = `${startRef.split(':')[0]}:${endRef}`;
  const relativeColumns = keyColumns.map((column) => column - range.c0 + 1);
  const keyFormula =
    relativeColumns.length === range.c1 - range.c0 + 1
      ? sourceRange
      : `CHOOSECOLS(${sourceRange},${relativeColumns.join(',')})`;
  try {
    const result = wb.evaluateFormulaArray(
      nativeEvaluationAnchor(state, range),
      `=UNIQUE(${keyFormula})`,
    );
    if (!result.status.ok || result.cols !== keyColumns.length || result.rows < 1) return null;
    if (result.rows > count || result.cells.length !== result.rows) return null;
    const uniqueValues = result.cells.map((row) => {
      if (row.length !== keyColumns.length) return null;
      return row.map((value) => fromEngineValue(value));
    });
    if (uniqueValues.some((row): row is null => row === null)) return null;

    const candidates = new Map<string, number[]>();
    for (let row = firstDataRow; row <= range.r1; row += 1) {
      const sourceKey = nativeUniqueRowKey(
        keyColumns.map(
          (column) =>
            state.data.cells.get(addrKey({ sheet: range.sheet, row, col: column }))?.value ??
            ({ kind: 'blank' } as const),
        ),
      );
      const rows = candidates.get(sourceKey);
      if (rows) rows.push(row);
      else candidates.set(sourceKey, [row]);
    }
    const offsets = new Map<string, number>();
    const representativeRows: number[] = [];
    for (const values of uniqueValues) {
      if (!values) return null;
      const outputKey = nativeUniqueRowKey(values);
      const rows = candidates.get(outputKey);
      const offset = offsets.get(outputKey) ?? 0;
      const found = rows?.[offset];
      if (found === undefined) return null;
      offsets.set(outputKey, offset + 1);
      representativeRows.push(found);
    }
    return representativeRows;
  } catch {
    return null;
  }
};

const normalizedSortKeys = (opts: SortOptions): SortKey[] => {
  const keys = opts.keys?.length ? opts.keys : [{ byCol: opts.byCol, direction: opts.direction }];
  const out: SortKey[] = [];
  for (const key of keys) {
    if (!Number.isInteger(key.byCol)) continue;
    const direction = key.direction === 'desc' ? 'desc' : 'asc';
    const sortOn = key.sortOn ?? 'value';
    out.push({
      byCol: key.byCol,
      direction,
      ...(sortOn !== 'value' ? { sortOn } : {}),
      ...(key.color ? { color: key.color } : {}),
      ...(key.customList ? { customList: [...key.customList] } : {}),
    });
  }
  return out;
};

const normalizeColorKey = (color: string): string => color.trim().toLocaleLowerCase();
const normalizeListKey = (value: string): string => value.trim().toLocaleLowerCase();

type SortValueKind = 'number' | 'text' | 'bool' | 'error' | 'blank' | 'color' | 'custom';

const customListIndex = (value: string, lists: readonly (readonly string[])[]): number | null => {
  const key = normalizeListKey(value);
  for (const candidate of lists) {
    const idx = candidate.findIndex((item) => normalizeListKey(item) === key);
    if (idx >= 0) return idx;
  }
  return null;
};

const sortKeyForCell = (
  state: State,
  sheet: number,
  row: number,
  key: SortKey,
  builtInLists: readonly (readonly string[])[],
): {
  kind: SortValueKind;
  n?: number;
  s?: string;
  match?: boolean;
  customIndex?: number | null;
} => {
  const sortOn = key.sortOn ?? 'value';
  if (sortOn === 'cellColor' || sortOn === 'fontColor') {
    const fmt = state.format.formats.get(addrKey({ sheet, row, col: key.byCol }));
    const actual = sortOn === 'cellColor' ? fmt?.fill : fmt?.color;
    return {
      kind: 'color',
      match: !!actual && !!key.color && normalizeColorKey(actual) === normalizeColorKey(key.color),
    };
  }
  const col = key.byCol;
  const keyCell = state.data.cells.get(addrKey({ sheet, row, col }));
  if (!keyCell) return { kind: 'blank' };
  const v = keyCell.value;
  if (sortOn === 'customList') {
    const value = v.kind === 'text' ? v.value : v.kind === 'number' ? String(v.value) : '';
    return {
      kind: 'custom',
      customIndex: value
        ? customListIndex(value, key.customList ? [key.customList] : builtInLists)
        : null,
    };
  }
  if (v.kind === 'number') return { kind: 'number', n: v.value };
  if (v.kind === 'text') return { kind: 'text', s: v.value };
  if (v.kind === 'bool') return { kind: 'bool', n: v.value ? 1 : 0 };
  if (v.kind === 'error') return { kind: 'error' };
  return { kind: 'blank' };
};

const compareSortKeys = (
  a: {
    kind: SortValueKind;
    n?: number;
    s?: string;
    match?: boolean;
    customIndex?: number | null;
  },
  b: {
    kind: SortValueKind;
    n?: number;
    s?: string;
    match?: boolean;
    customIndex?: number | null;
  },
  direction: SortDirection,
  compareText: (left: string, right: string) => number,
): number => {
  const dir = direction === 'desc' ? -1 : 1;
  if (a.kind === 'color' || b.kind === 'color') {
    const av = a.kind === 'color' && a.match === true ? 0 : 1;
    const bv = b.kind === 'color' && b.match === true ? 0 : 1;
    return dir * (av - bv);
  }
  if (a.kind === 'custom' || b.kind === 'custom') {
    const av =
      a.kind === 'custom' && a.customIndex != null ? a.customIndex : Number.POSITIVE_INFINITY;
    const bv =
      b.kind === 'custom' && b.customIndex != null ? b.customIndex : Number.POSITIVE_INFINITY;
    return dir * (av - bv);
  }
  // Blanks sink to bottom regardless of direction, matching spreadsheet sort.
  if (a.kind === 'blank' && b.kind === 'blank') return 0;
  if (a.kind === 'blank') return 1;
  if (b.kind === 'blank') return -1;
  const rank: Record<'number' | 'text' | 'bool' | 'error', number> = {
    number: 0,
    text: 1,
    bool: 2,
    error: 3,
  };
  if (a.kind !== b.kind) return dir * (rank[a.kind] - rank[b.kind]);
  switch (a.kind) {
    case 'number':
    case 'bool':
      return dir * ((a.n ?? 0) - (b.n ?? 0));
    case 'text':
      return dir * compareText(a.s ?? '', b.s ?? '');
    case 'error':
      // Errors retain their source order, matching Excel's stable sort.
      return 0;
    default:
      return 0;
  }
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
  let bothText = 0;

  for (let col = range.c0; col <= range.c1; col += 1) {
    const headerKind = cellKindAt(state, range.sheet, range.r0, col);
    if (headerKind !== 'blank') firstRowNonBlank += 1;
    if (headerKind === 'text') firstRowText += 1;

    const dataKind = cellKindAt(state, range.sheet, range.r0 + 1, col);
    if (headerKind !== 'blank' && dataKind !== 'blank') comparableColumns += 1;
    if (headerKind === 'text' && dataKind !== 'blank' && dataKind !== 'text') typeMismatch += 1;
    if (headerKind === 'text' && dataKind === 'text') bothText += 1;
    if (hasDistinctHeaderFormat(state, range.sheet, range.r0, range.r0 + 1, col)) {
      distinctHeaderFormats += 1;
    }
  }

  if (firstRowNonBlank === 0 || firstRowText === 0) return false;
  if (typeMismatch > 0) return true;
  if (distinctHeaderFormats > 0 && firstRowText >= comparableColumns) return true;
  // All-text table with no numeric or format contrast: the spreadsheet's Sort
  // dialog still defaults to "my data has headers", so keep the label row out of
  // the sort rather than mixing it into the data.
  return (
    comparableColumns > 0 && bothText === comparableColumns && firstRowText >= comparableColumns
  );
}

/** Sort the rows of `range` in place by the values in `byCol`. Writes the
 *  resulting cells back through `wb`. Ascending value order is number, text,
 *  bool, error, blank; blanks stay last for either direction. */
export function sortRange(
  state: State,
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  range: Range,
  opts: SortOptions,
  history: History | null = null,
): boolean {
  const start = opts.hasHeader ? range.r0 + 1 : range.r0;
  if (start > range.r1) return false;
  const sortKeys = normalizedSortKeys(opts);
  if (sortKeys.length === 0) return false;
  if (sortKeys.some((key) => key.byCol < range.c0 || key.byCol > range.c1)) return false;
  if (!canRewriteExactRange(range, start)) return false;
  if (rangeIntersectsMerges(state, range)) return false;
  if (!ensureWritableRange(state, { ...range, r0: start })) return false;

  // A real formulon bundle owns the comparison semantics for formula-driven
  // sorts. Do this before hydrating comments or touching any cell so an
  // unsupported/malformed native result is an atomic no-op. Stub workbooks
  // intentionally use the explicit-locale JS comparator below.
  const nativeBundle = wb.isStub === false && wb.capabilities.arrayFormulaEvaluation === true;
  const nativeOrder = nativeBundle
    ? nativeSortPermutation(state, wb, range, start, sortKeys)
    : null;
  if (nativeBundle && nativeOrder === null) return false;

  state = hydrateSortComments(store, wb, range, start);
  const builtInLists = customListsFor(localeForWorkbook(wb));

  // Snapshot the rows we'll move, including formula text so we can write it back.
  interface RowSnap {
    cells: Array<{
      value: CellValue;
      formula: string | null;
      col: number;
      format?: CellFormat;
    }>;
    sortKeys: Array<{
      kind: SortValueKind;
      n?: number;
      s?: string;
      match?: boolean;
      customIndex?: number | null;
    }>;
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
        sortKeys: sortKeys.map((key) => sortKeyForCell(state, range.sheet, r, key, builtInLists)),
      },
    });
  }

  if (nativeOrder) {
    const bySourceIndex = new Map(snaps.map((entry, index) => [entry.row, index]));
    const ordered = nativeOrder.map((sourceIndex) => snaps[sourceIndex - 1]);
    // nativeSortPermutation validates the sequence, but retain a defensive
    // check at this boundary so a future result-shape change still cannot
    // cause a partial write.
    if (ordered.some((entry) => !entry) || bySourceIndex.size !== snaps.length) return false;
    snaps.splice(
      0,
      snaps.length,
      ...ordered.filter((entry): entry is (typeof snaps)[number] => !!entry),
    );
  } else {
    const textSemantics = createProfileTextSemantics(localeForWorkbook(wb));
    snaps.sort((a, b) => {
      for (let i = 0; i < sortKeys.length; i += 1) {
        const result = compareSortKeys(
          a.snap.sortKeys[i] ?? { kind: 'blank' },
          b.snap.sortKeys[i] ?? { kind: 'blank' },
          sortKeys[i]?.direction ?? 'asc',
          textSemantics.compare,
        );
        if (result !== 0) return result;
      }
      return a.row - b.row;
    });
  }

  // Write back into wb in sorted order.
  wb.withBatchedRecalc(() => {
    for (let i = 0; i < snaps.length; i += 1) {
      const dstRow = start + i;
      const snap = snaps[i]?.snap;
      if (!snap) continue;
      for (const cell of snap.cells) {
        const addr = { sheet: range.sheet, row: dstRow, col: cell.col };
        writeCellSnapshot(wb, addr, cell, snaps[i]?.row ?? dstRow);
      }
    }
  });
  const affected: Addr[] = [];
  for (let r = start; r <= range.r1; r += 1) {
    for (let c = range.c0; c <= range.c1; c += 1) {
      affected.push({ sheet: range.sheet, row: r, col: c });
    }
  }
  const commentTargets = new Map<string, Addr>();
  const rememberComment = (addr: Addr, format: CellFormat | undefined): void => {
    if (typeof format?.comment !== 'string' || format.comment.length === 0) return;
    commentTargets.set(addrKey(addr), addr);
  };
  for (const addr of affected) {
    rememberComment(addr, state.format.formats.get(addrKey(addr)));
  }
  for (let i = 0; i < snaps.length; i += 1) {
    const dstRow = start + i;
    const snap = snaps[i]?.snap;
    if (!snap) continue;
    for (const cell of snap.cells) {
      rememberComment({ sheet: range.sheet, row: dstRow, col: cell.col }, cell.format);
    }
  }
  const applyFormats = (): void => {
    store.setState((s) => {
      const formats = new Map(s.format.formats);
      for (const addr of affected) formats.delete(addrKey(addr));
      for (let i = 0; i < snaps.length; i += 1) {
        const dstRow = start + i;
        const snap = snaps[i]?.snap;
        if (!snap) continue;
        for (const cell of snap.cells) {
          if (!cell.format) continue;
          formats.set(addrKey({ sheet: range.sheet, row: dstRow, col: cell.col }), cell.format);
        }
      }
      return { ...s, format: { ...s.format, formats } };
    });
    for (const addr of commentTargets.values()) {
      const format = store.getState().format.formats.get(addrKey(addr));
      wb.setCommentEntry(
        addr.sheet,
        addr.row,
        addr.col,
        format?.commentAuthor ?? '',
        format?.comment ?? '',
      );
    }
  };
  recordFormatChange(history, store, () => {
    recordCommentChange(history, store, wb, [...commentTargets.values()], applyFormats);
  });
  wb.recalcAuto();
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
  if (rangeArea(range) > MAX_SORT_CELLS) return 0;
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
  const nativeBundle = wb.isStub === false && wb.capabilities.arrayFormulaEvaluation === true;
  const nativeKeep = nativeBundle
    ? nativeUniqueRows(state, wb, range, firstDataRow, keyColumns)
    : null;
  if (nativeBundle && nativeKeep === null) return 0;
  const keep: number[] = options.hasHeader ? [range.r0] : [];
  if (nativeKeep) keep.push(...nativeKeep);
  if (!nativeKeep) {
    const semantics = createProfileTextSemantics(localeForWorkbook(wb));
    const seen = new Set<string>();
    for (let r = firstDataRow; r <= range.r1; r += 1) {
      const sig: string[] = [];
      for (const c of keyColumns) {
        const cell = state.data.cells.get(addrKey({ sheet: range.sheet, row: r, col: c }));
        if (!cell) sig.push('blank:');
        else {
          const v = cell.value;
          if (v.kind === 'number') sig.push(`number:${v.value}`);
          else if (v.kind === 'text') sig.push(`text:${semantics.key(v.value)}`);
          else if (v.kind === 'bool') sig.push(`bool:${v.value ? 1 : 0}`);
          else if (v.kind === 'error') sig.push(`error:${v.code}:${v.text}`);
          else sig.push('blank:');
        }
      }
      const key = sig.join('');
      if (seen.has(key)) continue;
      seen.add(key);
      keep.push(r);
    }
  }
  if (keep.length === range.r1 - range.r0 + 1) return 0;

  // Snapshot kept rows then rewrite from r0.
  interface RowSnap {
    cells: Array<{
      value: CellValue;
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

  wb.withBatchedRecalc(() => {
    for (let i = 0; i < snaps.length; i += 1) {
      const dstRow = range.r0 + i;
      const snap = snaps[i];
      if (!snap) continue;
      for (let offset = 0; offset < snap.cells.length; offset += 1) {
        const addr = { sheet: range.sheet, row: dstRow, col: range.c0 + offset };
        writeCellSnapshot(wb, addr, snap.cells[offset] ?? null);
      }
    }
  });
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
    return { ...s, format: { ...s.format, formats } };
  });
  // Clear tail rows that were dropped.
  wb.withBatchedRecalc(() => {
    for (let r = range.r0 + snaps.length; r <= range.r1; r += 1) {
      for (let c = range.c0; c <= range.c1; c += 1) {
        wb.setBlank({ sheet: range.sheet, row: r, col: c });
      }
    }
  });
  wb.recalcAuto();
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
    ok = sortRange(state, store, workbook, range, options, history);
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
