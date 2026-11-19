import { addrKey, colLetter, MAX_COL } from '../engine/address.js';
import { createProfileTextSemantics, localeForWorkbook } from '../engine/profile-text-semantics.js';
import type { Addr, CellValue, Range } from '../engine/types.js';
import { fromEngineValue, writeCell } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { type ConditionalCellOverlay, evaluateConditional } from '../render/conditional.js';
import { isSingleCell, rangeArea, rangesIntersect } from '../store/selection-geometry.js';
import {
  type CellFormat,
  type FilterSortKey,
  mutators,
  type SpreadsheetStore,
  type State,
} from '../store/store.js';
import type { ConditionalIconRef } from '../store/types/conditional-rule.js';
import { type AuthorizedIntentFactory, authorizedWrite } from './authorized-write.js';
import { recordCommentChange } from './comment.js';
import { DEFAULT_COMMENT_AUTHOR } from './comment-author.js';
import { allCustomLists } from './custom-lists.js';
import { blockDataTableEdit } from './data-table-guard.js';
import { filterRangeForSort, inferAutoFilterRange, setFilterSort } from './filter.js';
import { cellColorKey, normalizeColorKey } from './filter-colors.js';
import { filterRangeAt } from './filter-ranges.js';
import { shiftFormulaRefs } from './formula-refs.js';
import type { History } from './history.js';
import type { InteractionOrigin } from './interaction-policy.js';
import { isCellWritable, warnProtected } from './protection.js';
import { recordFormatChange } from './slice-history.js';

/** Spreadsheet parity: refuse to sort when the range intersects any merge —
 *  rearranging rows would tear the merged rectangle apart. */
const rangeIntersectsMerges = (state: State, range: Range): boolean =>
  [...state.merges.byAnchor.values()].some((m) => rangesIntersect(m, range));

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
  sourceCol = addr.col,
): void => {
  if (!cell) {
    wb.setBlank(addr);
    return;
  }
  const formula = cell.formula
    ? shiftFormulaRefs(cell.formula, addr.row - sourceRow, addr.col - sourceCol)
    : null;
  writeCell(wb, addr, cell.value, formula);
};

export type SortDirection = 'asc' | 'desc';

export type SortOrientation = 'topToBottom' | 'leftToRight';

export interface SortKey {
  /** Column to sort by, in absolute sheet coords. Must lie within `range`.
   *  For `orientation: 'leftToRight'` this is the key row instead. */
  byCol: number;
  direction: SortDirection;
  /** Sort by values by default, or by cell/font color or conditional formatting icon. */
  sortOn?: 'value' | 'cellColor' | 'fontColor' | 'cellIcon' | 'customList';
  /** Target color for color sorts. Matching rows sort first for asc, last for desc. */
  color?: string;
  /** Target icon for icon sorts, matched as the color is. */
  icon?: ConditionalIconRef;
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
  /** When true, the first row is treated as a header and not moved. Ignored
   *  for `leftToRight`, which never keeps a header column. */
  hasHeader?: boolean;
  /** `leftToRight` sorts columns by the key row(s) named in `byCol`.
   *  Defaults to `topToBottom`. */
  orientation?: SortOrientation;
  /** Case-sensitive text order: lower case sorts before upper case among
   *  otherwise equal text (reversed for descending). Defaults to false. */
  matchCase?: boolean;
}

export interface RemoveDuplicatesOptions {
  /** Absolute sheet columns to compare. Defaults to every column in `range`. */
  columns?: readonly number[];
  /** When true, the first row is preserved and not compared as data. */
  hasHeader?: boolean;
}

/** What Remove Duplicates did: how many data rows it removed and how many unique rows remain, or
 *  why it refused to run. */
export type RemoveDuplicatesResult =
  | { ok: true; removed: number; remaining: number }
  | { ok: false; reason: 'too-large' | 'merged-cells' | 'protected' | 'engine-failed' };

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
                ...(comment.visible ? { visible: true as const } : {}),
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
    const visible = comment.visible === true;
    if (
      previous.comment === comment.text &&
      previous.commentAuthor === comment.author &&
      (previous.commentVisible === true) === visible
    ) {
      continue;
    }
    const { commentVisible: _visible, ...rest } = previous;
    formats.set(key, {
      ...rest,
      comment: comment.text,
      commentAuthor: comment.author ?? undefined,
      ...(visible ? { commentVisible: true } : {}),
    });
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

const absoluteColumnRange = (column: number, firstRow: number, lastRow: number): string => {
  const name = colLetter(column);
  return `$${name}$${firstRow + 1}:$${name}$${lastRow + 1}`;
};

const absoluteRowRange = (row: number, firstCol: number, lastCol: number): string =>
  `$${colLetter(firstCol)}$${row + 1}:$${colLetter(lastCol)}$${row + 1}`;

const numberArray = (values: readonly number[], horizontal: boolean): string =>
  `{${values.join(horizontal ? ',' : ';')}}`;

const compareCaseVariants = (left: string, right: string): number => {
  const n = Math.min(left.length, right.length);
  for (let i = 0; i < n; i += 1) {
    const l = left.charAt(i);
    const r = right.charAt(i);
    if (l === r) continue;
    const leftUpper = l !== l.toLowerCase();
    const rightUpper = r !== r.toLowerCase();
    if (leftUpper !== rightUpper) return leftUpper ? 1 : -1;
  }
  return 0;
};

const nativeEvaluationAnchor = (
  state: State,
  range: Range,
  rows: number,
  cols = 1,
): Addr | null => {
  // evaluateFormulaArray is read-only, but the engine still rejects a spill
  // that would land on a filled cell. Anchor at row 0 in the first run of
  // `cols` columns with nothing in the top `rows` rows; the formula uses only
  // absolute references, so the anchor does not change its meaning.
  const occupied = new Set<number>();
  const prefix = `${range.sheet}:`;
  for (const key of state.data.cells.keys()) {
    if (!key.startsWith(prefix)) continue;
    const [, row, col] = key.split(':');
    if (Number(row) < rows) occupied.add(Number(col));
  }
  let run = 0;
  for (let col = 0; col <= MAX_COL; col += 1) {
    run = occupied.has(col) ? 0 : run + 1;
    if (run === cols) return { sheet: range.sheet, row: 0, col: col - cols + 1 };
  }
  return null;
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
  start: number,
  sortKeys: readonly SortKey[],
  horizontal: boolean,
  matchCase: boolean,
  lists: readonly (readonly string[])[],
): number[] | null => {
  const last = horizontal ? range.c1 : range.r1;
  const count = last - start + 1;
  if (count <= 0) return [];
  if (count > MAX_NATIVE_SORT_ROWS) return null;

  const builtInLists = lists;
  const byKey = sortKeys.map((key) => {
    const direction = key.direction === 'desc' ? -1 : 1;
    const sortOn = key.sortOn ?? 'value';
    if (sortOn === 'value') {
      const keyRange = horizontal
        ? absoluteRowRange(key.byCol, start, last)
        : absoluteColumnRange(key.byCol, start, last);
      if (!matchCase) return `${keyRange},${direction}`;
      // SORTBY compares text without case; a second key of per-text ranks
      // orders the case variants that tie (lower case first).
      const texts: string[] = [];
      for (let line = start; line <= last; line += 1) {
        const v = sortKeyForCell(state, range.sheet, line, key, builtInLists, horizontal);
        texts.push(v.kind === 'text' ? (v.s ?? '') : '');
      }
      const distinct = [...new Set(texts)].sort(
        (a, b) =>
          (a.toLowerCase() < b.toLowerCase() ? -1 : a.toLowerCase() > b.toLowerCase() ? 1 : 0) ||
          compareCaseVariants(a, b),
      );
      const rank = new Map(distinct.map((text, index) => [text, index]));
      const caseRanks = texts.map((text) => rank.get(text) ?? 0);
      return `${keyRange},${direction},${numberArray(caseRanks, horizontal)},${direction}`;
    }
    const ranks: number[] = [];
    for (let line = start; line <= last; line += 1) {
      const sortValue = sortKeyForCell(state, range.sheet, line, key, builtInLists, horizontal);
      if (sortOn === 'cellColor' || sortOn === 'fontColor' || sortOn === 'cellIcon') {
        ranks.push(sortValue.match === true ? 0 : 1);
      } else {
        // Unknown custom-list entries remain after all listed entries in both
        // directions, matching the existing JS path's explicit sentinel.
        ranks.push(sortValue.customIndex ?? sortKeys.length + count);
      }
    }
    return `${numberArray(ranks, horizontal)},${direction}`;
  });
  const sequence = horizontal ? `SEQUENCE(1,${count})` : `SEQUENCE(${count})`;
  const formula = `=SORTBY(${sequence},${byKey.join(',')})`;
  try {
    const anchor = horizontal
      ? nativeEvaluationAnchor(state, range, 1, count)
      : nativeEvaluationAnchor(state, range, count);
    if (!anchor) return null;
    const result = wb.evaluateFormulaArray(anchor, formula);
    const rows = horizontal ? 1 : count;
    const cols = horizontal ? count : 1;
    if (!result || result.length !== rows || result.some((row) => row.length !== cols)) {
      return null;
    }
    const permutation: number[] = [];
    const seen = new Set<number>();
    const flat = horizontal ? (result[0] ?? []) : result.map((row) => row?.[0]);
    if (flat.length !== count) return null;
    for (const raw of flat) {
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
    const anchor = nativeEvaluationAnchor(state, range, count, keyColumns.length);
    if (!anchor) return null;
    const result = wb.evaluateFormulaArray(anchor, `=UNIQUE(${keyFormula})`);
    if (!result || result.length < 1 || result.length > count) return null;
    if (result.some((row) => row.length !== keyColumns.length)) return null;
    const uniqueValues = result.map((row) => {
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
      ...(key.icon ? { icon: { ...key.icon } } : {}),
      ...(key.customList ? { customList: [...key.customList] } : {}),
    });
  }
  return out;
};

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

interface CellSortKey {
  kind: SortValueKind;
  n?: number;
  s?: string;
  match?: boolean;
  customIndex?: number | null;
}

/** The conditional formatting overlays of `state`, computed once per state for an icon sort. */
const overlayCache = new WeakMap<State, ReadonlyMap<string, ConditionalCellOverlay>>();
const conditionalOverlays = (state: State): ReadonlyMap<string, ConditionalCellOverlay> => {
  let overlays = overlayCache.get(state);
  if (!overlays) {
    overlays = evaluateConditional(state);
    overlayCache.set(state, overlays);
  }
  return overlays;
};

const sortKeyForCell = (
  state: State,
  sheet: number,
  line: number,
  key: SortKey,
  builtInLists: readonly (readonly string[])[],
  horizontal = false,
): CellSortKey => {
  const sortOn = key.sortOn ?? 'value';
  const row = horizontal ? key.byCol : line;
  const col = horizontal ? line : key.byCol;
  if (sortOn === 'cellIcon') {
    const overlay = conditionalOverlays(state).get(addrKey({ sheet, row, col }));
    return {
      kind: 'color',
      match: !!key.icon && overlay?.iconKind === key.icon.set && overlay.iconSlot === key.icon.slot,
    };
  }
  if (sortOn === 'cellColor' || sortOn === 'fontColor') {
    const fmt = state.format.formats.get(addrKey({ sheet, row, col }));
    return {
      kind: 'color',
      match: !!key.color && cellColorKey(fmt, sortOn) === normalizeColorKey(key.color),
    };
  }
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
  a: CellSortKey,
  b: CellSortKey,
  direction: SortDirection,
  compareText: (left: string, right: string) => number,
  matchCase = false,
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
      return (
        dir *
        (compareText(a.s ?? '', b.s ?? '') ||
          (matchCase ? compareCaseVariants(a.s ?? '', b.s ?? '') : 0))
      );
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
export function inferSortHasHeader(
  state: State,
  range: Range,
  orientation: SortOrientation = 'topToBottom',
): boolean {
  // Excel never guesses a header column for a left-to-right sort.
  if (orientation === 'leftToRight') return false;
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

interface PlacedFormat {
  addr: Addr;
  format: CellFormat;
}

/** Replace the formats of `affected` with `placed`, keep the engine's comments
 *  in step, and record both as undoable changes. `state` carries the formats
 *  as they were before the rewrite. */
const rewriteRangeFormats = (
  state: State,
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  affected: readonly Addr[],
  placed: readonly PlacedFormat[],
): void => {
  const commentTargets = new Map<string, Addr>();
  const rememberComment = (addr: Addr, format: CellFormat | undefined): void => {
    if (typeof format?.comment !== 'string' || format.comment.length === 0) return;
    commentTargets.set(addrKey(addr), addr);
  };
  for (const addr of affected) rememberComment(addr, state.format.formats.get(addrKey(addr)));
  for (const { addr, format } of placed) rememberComment(addr, format);
  const applyFormats = (): void => {
    store.setState((s) => {
      const formats = new Map(s.format.formats);
      for (const addr of affected) formats.delete(addrKey(addr));
      for (const { addr, format } of placed) formats.set(addrKey(addr), format);
      return { ...s, format: { ...s.format, formats } };
    });
    for (const addr of commentTargets.values()) {
      const format = store.getState().format.formats.get(addrKey(addr));
      wb.setCommentEntry(
        addr.sheet,
        addr.row,
        addr.col,
        format?.commentAuthor ?? DEFAULT_COMMENT_AUTHOR,
        format?.comment ?? '',
        format?.commentVisible === true,
      );
    }
  };
  recordFormatChange(history, store, () => {
    recordCommentChange(history, store, wb, [...commentTargets.values()], applyFormats);
  });
};

/** True when `row` is hidden and lies inside an AutoFilter whose columns
 *  overlap `range`, i.e. the filter (not a manual hide) is what hides it. */
const filterHidesRow = (state: State, range: Range, row: number): boolean =>
  state.layout.hiddenRows.has(row) &&
  state.ui.filterRanges.some(
    (f) =>
      f.sheet === range.sheet && row > f.r0 && row <= f.r1 && f.c0 <= range.c1 && f.c1 >= range.c0,
  );

/** The cells a sort of `range` rewrites: everything below a kept header row
 *  (a left-to-right sort keeps none), or null when nothing is left to sort. */
export const sortedDataRange = (range: Range, opts: SortOptions): Range | null => {
  const horizontal = opts.orientation === 'leftToRight';
  if (horizontal || !opts.hasHeader) return { ...range };
  return range.r0 + 1 > range.r1 ? null : { ...range, r0: range.r0 + 1 };
};

/** The intents of a sort that rewrites `dataRange`. */
export const sortIntents =
  (dataRange: Range, commandId: string): AuthorizedIntentFactory =>
  (origin) => [
    { operation: 'sort', origin, commandId, effects: [{ kind: 'range', range: { ...dataRange } }] },
  ];

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
  const horizontal = opts.orientation === 'leftToRight';
  const matchCase = opts.matchCase === true;
  // A "line" is one sortable unit: a row, or a column when sorting left to
  // right. `pos` walks across the line.
  const dataRange = sortedDataRange(range, opts);
  if (!dataRange) return false;
  const start = horizontal ? dataRange.c0 : dataRange.r0;
  const last = horizontal ? range.c1 : range.r1;
  const sortKeys = normalizedSortKeys(opts);
  if (sortKeys.length === 0) return false;
  const keyLo = horizontal ? range.r0 : range.c0;
  const keyHi = horizontal ? range.r1 : range.c1;
  if (sortKeys.some((key) => key.byCol < keyLo || key.byCol > keyHi)) return false;
  if (!canRewriteExactRange(dataRange)) return false;
  if (rangeIntersectsMerges(state, range)) return false;
  if (!ensureWritableRange(state, dataRange)) return false;
  if (blockDataTableEdit(wb, dataRange)) return false;

  // Inside an AutoFilter, rows the filter hides keep their place and only the
  // visible rows trade positions among themselves.
  const lines: number[] = [];
  for (let l = start; l <= last; l += 1) {
    if (!horizontal && filterHidesRow(state, range, l)) continue;
    lines.push(l);
  }
  const skipsLines = lines.length !== last - start + 1;

  // A real formulon bundle owns the comparison semantics for formula-driven
  // sorts. Do this before hydrating comments or touching any cell so an
  // unsupported/malformed native result is an atomic no-op. Stub workbooks
  // intentionally use the explicit-locale JS comparator below.
  const nativeBundle =
    !skipsLines && wb.isStub === false && wb.capabilities.arrayFormulaEvaluation === true;
  const nativeOrder = nativeBundle
    ? nativeSortPermutation(
        state,
        wb,
        range,
        start,
        sortKeys,
        horizontal,
        matchCase,
        allCustomLists(store, localeForWorkbook(wb)),
      )
    : null;
  if (nativeBundle && nativeOrder === null) return false;

  state = hydrateSortComments(store, wb, dataRange, horizontal ? range.r0 : start);
  const builtInLists = allCustomLists(store, localeForWorkbook(wb));
  const addrAt = (line: number, pos: number): Addr =>
    horizontal
      ? { sheet: range.sheet, row: pos, col: line }
      : { sheet: range.sheet, row: line, col: pos };
  const posLo = horizontal ? range.r0 : range.c0;
  const posHi = horizontal ? range.r1 : range.c1;

  // Snapshot the lines we'll move, including formula text so we can write it back.
  interface LineSnap {
    cells: Array<{
      value: CellValue;
      formula: string | null;
      pos: number;
      format?: CellFormat;
    }>;
    sortKeys: CellSortKey[];
  }
  const snaps: { line: number; snap: LineSnap }[] = [];
  for (const l of lines) {
    const cells: LineSnap['cells'] = [];
    for (let p = posLo; p <= posHi; p += 1) {
      const key = addrKey(addrAt(l, p));
      const cell = state.data.cells.get(key);
      cells.push({
        value: cell?.value ?? { kind: 'blank' },
        formula: cell?.formula ?? null,
        pos: p,
        format: cloneFormat(state.format.formats.get(key)),
      });
    }
    snaps.push({
      line: l,
      snap: {
        cells,
        sortKeys: sortKeys.map((key) =>
          sortKeyForCell(state, range.sheet, l, key, builtInLists, horizontal),
        ),
      },
    });
  }

  if (nativeOrder) {
    const bySourceIndex = new Map(snaps.map((entry, index) => [entry.line, index]));
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
          matchCase,
        );
        if (result !== 0) return result;
      }
      return a.line - b.line;
    });
  }

  // Write back into wb in sorted order.
  wb.withBatchedRecalc(() => {
    for (let i = 0; i < snaps.length; i += 1) {
      const dstLine = lines[i] ?? start + i;
      const entry = snaps[i];
      if (!entry) continue;
      for (const cell of entry.snap.cells) {
        const addr = addrAt(dstLine, cell.pos);
        const from = addrAt(entry.line, cell.pos);
        writeCellSnapshot(wb, addr, cell, from.row, from.col);
      }
    }
  });
  const affected: Addr[] = [];
  for (const l of lines) {
    for (let p = posLo; p <= posHi; p += 1) affected.push(addrAt(l, p));
  }
  const placed: PlacedFormat[] = [];
  for (let i = 0; i < snaps.length; i += 1) {
    const snap = snaps[i]?.snap;
    if (!snap) continue;
    for (const cell of snap.cells) {
      if (!cell.format) continue;
      placed.push({ addr: addrAt(lines[i] ?? start + i, cell.pos), format: cell.format });
    }
  }
  rewriteRangeFormats(state, store, wb, history, affected, placed);
  wb.recalcAuto();
  const filterRange = horizontal
    ? null
    : filterRangeForSort(store.getState(), range, opts.hasHeader === true);
  if (filterRange) {
    const keys = sortKeys.flatMap((key): FilterSortKey[] => {
      const sortOn = key.sortOn ?? 'value';
      if (sortOn === 'value') return [{ byCol: key.byCol, direction: key.direction }];
      if ((sortOn === 'cellColor' || sortOn === 'fontColor') && key.color) {
        return [
          {
            byCol: key.byCol,
            direction: key.direction,
            sortOn,
            color: normalizeColorKey(key.color),
          },
        ];
      }
      if (sortOn === 'cellIcon' && key.icon) {
        return [{ byCol: key.byCol, direction: key.direction, sortOn, icon: { ...key.icon } }];
      }
      if (sortOn === 'customList' && key.customList?.length) {
        return [
          { byCol: key.byCol, direction: key.direction, sortOn, customList: [...key.customList] },
        ];
      }
      return [];
    });
    setFilterSort(store, history, filterRange, keys, wb);
  }
  return true;
}

/** Remove duplicate rows from a range (keeping the first occurrence). */
export function removeDuplicates(
  state: State,
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  range: Range,
  options: RemoveDuplicatesOptions = {},
  history: History | null = null,
): RemoveDuplicatesResult {
  if (rangeArea(range) > MAX_SORT_CELLS) return { ok: false, reason: 'too-large' };
  if (rangeIntersectsMerges(state, range)) return { ok: false, reason: 'merged-cells' };
  if (!ensureWritableRange(state, range)) return { ok: false, reason: 'protected' };
  const columns = (options.columns?.length ? options.columns : undefined)?.filter(
    (col, index, arr) =>
      col >= range.c0 && col <= range.c1 && Number.isInteger(col) && arr.indexOf(col) === index,
  );
  const keyColumns = columns?.length
    ? columns
    : Array.from({ length: range.c1 - range.c0 + 1 }, (_, i) => range.c0 + i);
  const firstDataRow = options.hasHeader ? range.r0 + 1 : range.r0;
  if (firstDataRow > range.r1) return { ok: true, removed: 0, remaining: 0 };
  const nativeBundle = wb.isStub === false && wb.capabilities.arrayFormulaEvaluation === true;
  if (nativeBundle && range.r1 - firstDataRow + 1 > MAX_NATIVE_UNIQUE_ROWS) {
    return { ok: false, reason: 'too-large' };
  }
  const nativeKeep = nativeBundle
    ? nativeUniqueRows(state, wb, range, firstDataRow, keyColumns)
    : null;
  if (nativeBundle && nativeKeep === null) return { ok: false, reason: 'engine-failed' };
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
      const key = JSON.stringify(sig);
      if (seen.has(key)) continue;
      seen.add(key);
      keep.push(r);
    }
  }
  const remaining = keep.length - (options.hasHeader ? 1 : 0);
  if (keep.length === range.r1 - range.r0 + 1) return { ok: true, removed: 0, remaining };

  state = hydrateSortComments(store, wb, range, range.r0);

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
  const affected: Addr[] = [];
  for (let r = range.r0; r <= range.r1; r += 1) {
    for (let c = range.c0; c <= range.c1; c += 1)
      affected.push({ sheet: range.sheet, row: r, col: c });
  }
  const placed: PlacedFormat[] = [];
  for (let i = 0; i < snaps.length; i += 1) {
    const cells = snaps[i]?.cells ?? [];
    for (let offset = 0; offset < cells.length; offset += 1) {
      const format = cells[offset]?.format;
      if (!format) continue;
      placed.push({
        addr: { sheet: range.sheet, row: range.r0 + i, col: range.c0 + offset },
        format,
      });
    }
  }
  rewriteRangeFormats(state, store, wb, history, affected, placed);
  // Clear tail rows that were dropped.
  wb.withBatchedRecalc(() => {
    for (let r = range.r0 + snaps.length; r <= range.r1; r += 1) {
      for (let c = range.c0; c <= range.c1; c += 1) {
        wb.setBlank({ sheet: range.sheet, row: r, col: c });
      }
    }
  });
  wb.recalcAuto();
  return { ok: true, removed: range.r1 - range.r0 + 1 - snaps.length, remaining };
}

export interface SortRangeWithHistoryDeps {
  store: SpreadsheetStore;
  workbook: WorkbookHandle;
  history: History;
  range: Range;
  options: SortOptions;
  origin: InteractionOrigin;
  commandId: string;
}

/** Run [[sortRange]] as one authorized history step and refresh the affected
 *  sheet's cells when anything actually moved. Returns whether the sort
 *  changed any data, so callers can short-circuit follow-up side effects. */
export const sortRangeWithHistory = (deps: SortRangeWithHistoryDeps): boolean => {
  const { store, workbook, history, range, options } = deps;
  const dataRange = sortedDataRange(range, options);
  if (!dataRange) return false;
  const state = store.getState();
  const result = authorizedWrite({
    store,
    history,
    intents: sortIntents(dataRange, deps.commandId),
    origin: deps.origin,
    write: () => sortRange(state, store, workbook, range, options, history),
  });
  const ok = result.ok && result.value;
  if (ok) mutators.replaceCells(store, workbook.cells(state.data.sheetIndex));
  return ok;
};

export interface SortActiveColumnAutoOptions {
  store: SpreadsheetStore;
  workbook: WorkbookHandle;
  history: History;
  direction: SortDirection;
  origin: InteractionOrigin;
  commandId: string;
}

/** "Sort A→Z / Z→A" toolbar action. Auto-detects the contiguous range around
 *  the selection, picks header behavior automatically, sorts by the active
 *  cell's column, and refreshes the sheet cells. Returns whether anything
 *  actually changed so the host can skip side effects on a no-op. */
export const sortActiveColumnAuto = (deps: SortActiveColumnAutoOptions): boolean => {
  const { store, workbook, history, direction, origin, commandId } = deps;
  const state = store.getState();
  const sel = state.selection.range;
  const single = isSingleCell(sel);
  // A cell inside an AutoFilter sorts that whole filter, blank rows included,
  // with the filter's first row as the header.
  const filter = single ? filterRangeAt(state, state.selection.active) : null;
  const range = filter ?? inferAutoFilterRange(state);
  return sortRangeWithHistory({
    store,
    workbook,
    history,
    range,
    options: {
      byCol: state.selection.active.col,
      direction,
      hasHeader: filter ? true : inferSortHasHeader(state, range),
    },
    origin,
    commandId,
  });
};
