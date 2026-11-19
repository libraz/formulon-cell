import { addrKey, colLetter, MAX_COL, parseAddrKey } from '../engine/address.js';
import { createProfileTextSemantics } from '../engine/profile-text-semantics.js';
import type { Addr, CellValue, Range } from '../engine/types.js';
import { formatBoolean, formatCell, fromEngineValue } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { State } from '../store/store.js';
import { canEvaluateFilterRange, dataRowCount, rowInRange } from './filter-ranges.js';
import { formatNumber } from './format.js';

// UNIQUE in the bundled engine grows much faster than SORTBY. The matching
// budget is kept explicit here so a large checklist never blocks the UI while
// asking the engine for an unbounded dynamic array.
const MAX_NATIVE_DISTINCT_ROWS = 20_000;

type FilterItemKind = 'number' | 'text' | 'blank';

export interface FilterValueItem {
  /** Matching key — identical to `filterValueKey(cell.value)`, so it can be fed
   *  straight back into the hidden-value set that {@link applyValueFilter}
   *  matches against. */
  key: string;
  /** Display label honouring the column's number format (dates render as dates,
   *  currency keeps its symbol) so the checklist mirrors the grid rather than
   *  showing a raw serial. */
  label: string;
}

export type DistinctFilterFailureReason = 'native-limit' | 'native-result';

export interface DistinctFilterItemsResult {
  items: FilterValueItem[];
  /** Native dynamic-array failure, while items remains a safe explicit
   * locale approximation for checklist rendering. */
  failure: DistinctFilterFailureReason | null;
}

const classifyFilterValue = (value: CellValue | undefined): FilterItemKind => {
  if (!value || value.kind === 'blank' || value.kind === 'error') return 'blank';
  if (value.kind === 'number') return 'number';
  return 'text';
};

export const filterItemLabel = (
  state: State,
  addr: Addr,
  value: CellValue,
  locale: string,
): string => {
  if (value.kind === 'number') {
    const fmt = state.format.formats.get(addrKey(addr));
    if (fmt?.numFmt)
      return formatNumber(value.value, fmt.numFmt, locale, state.ui.date1904 === true);
  }
  return formatCell(value, locale);
};

const FILTER_GROUP_ORDER: Record<FilterItemKind, number> = { number: 0, text: 1, blank: 2 };

const nativeFilterAnchor = (state: State, range: Range): Addr => {
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

/** Native UNIQUE/SORT can expose a source blank as numeric zero through the
 * engine value ABI. Keep those two source forms in one representative queue;
 * the queue remains source ordered, so a real zero and a blank stay distinct
 * when the native result contains both. */
const nativeFilterValueKey = (value: CellValue): string => {
  if (value.kind === 'blank' || (value.kind === 'number' && value.value === 0)) {
    return 'zero-or-blank';
  }
  if (value.kind === 'number') return `number:${String(value.value)}`;
  if (value.kind === 'text') return `text:${value.value}`;
  if (value.kind === 'bool') return `bool:${value.value ? '1' : '0'}`;
  return `error:${value.code}:${value.text}`;
};

const nativeDistinctFilterItems = (
  state: State,
  wb: WorkbookHandle,
  range: Range,
  byCol: number,
  locale: string,
): FilterValueItem[] | null => {
  const count = Math.max(0, range.r1 - range.r0);
  if (count === 0) return [];
  const column = colLetter(byCol);
  const first = range.r0 + 2;
  const last = range.r1 + 1;
  const formula = `=SORT(UNIQUE($${column}$${first}:$${column}$${last}))`;
  try {
    const result = wb.evaluateFormulaArray(nativeFilterAnchor(state, range), formula);
    if (!result || result.length < 1 || result.length > count) return null;
    if (result.some((row) => row.length !== 1)) return null;
    const values = result.map((row) => (row?.[0] ? fromEngineValue(row[0]) : null));
    if (values.some((value): value is null => value === null)) return null;
    const candidates = new Map<string, number[]>();
    for (let row = range.r0 + 1; row <= range.r1; row += 1) {
      const source =
        state.data.cells.get(addrKey({ sheet: range.sheet, row, col: byCol }))?.value ??
        ({ kind: 'blank' } as const);
      const key = nativeFilterValueKey(source);
      const rows = candidates.get(key);
      if (rows) rows.push(row);
      else candidates.set(key, [row]);
    }
    const sourceAt = (row: number): CellValue =>
      state.data.cells.get(addrKey({ sheet: range.sheet, row, col: byCol }))?.value ?? {
        kind: 'blank',
      };
    const used = new Set<number>();
    const items: FilterValueItem[] = [];
    for (const value of values) {
      if (!value) return null;
      const rows = (candidates.get(nativeFilterValueKey(value)) ?? []).filter(
        (row) => !used.has(row),
      );
      // Zero and blank share a key; a blank the engine returns as such takes a blank row.
      const representative =
        rows.find((row) => (sourceAt(row).kind === 'blank') === (value.kind === 'blank')) ??
        rows[0];
      if (representative === undefined) return null;
      used.add(representative);
      const addr = { sheet: range.sheet, row: representative, col: byCol };
      const sourceValue = sourceAt(representative);
      const key = sourceValue.kind === 'blank' ? '' : filterValueKey(value);
      items.push({
        key,
        label: key === '' ? '' : filterItemLabel(state, addr, value, locale),
      });
    }
    return items;
  } catch {
    return null;
  }
};

/** Distinct values of `byCol` in `range` (excluding the header row) as
 *  {@link FilterValueItem}s. De-duplicated by filter key and ordered the way the
 *  AutoFilter checklist orders them — numbers ascending numerically, then text
 *  ascending, with blanks last — instead of a raw lexical sort of serials. */
function fallbackDistinctFilterItems(
  state: State,
  range: Range,
  byCol: number,
  locale = 'en-US',
): FilterValueItem[] {
  const semantics = createProfileTextSemantics(locale);
  const items = new Map<
    string,
    { key: string; label: string; kind: FilterItemKind; num: number }
  >();
  const visitValue = (row: number, value: CellValue | undefined): void => {
    const addr = { sheet: range.sheet, row, col: byCol };
    const key = filterValueKey(value);
    // Keep number/text/blank types separate. Within the text class the
    // locale key folds case and width for linear-time checklist membership.
    const matchKey =
      value?.kind === 'text' ? semantics.key(key) : `${value?.kind ?? 'blank'}:${key}`;
    if (items.has(matchKey)) return;
    const kind = key === '' ? 'blank' : classifyFilterValue(value);
    const num = value?.kind === 'number' ? value.value : 0;
    const label = key === '' || !value ? '' : filterItemLabel(state, addr, value, locale);
    items.set(matchKey, { key, label, kind, num });
  };

  if (canEvaluateFilterRange(range)) {
    for (let r = range.r0 + 1; r <= range.r1; r += 1) {
      const addr = { sheet: range.sheet, row: r, col: byCol };
      visitValue(r, state.data.cells.get(addrKey(addr))?.value);
    }
  } else {
    for (const [key, cell] of state.data.cells) {
      const addr = parseAddrKey(key);
      if (!addr || addr.sheet !== range.sheet || addr.col !== byCol || !rowInRange(addr.row, range))
        continue;
      visitValue(addr.row, cell.value);
    }
  }
  return Array.from(items.values())
    .sort((a, b) => {
      const groupDelta = FILTER_GROUP_ORDER[a.kind] - FILTER_GROUP_ORDER[b.kind];
      if (groupDelta !== 0) return groupDelta;
      if (a.kind === 'number') {
        return a.num - b.num || semantics.compare(a.label, b.label);
      }
      if (a.kind === 'text') return semantics.compare(a.label, b.label);
      return 0;
    })
    .map(({ key, label }) => ({ key, label }));
}

export function distinctFilterItemsResult(
  state: State,
  range: Range,
  byCol: number,
  locale = 'en-US',
  wb?: WorkbookHandle,
): DistinctFilterItemsResult {
  if (wb?.isStub === false && wb.capabilities.arrayFormulaEvaluation === true) {
    // A real bundle owns UNIQUE/SORT equality and collation for bounded
    // ranges. Above the native budget, or when the result is malformed, use
    // the explicit-locale store path so a checklist failure cannot turn into
    // an empty list that hides every row when the user applies it.
    if (dataRowCount(range) > MAX_NATIVE_DISTINCT_ROWS) {
      return {
        items: fallbackDistinctFilterItems(state, range, byCol, locale),
        failure: 'native-limit',
      };
    }
    const native = nativeDistinctFilterItems(state, wb, range, byCol, locale);
    if (native !== null) return { items: native, failure: null };
    return {
      items: fallbackDistinctFilterItems(state, range, byCol, locale),
      failure: 'native-result',
    };
  }
  return { items: fallbackDistinctFilterItems(state, range, byCol, locale), failure: null };
}

export function distinctFilterItems(
  state: State,
  range: Range,
  byCol: number,
  locale = 'en-US',
  wb?: WorkbookHandle,
): FilterValueItem[] {
  return distinctFilterItemsResult(state, range, byCol, locale, wb).items;
}

/** Distinct (string-coerced) filter keys found in `byCol` of `range`, excluding
 *  the header row, ordered like the AutoFilter checklist. */
export function distinctValues(
  state: State,
  range: Range,
  byCol: number,
  locale = 'en-US',
  wb?: WorkbookHandle,
): string[] {
  return distinctFilterItems(state, range, byCol, locale, wb).map((item) => item.key);
}

export const filterValueKey = (v: unknown): string => {
  if (!v || typeof v !== 'object') return '';
  const cv = v as { kind: string; value?: unknown };
  if (cv.kind === 'number') return String(cv.value);
  if (cv.kind === 'text') return String(cv.value ?? '');
  if (cv.kind === 'bool') return formatBoolean(Boolean(cv.value));
  return '';
};
