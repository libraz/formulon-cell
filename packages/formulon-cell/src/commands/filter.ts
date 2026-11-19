import { addrKey, MAX_COL, MAX_ROW } from '../engine/address.js';
import { syncAutoFilterToEngine } from '../engine/auto-filter-sync.js';
import {
  createProfileTextSemantics,
  localeForWorkbook,
  profileTextKey,
} from '../engine/profile-text-semantics.js';
import type { Addr, CellValue, Range } from '../engine/types.js';
import { formatCell, writeCell } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { isDataTableFormula } from '../engine/workbook-handle-data-tables.js';
import { evaluateConditional } from '../render/conditional.js';
import { rangeContainsRange, sameRange } from '../store/selection-geometry.js';
import type {
  FilterColorSpec,
  FilterSortSpec,
  SpreadsheetStore,
  State,
  ValueFilterCriteria,
} from '../store/store.js';
import { parseLocalizedNumericInput } from './coerce-input.js';
import { blockDataTableEdit } from './data-table-guard.js';
import { colorMatches, normalizeColorKey } from './filter-colors.js';
import {
  type ConditionFilterOptions,
  canonicalNumericCriterion,
  cellNumber,
  criteriaConditionMatches,
  numericConditionOperator,
  periodMatcher,
  rankedBound,
} from './filter-conditions.js';
import { distinctFilterItemsResult, filterItemLabel, filterValueKey } from './filter-items.js';
import {
  activeFilterRange,
  canEvaluateFilterRange,
  rowInRange,
  withFilterRange,
  withoutFilterRange,
} from './filter-ranges.js';
import type { History } from './history.js';

export type FilterPredicate = (
  cell: { value: unknown; formula: string | null } | undefined,
) => boolean;

/** Mirror the store's filters on `range`'s sheet to the engine. With `range`,
 *  only that filter is rewritten; removed filters are reconciled either way. */
function syncFilterToEngine(
  wb: WorkbookHandle | undefined,
  store: SpreadsheetStore,
  range?: Range,
): void {
  if (!wb) return;
  const state = store.getState();
  syncAutoFilterToEngine(
    wb,
    state,
    range?.sheet ?? state.data.sheetIndex,
    store,
    localeForWorkbook(wb),
    range,
  );
}

const clearHiddenRowsInRange = (hiddenRows: Set<number>, range: Range): Set<number> => {
  const next = new Set(hiddenRows);
  for (const row of hiddenRows) {
    if (rowInRange(row, range)) next.delete(row);
  }
  return next;
};

function stampFilterRangeWithoutCriteria(store: SpreadsheetStore, range: Range): void {
  store.setState((s) => ({
    ...s,
    layout: { ...s.layout, hiddenRows: clearHiddenRowsInRange(s.layout.hiddenRows, range) },
    ui: {
      ...s.ui,
      filterRanges: withFilterRange(s, range),
      filterCriteria: filterCriteriaAfterClear(s.ui.filterCriteria, range),
    },
  }));
}

/** Hide rows in `range` whose `byCol` value fails `predicate`. The first
 *  row of the range is treated as a header and stays visible.
 *  Excel-style replace semantics: rows previously hidden inside `range` are
 *  revealed first, so re-filtering the same range never accumulates stale
 *  hides. Hidden rows live in `layout.hiddenRows` (UI-only — engine state is
 *  untouched). The range is also added to `ui.filterRanges` so the column
 *  headers paint the autofilter chevron. */
export function applyFilter(
  state: State,
  store: SpreadsheetStore,
  range: Range,
  byCol: number,
  predicate: FilterPredicate,
): number {
  return applyFilterColumns(state, store, range, [{ byCol, predicate }]);
}

/** Multi-column variant of {@link applyFilter}: a row survives only when it
 *  passes *every* column predicate, the way Excel ANDs filtered columns. Prior
 *  hides inside `range` are cleared first so the result is a full replace. */
export function applyFilterColumns(
  state: State,
  store: SpreadsheetStore,
  range: Range,
  columns: ReadonlyArray<{ byCol: number; predicate: FilterPredicate }>,
): number {
  if (!canEvaluateFilterRange(range)) {
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, hiddenRows: clearHiddenRowsInRange(s.layout.hiddenRows, range) },
      ui: { ...s.ui, filterRanges: withFilterRange(s, range) },
    }));
    return 0;
  }

  const newHidden = clearHiddenRowsInRange(state.layout.hiddenRows, range);
  let hiddenCount = 0;
  for (let r = range.r0 + 1; r <= range.r1; r += 1) {
    const survives = columns.every(({ byCol, predicate }) =>
      predicate(state.data.cells.get(addrKey({ sheet: range.sheet, row: r, col: byCol }))),
    );
    if (!survives) {
      newHidden.add(r);
      hiddenCount += 1;
    }
  }
  store.setState((s) => ({
    ...s,
    layout: { ...s.layout, hiddenRows: newHidden },
    ui: { ...s.ui, filterRanges: withFilterRange(s, range) },
  }));
  return hiddenCount;
}

export function applyValueFilter(
  state: State,
  store: SpreadsheetStore,
  range: Range,
  byCol: number,
  hiddenValues: readonly string[],
  wb?: WorkbookHandle,
): number {
  if (!canEvaluateFilterRange(range)) {
    stampFilterRangeWithoutCriteria(store, range);
    syncFilterToEngine(wb, store, range);
    return 0;
  }
  const locale = localeForWorkbook(wb);
  const semantics = createProfileTextSemantics(locale);
  const uniqueHiddenValues: string[] = [];
  const seenHiddenValues = new Set<string>();
  for (const value of hiddenValues) {
    const key = semantics.key(value);
    if (seenHiddenValues.has(key)) continue;
    seenHiddenValues.add(key);
    uniqueHiddenValues.push(value);
  }
  const criteria = upsertCriteria(state.ui.filterCriteria, {
    range,
    byCol,
    hiddenValues: uniqueHiddenValues.sort((a, b) => semantics.compare(a, b)),
  });
  // Recomputing with no remaining criterion would otherwise keep the rows
  // hidden by the just-cleared column. Start from a revealed target range and
  // let any other active criteria add their rows back as needed.
  const recomputeState: State = {
    ...state,
    layout: {
      ...state.layout,
      hiddenRows: clearHiddenRowsInRange(state.layout.hiddenRows, range),
    },
  };
  const { hidden, count } = recomputeHiddenFromCriteria(recomputeState, criteria, locale);
  store.setState((s) => ({
    ...s,
    layout: { ...s.layout, hiddenRows: hidden },
    ui: { ...s.ui, filterRanges: withFilterRange(s, range), filterCriteria: criteria },
  }));
  syncFilterToEngine(wb, store, range);
  return count;
}

export function applyConditionFilter(
  state: State,
  store: SpreadsheetStore,
  range: Range,
  byCol: number,
  condition: ConditionFilterOptions,
  wb?: WorkbookHandle,
): number {
  if (!canEvaluateFilterRange(range)) {
    stampFilterRangeWithoutCriteria(store, range);
    syncFilterToEngine(wb, store, range);
    return 0;
  }
  const locale = localeForWorkbook(wb);
  const value = numericConditionOperator(condition.op)
    ? canonicalNumericCriterion(condition.value, locale)
    : condition.value;
  const criteria = upsertCriteria(state.ui.filterCriteria, {
    range,
    byCol,
    hiddenValues: [],
    condition: {
      op: condition.op,
      value,
      ...(condition.op === 'between'
        ? { value2: canonicalNumericCriterion(condition.value2 ?? '', locale) }
        : {}),
      ...(condition.op === 'top10' || condition.op === 'bottom10'
        ? { unit: condition.unit ?? 'items' }
        : {}),
      ...(condition.op === 'datePeriod' && condition.period ? { period: condition.period } : {}),
      ...(condition.second
        ? {
            second: {
              op: condition.second.op,
              value: numericConditionOperator(condition.second.op)
                ? canonicalNumericCriterion(condition.second.value, locale)
                : condition.second.value,
            },
            join: condition.join ?? 'and',
          }
        : {}),
    },
  });
  const { hidden, count } = recomputeHiddenFromCriteria(state, criteria, locale);
  store.setState((s) => ({
    ...s,
    layout: { ...s.layout, hiddenRows: hidden },
    ui: { ...s.ui, filterRanges: withFilterRange(s, range), filterCriteria: criteria },
  }));
  syncFilterToEngine(wb, store, range);
  return count;
}

export function applyColorFilter(
  state: State,
  store: SpreadsheetStore,
  range: Range,
  byCol: number,
  color: FilterColorSpec,
  wb?: WorkbookHandle,
): number {
  if (!canEvaluateFilterRange(range)) {
    stampFilterRangeWithoutCriteria(store, range);
    syncFilterToEngine(wb, store, range);
    return 0;
  }
  const criteria = upsertCriteria(state.ui.filterCriteria, {
    range,
    byCol,
    hiddenValues: [],
    color:
      color.kind === 'cellIcon'
        ? { kind: 'cellIcon', icon: { set: color.icon.set, slot: color.icon.slot } }
        : { kind: color.kind, color: normalizeColorKey(color.color) },
  });
  const { hidden, count } = recomputeHiddenFromCriteria(state, criteria, localeForWorkbook(wb));
  store.setState((s) => ({
    ...s,
    layout: { ...s.layout, hiddenRows: hidden },
    ui: { ...s.ui, filterRanges: withFilterRange(s, range), filterCriteria: criteria },
  }));
  syncFilterToEngine(wb, store, range);
  return count;
}

/**
 * Recompute the hidden-row set from the full criteria list, ANDing across
 * every filtered column the way Excel's AutoFilter does: a row survives only
 * when it passes *all* active column criteria. Rows inside the criteria ranges
 * are re-evaluated from scratch (so replacing one column's filter never leaves
 * another column's hides stale), while hides outside the ranges are preserved.
 */
function recomputeHiddenFromCriteria(
  state: State,
  criteria: readonly ValueFilterCriteria[],
  locale = 'en-US',
): { hidden: Set<number>; count: number } {
  const baseline = new Set(state.layout.hiddenRows);
  for (const c of criteria) {
    for (const row of state.layout.hiddenRows) {
      if (rowInRange(row, c.range)) baseline.delete(row);
    }
  }
  const next = new Set(baseline);
  for (const c of criteria) {
    if (!canEvaluateFilterRange(c.range)) continue;
    const bound = c.condition ? rankedBound(state, c.range, c.byCol, c.condition) : null;
    const inPeriod = periodMatcher(c.condition, state.ui.date1904 === true);
    const icons = c.color?.kind === 'cellIcon' ? evaluateConditional(state) : null;
    const hiddenSet =
      c.condition || c.color
        ? null
        : new Set(c.hiddenValues.map((value) => filterMatchKey(value, locale)));
    for (let r = c.range.r0 + 1; r <= c.range.r1; r += 1) {
      if (next.has(r)) continue;
      const cell = state.data.cells.get(addrKey({ sheet: c.range.sheet, row: r, col: c.byCol }));
      const hide = c.color
        ? !colorMatches(state, c.range.sheet, r, c.byCol, c.color, icons)
        : c.condition
          ? !criteriaConditionMatches(cell, c.condition, locale, bound, inPeriod)
          : (hiddenSet as Set<string>).has(filterMatchKey(filterValueKey(cell?.value), locale));
      if (hide) next.add(r);
    }
  }
  return { hidden: next, count: next.size - baseline.size };
}

/** Formulon's AutoFilter text matcher follows Excel and compares text without
 * regard to case. Keep checklist labels in their original case, but fold
 * value keys whenever a selection is evaluated. */
const filterMatchKey = (value: string, locale = 'en-US'): string => profileTextKey(value, locale);

/** Apply an Excel-style "Filter by Selected Cell's Value" action. The active
 *  cell must be inside the active filter range (or current selection range)
 *  and below the header row. The resulting criteria is stored as a normal
 *  value filter so Reapply and sheet-view snapshots keep working. */
export function filterBySelectedCellValue(
  state: State,
  store: SpreadsheetStore,
  range: Range = activeFilterRange(state) ?? inferAutoFilterRange(state),
  wb?: WorkbookHandle,
): number {
  const active = state.selection.active;
  if (
    active.sheet !== range.sheet ||
    active.row <= range.r0 ||
    active.row > range.r1 ||
    active.col < range.c0 ||
    active.col > range.c1
  ) {
    return 0;
  }
  const selected = state.data.cells.get(addrKey(active));
  const selectedValue = selected?.value ?? { kind: 'blank' as const };
  const locale = localeForWorkbook(wb);
  const semantics = createProfileTextSemantics(locale);
  const selectedLabel = filterItemLabel(state, active, selectedValue, locale);
  const distinct = distinctFilterItemsResult(state, range, active.col, locale, wb);
  // Native evaluation refusal/failure must be atomic for this command. An
  // empty checklist would otherwise be interpreted as "hide nothing".
  if (distinct.failure) return 0;
  const hiddenValues = distinct.items
    .filter((item) => !semantics.equals(item.label, selectedLabel))
    .map((item) => item.key);
  if (hiddenValues.length === 0) {
    setAutoFilter(store, range, wb);
    return 0;
  }
  return applyValueFilter(state, store, range, active.col, hiddenValues, wb);
}

const isNonBlankCell = (state: State, sheet: number, row: number, col: number): boolean => {
  const cell = state.data.cells.get(addrKey({ sheet, row, col }));
  return cell != null && cell.value.kind !== 'blank';
};

const hasNonBlankInRow = (
  state: State,
  sheet: number,
  row: number,
  c0: number,
  c1: number,
): boolean => {
  for (let col = c0; col <= c1; col += 1) {
    if (isNonBlankCell(state, sheet, row, col)) return true;
  }
  return false;
};

const hasNonBlankInCol = (
  state: State,
  sheet: number,
  col: number,
  r0: number,
  r1: number,
): boolean => {
  for (let row = r0; row <= r1; row += 1) {
    if (isNonBlankCell(state, sheet, row, col)) return true;
  }
  return false;
};

/** Infer the Excel Current Region for a Filter toggle. A selection inside a
 *  table resolves to that table. Explicit multi-cell
 *  selections are honored as-is; a single active cell expands to the
 *  surrounding contiguous non-blank block bounded by blank rows/columns. */
export function inferAutoFilterRange(state: State, range: Range = state.selection.range): Range {
  // Inside a table the filter belongs to the table, whatever surrounds it.
  const table = state.tables.tables.find((t) => t.showHeader && rangeContainsRange(t.range, range));
  if (table) return { ...table.range };
  if (range.r0 !== range.r1 || range.c0 !== range.c1) return { ...range };
  const sheet = range.sheet;
  const row = range.r0;
  const col = range.c0;
  if (!isNonBlankCell(state, sheet, row, col)) return { ...range };

  let r0 = row;
  let r1 = row;
  let c0 = col;
  let c1 = col;
  let changed = true;
  while (changed) {
    changed = false;
    if (r0 > 0 && hasNonBlankInRow(state, sheet, r0 - 1, c0, c1)) {
      r0 -= 1;
      changed = true;
    }
    if (r1 < MAX_ROW && hasNonBlankInRow(state, sheet, r1 + 1, c0, c1)) {
      r1 += 1;
      changed = true;
    }
    if (c0 > 0 && hasNonBlankInCol(state, sheet, c0 - 1, r0, r1)) {
      c0 -= 1;
      changed = true;
    }
    if (c1 < MAX_COL && hasNonBlankInCol(state, sheet, c1 + 1, r0, r1)) {
      c1 += 1;
      changed = true;
    }
  }

  return { sheet, r0, c0, r1, c1 };
}

type ComparisonOp = '=' | '<>' | '>' | '>=' | '<' | '<=';

type AdvancedCriterion =
  /** Bare text: Excel matches entries that *begin with* the pattern (wildcards
   *  honoured), case-insensitively. */
  | { kind: 'beginsWith'; value: string }
  /** Operator-prefixed text, e.g. `=Smith` (exact), `<>Smith`, `>M` (lexical). */
  | { kind: 'text'; op: ComparisonOp; value: string }
  | { kind: 'number'; op: ComparisonOp; value: number }
  /** `=` alone → is blank. */
  | { kind: 'blank' }
  /** `<>` alone → is not blank. */
  | { kind: 'nonBlank' };

export interface AdvancedFilterCopyOptions {
  uniqueOnly?: boolean;
}

/** Apply an Excel-style Advanced Filter in-place. The list range's first row is
 *  treated as headers. The criteria range must also have headers in its first
 *  row; criteria in the same row are ANDed, while criteria rows are ORed. */
export function applyAdvancedFilter(
  state: State,
  store: SpreadsheetStore,
  listRange: Range,
  criteriaRange: Range,
  wb?: WorkbookHandle,
): number {
  if (!canEvaluateFilterRange(listRange)) {
    stampFilterRangeWithoutCriteria(store, listRange);
    syncFilterToEngine(wb, store, listRange);
    return 0;
  }
  const locale = localeForWorkbook(wb);
  const criteriaRows = buildAdvancedCriteriaRows(state, listRange, criteriaRange, locale);
  if (criteriaRows.length === 0) {
    setAutoFilter(store, listRange, wb);
    return 0;
  }

  const nextHidden = new Set(state.layout.hiddenRows);
  for (let r = listRange.r0 + 1; r <= listRange.r1; r += 1) nextHidden.delete(r);

  let hiddenCount = 0;
  for (let r = listRange.r0 + 1; r <= listRange.r1; r += 1) {
    const rowMatches = criteriaRows.some((criteria) =>
      criteria.every(({ col, criterion }) =>
        advancedCriterionMatches(
          state.data.cells.get(addrKey({ sheet: listRange.sheet, row: r, col })),
          criterion,
          locale,
        ),
      ),
    );
    if (!rowMatches) {
      nextHidden.add(r);
      hiddenCount += 1;
    }
  }

  store.setState((s) => ({
    ...s,
    layout: { ...s.layout, hiddenRows: nextHidden },
    ui: { ...s.ui, filterRanges: withFilterRange(s, listRange) },
  }));
  syncFilterToEngine(wb, store, listRange);
  return hiddenCount;
}

/** Copy an Advanced Filter result to another location. The destination address
 *  is the top-left output cell; the list header row is copied first, followed
 *  by matching data rows. `uniqueOnly` mirrors Excel's "Unique records only". */
export function copyAdvancedFilterResult(
  state: State,
  store: SpreadsheetStore,
  listRange: Range,
  criteriaRange: Range,
  dest: Addr,
  options: AdvancedFilterCopyOptions = {},
  wb?: WorkbookHandle | null,
): number {
  if (!canEvaluateFilterRange(listRange)) return 0;
  const locale = localeForWorkbook(wb);
  const criteriaRows = buildAdvancedCriteriaRows(state, listRange, criteriaRange, locale);
  const width = listRange.c1 - listRange.c0 + 1;
  const rowIndexes: number[] = [listRange.r0];
  const seen = new Set<string>();

  for (let r = listRange.r0 + 1; r <= listRange.r1; r += 1) {
    const rowMatches =
      criteriaRows.length === 0 ||
      criteriaRows.some((criteria) =>
        criteria.every(({ col, criterion }) =>
          advancedCriterionMatches(
            state.data.cells.get(addrKey({ sheet: listRange.sheet, row: r, col })),
            criterion,
            locale,
          ),
        ),
      );
    if (!rowMatches) continue;
    if (options.uniqueOnly) {
      const key = advancedRowKey(state, listRange, r, locale);
      if (seen.has(key)) continue;
      seen.add(key);
    }
    rowIndexes.push(r);
  }

  const output = {
    sheet: dest.sheet,
    r0: dest.row,
    c0: dest.col,
    r1: dest.row + rowIndexes.length - 1,
    c1: dest.col + width - 1,
  };
  if (wb && blockDataTableEdit(wb, output)) return 0;
  const writeMatchedRows = (): void => {
    store.setState((s) => {
      const cells = new Map(s.data.cells);
      rowIndexes.forEach((sourceRow, outOffset) => {
        for (let offset = 0; offset < width; offset += 1) {
          const sourceCol = listRange.c0 + offset;
          const target = { sheet: dest.sheet, row: dest.row + outOffset, col: dest.col + offset };
          const source = state.data.cells.get(
            addrKey({ sheet: listRange.sheet, row: sourceRow, col: sourceCol }),
          );
          if (source) cells.set(addrKey(target), cloneCellRecord(source));
          else cells.delete(addrKey(target));
          if (wb) writeCellRecord(wb, target, source);
        }
      });
      return { ...s, data: { ...s.data, cells } };
    });
  };
  // The engine writes ride along with the store update, so batch them into
  // one recalc instead of one per copied cell.
  if (wb) wb.withBatchedRecalc(writeMatchedRows);
  else writeMatchedRows();

  return rowIndexes.length;
}

function buildAdvancedCriteriaRows(
  state: State,
  listRange: Range,
  criteriaRange: Range,
  locale = 'en-US',
): Array<Array<{ col: number; criterion: AdvancedCriterion }>> {
  const headerToCol = new Map<string, number>();
  for (let c = listRange.c0; c <= listRange.c1; c += 1) {
    const header = cellText(state, listRange.sheet, listRange.r0, c).trim().toLowerCase();
    if (header) headerToCol.set(header, c);
  }

  const criteriaCols: Array<{ col: number; listCol: number }> = [];
  for (let c = criteriaRange.c0; c <= criteriaRange.c1; c += 1) {
    const header = cellText(state, criteriaRange.sheet, criteriaRange.r0, c).trim().toLowerCase();
    const listCol = headerToCol.get(header);
    if (listCol != null) criteriaCols.push({ col: c, listCol });
  }

  const rows: Array<Array<{ col: number; criterion: AdvancedCriterion }>> = [];
  for (let r = criteriaRange.r0 + 1; r <= criteriaRange.r1; r += 1) {
    const criteria: Array<{ col: number; criterion: AdvancedCriterion }> = [];
    for (const { col, listCol } of criteriaCols) {
      const raw = cellText(state, criteriaRange.sheet, r, col).trim();
      if (!raw) continue;
      criteria.push({ col: listCol, criterion: parseAdvancedCriterion(raw, locale) });
    }
    if (criteria.length > 0) rows.push(criteria);
  }
  return rows;
}

function advancedRowKey(state: State, range: Range, row: number, locale = 'en-US'): string {
  const parts: string[] = [];
  for (let col = range.c0; col <= range.c1; col += 1) {
    const value = state.data.cells.get(addrKey({ sheet: range.sheet, row, col }))?.value;
    const key = filterValueKey(value);
    parts.push(
      `${value?.kind ?? 'blank'}:${value?.kind === 'text' ? profileTextKey(key, locale) : key}`,
    );
  }
  return parts.join('\u001f');
}

function cloneCellRecord(cell: { value: CellValue; formula: string | null }): {
  value: CellValue;
  formula: string | null;
} {
  // A data table's TABLE() cells copy as their values.
  return {
    value: { ...cell.value } as CellValue,
    formula: isDataTableFormula(cell.formula) ? null : cell.formula,
  };
}

function writeCellRecord(
  wb: WorkbookHandle,
  addr: Addr,
  cell: { value: CellValue; formula: string | null } | undefined,
): void {
  if (cell) writeCell(wb, addr, cell.value, isDataTableFormula(cell.formula) ? null : cell.formula);
  else wb.setBlank(addr);
}

function parseAdvancedCriterion(raw: string, locale = 'en-US'): AdvancedCriterion {
  const match = /^(<=|>=|<>|=|<|>)(.*)$/.exec(raw);
  if (match) {
    const op = match[1] as ComparisonOp;
    const body = (match[2] ?? '').trim();
    // A bare operator means blank / non-blank (Excel: `=` → is blank).
    if (body === '') return op === '<>' ? { kind: 'nonBlank' } : { kind: 'blank' };
    const numeric = parseLocalizedNumericInput(body, locale);
    if (numeric !== null) return { kind: 'number', op, value: numeric };
    return { kind: 'text', op, value: body };
  }
  // No operator: bare numbers match by equality, bare text is "begins with".
  const numeric = parseLocalizedNumericInput(raw, locale);
  if (raw !== '' && numeric !== null) return { kind: 'number', op: '=', value: numeric };
  return { kind: 'beginsWith', value: raw };
}

function advancedCriterionMatches(
  cell: { value: unknown; formula: string | null } | undefined,
  criterion: AdvancedCriterion,
  locale = 'en-US',
): boolean {
  const text = filterValueKey(cell?.value);
  switch (criterion.kind) {
    case 'blank':
      return text === '';
    case 'nonBlank':
      return text !== '';
    case 'beginsWith':
      return textLikeMatch(text, criterion.value, false, locale);
    case 'text': {
      switch (criterion.op) {
        case '=':
          return textLikeMatch(text, criterion.value, true, locale);
        case '<>':
          return !textLikeMatch(text, criterion.value, true, locale);
        default: {
          const compare = createProfileTextSemantics(locale).compare(text, criterion.value);
          if (criterion.op === '>') return compare > 0;
          if (criterion.op === '>=') return compare >= 0;
          if (criterion.op === '<') return compare < 0;
          return compare <= 0;
        }
      }
    }
    default: {
      const value = cellNumber(cell?.value);
      if (value == null) return false;
      switch (criterion.op) {
        case '<>':
          return value !== criterion.value;
        case '>':
          return value > criterion.value;
        case '>=':
          return value >= criterion.value;
        case '<':
          return value < criterion.value;
        case '<=':
          return value <= criterion.value;
        default:
          return value === criterion.value;
      }
    }
  }
}

/** Case-insensitive wildcard match. `*`/`?` are Excel wildcards; when
 *  `anchorEnd` is false the pattern only has to match the start of `text`
 *  (Excel's "begins with" semantics for bare-text criteria). */
function textLikeMatch(
  text: string,
  pattern: string,
  anchorEnd: boolean,
  locale = 'en-US',
): boolean {
  const loweredPattern = pattern.toLocaleLowerCase(locale);
  const escaped = loweredPattern.replace(/[.+^${}()|[\]\\]/g, '\\$&');
  const body = escaped.replace(/\*/g, '.*').replace(/\?/g, '.');
  const regex = new RegExp(`^${body}${anchorEnd ? '$' : ''}`);
  return regex.test(text.toLocaleLowerCase(locale));
}

function cellText(state: State, sheet: number, row: number, col: number): string {
  const cell = state.data.cells.get(addrKey({ sheet, row, col }));
  return cell ? formatCell(cell.value) : '';
}

export function reapplyFilters(state: State, store: SpreadsheetStore, wb?: WorkbookHandle): number {
  const criteria = state.ui.filterCriteria;
  if (criteria.length === 0) {
    syncFilterToEngine(wb, store);
    return 0;
  }
  let hiddenCount = 0;
  store.setState((s) => ({
    ...s,
    layout: {
      ...s.layout,
      hiddenRows: rowsAfterReapply(
        s,
        criteria,
        (count) => {
          hiddenCount = count;
        },
        localeForWorkbook(wb),
      ),
    },
  }));
  syncFilterToEngine(wb, store);
  return hiddenCount;
}

function rowsAfterReapply(
  state: State,
  criteria: readonly ValueFilterCriteria[],
  setHiddenCount: (count: number) => void,
  locale = 'en-US',
): Set<number> {
  const { hidden, count } = recomputeHiddenFromCriteria(state, criteria, locale);
  setHiddenCount(count);
  return hidden;
}

function upsertCriteria(
  existing: readonly ValueFilterCriteria[],
  next: ValueFilterCriteria,
): ValueFilterCriteria[] {
  const out = existing.filter((c) => !(sameRange(c.range, next.range) && c.byCol === next.byCol));
  // A value filter that hides nothing clears the column; condition/color
  // filters are meaningful even though hiddenValues is empty.
  if (!next.condition && !next.color && next.hiddenValues.length === 0) return out;
  return [
    ...out,
    {
      range: { ...next.range },
      byCol: next.byCol,
      hiddenValues: [...next.hiddenValues],
      ...(next.condition ? { condition: { ...next.condition } } : {}),
      ...(next.color ? { color: { ...next.color } } : {}),
    },
  ];
}

/** The filter range a sort of `range` belongs to: same columns and last row,
 *  with the header row immediately above the sorted rows. */
export function filterRangeForSort(state: State, range: Range, hasHeader: boolean): Range | null {
  const headerRow = hasHeader ? range.r0 : range.r0 - 1;
  return (
    state.ui.filterRanges.find(
      (filter) =>
        filter.sheet === range.sheet &&
        filter.r0 === headerRow &&
        filter.r1 === range.r1 &&
        filter.c0 === range.c0 &&
        filter.c1 === range.c1,
    ) ?? null
  );
}

/** Record the sort order applied inside `filterRange` (an empty `keys` clears
 *  it), undoable with the rest of the sort. */
export function setFilterSort(
  store: SpreadsheetStore,
  history: History | null,
  filterRange: Range,
  keys: FilterSortSpec['keys'],
  wb?: WorkbookHandle,
): void {
  const before = store.getState().ui.filterSorts ?? [];
  const next = [
    ...before.filter((spec) => !sameRange(spec.range, filterRange)),
    ...(keys.length > 0
      ? [{ range: { ...filterRange }, keys: keys.map((key) => ({ ...key })) }]
      : []),
  ];
  const apply = (sorts: FilterSortSpec[]): void => {
    store.setState((s) => ({ ...s, ui: { ...s.ui, filterSorts: sorts } }));
    syncFilterToEngine(wb, store, filterRange);
  };
  apply(next);
  if (history && !history.isReplaying()) {
    history.push({ undo: () => apply(before), redo: () => apply(next) });
  }
}

/** Add `range` as an autofilter region without applying any predicate.
 *  Headers in the range paint the chevron so the user can open the dropdown
 *  per column. Equivalent to the "Filter" toggle (Ctrl+Shift+L). */
export function setAutoFilter(store: SpreadsheetStore, range: Range, wb?: WorkbookHandle): void {
  store.setState((s) => ({ ...s, ui: { ...s.ui, filterRanges: withFilterRange(s, range) } }));
  syncFilterToEngine(wb, store, range);
}

/** Reveal all rows in `range` (or all rows when range is omitted). Also removes
 *  the filter range matching `range` from `ui.filterRanges` (or every filter
 *  range when no range is supplied). */
export function clearFilter(
  state: State,
  store: SpreadsheetStore,
  range?: Range,
  wb?: WorkbookHandle,
): void {
  const next = range ? clearHiddenRowsInRange(state.layout.hiddenRows, range) : new Set<number>();
  store.setState((s) => ({
    ...s,
    layout: { ...s.layout, hiddenRows: next },
    ui: {
      ...s.ui,
      filterRanges: range ? withoutFilterRange(s.ui.filterRanges, range) : [],
      filterCriteria: filterCriteriaAfterClear(s.ui.filterCriteria, range),
    },
  }));
  syncFilterToEngine(wb, store, range);
}

/** Reveal `range`'s rows and drop its criteria, leaving its filter buttons
 *  (present or not) as they are. */
export function resetFilterCriteria(
  store: SpreadsheetStore,
  range: Range,
  wb?: WorkbookHandle,
): void {
  store.setState((s) => ({
    ...s,
    layout: { ...s.layout, hiddenRows: clearHiddenRowsInRange(s.layout.hiddenRows, range) },
    ui: { ...s.ui, filterCriteria: filterCriteriaAfterClear(s.ui.filterCriteria, range) },
  }));
  syncFilterToEngine(wb, store, range);
}

function filterCriteriaAfterClear(
  criteria: readonly ValueFilterCriteria[],
  range?: Range,
): ValueFilterCriteria[] {
  if (!range) return [];
  return criteria.filter((c) => !sameRange(c.range, range));
}
