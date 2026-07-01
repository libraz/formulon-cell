import { addrKey } from '../engine/address.js';
import type { Addr, CellValue, Range } from '../engine/types.js';
import { formatCell } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SpreadsheetStore, State, ValueFilterCriteria } from '../store/store.js';
import type { History } from './history.js';

export type FilterPredicate = (
  cell: { value: unknown; formula: string | null } | undefined,
) => boolean;

export type ConditionFilterOp =
  | 'equals'
  | 'notEquals'
  | 'contains'
  | 'notContains'
  | 'greaterThan'
  | 'greaterThanOrEqual'
  | 'lessThan'
  | 'lessThanOrEqual';

export interface ConditionFilterOptions {
  op: ConditionFilterOp;
  value: string;
}

interface FilterSnapshot {
  hiddenRows: Set<number>;
  filterRange: Range | null;
  filterCriteria: ValueFilterCriteria[];
}

const cloneRange = (range: Range | null): Range | null => (range ? { ...range } : null);
const cloneCriteria = (criteria: readonly ValueFilterCriteria[]): ValueFilterCriteria[] =>
  criteria.map((c) => ({
    range: { ...c.range },
    byCol: c.byCol,
    hiddenValues: [...c.hiddenValues],
    ...(c.condition ? { condition: { ...c.condition } } : {}),
  }));

function captureFilterSnapshot(state: State): FilterSnapshot {
  return {
    hiddenRows: new Set(state.layout.hiddenRows),
    filterRange: cloneRange(state.ui.filterRange),
    filterCriteria: cloneCriteria(state.ui.filterCriteria),
  };
}

function applyFilterSnapshot(store: SpreadsheetStore, snap: FilterSnapshot): void {
  store.setState((s) => ({
    ...s,
    layout: { ...s.layout, hiddenRows: new Set(snap.hiddenRows) },
    ui: {
      ...s.ui,
      filterRange: cloneRange(snap.filterRange),
      filterCriteria: cloneCriteria(snap.filterCriteria),
    },
  }));
}

function sameRange(a: Range | null, b: Range | null): boolean {
  if (a === b) return true;
  if (!a || !b) return false;
  return a.sheet === b.sheet && a.r0 === b.r0 && a.r1 === b.r1 && a.c0 === b.c0 && a.c1 === b.c1;
}

function sameFilterSnapshot(a: FilterSnapshot, b: FilterSnapshot): boolean {
  if (!sameRange(a.filterRange, b.filterRange)) return false;
  if (a.hiddenRows.size !== b.hiddenRows.size) return false;
  if (!sameCriteriaList(a.filterCriteria, b.filterCriteria)) return false;
  for (const row of a.hiddenRows) {
    if (!b.hiddenRows.has(row)) return false;
  }
  return true;
}

function sameCriteriaList(
  a: readonly ValueFilterCriteria[],
  b: readonly ValueFilterCriteria[],
): boolean {
  if (a.length !== b.length) return false;
  return a.every((left, i) => {
    const right = b[i];
    return (
      !!right &&
      left.byCol === right.byCol &&
      sameRange(left.range, right.range) &&
      left.condition?.op === right.condition?.op &&
      left.condition?.value === right.condition?.value &&
      left.hiddenValues.length === right.hiddenValues.length &&
      left.hiddenValues.every((value, index) => value === right.hiddenValues[index])
    );
  });
}

export function recordFilterChange<T>(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => T,
): T {
  const before = captureFilterSnapshot(store.getState());
  const result = mutate();
  const after = captureFilterSnapshot(store.getState());
  if (history && !history.isReplaying() && !sameFilterSnapshot(before, after)) {
    history.push({
      undo: () => applyFilterSnapshot(store, before),
      redo: () => applyFilterSnapshot(store, after),
    });
  }
  return result;
}

/** Hide rows in `range` whose `byCol` value fails `predicate`. The first
 *  row of the range is treated as a header and stays visible.
 *  Hidden rows are added to `layout.hiddenRows` (UI-only — engine state is
 *  untouched). The range is also stamped onto `ui.filterRange` so the column
 *  headers paint the autofilter chevron. */
export function applyFilter(
  state: State,
  store: SpreadsheetStore,
  range: Range,
  byCol: number,
  predicate: FilterPredicate,
): number {
  const newHidden = new Set<number>(state.layout.hiddenRows);
  let hiddenCount = 0;
  for (let r = range.r0 + 1; r <= range.r1; r += 1) {
    const cell = state.data.cells.get(addrKey({ sheet: range.sheet, row: r, col: byCol }));
    if (!predicate(cell)) {
      if (!newHidden.has(r)) {
        newHidden.add(r);
        hiddenCount += 1;
      }
    }
  }
  store.setState((s) => ({
    ...s,
    layout: { ...s.layout, hiddenRows: newHidden },
    ui: { ...s.ui, filterRange: { ...range } },
  }));
  return hiddenCount;
}

export function applyValueFilter(
  state: State,
  store: SpreadsheetStore,
  range: Range,
  byCol: number,
  hiddenValues: readonly string[],
): number {
  const criteria = upsertCriteria(state.ui.filterCriteria, {
    range,
    byCol,
    hiddenValues: Array.from(new Set(hiddenValues)).sort(),
  });
  const { hidden, count } = recomputeHiddenFromCriteria(state, criteria);
  store.setState((s) => ({
    ...s,
    layout: { ...s.layout, hiddenRows: hidden },
    ui: { ...s.ui, filterRange: { ...range }, filterCriteria: criteria },
  }));
  return count;
}

export function applyConditionFilter(
  state: State,
  store: SpreadsheetStore,
  range: Range,
  byCol: number,
  condition: ConditionFilterOptions,
): number {
  const criteria = upsertCriteria(state.ui.filterCriteria, {
    range,
    byCol,
    hiddenValues: [],
    condition: { op: condition.op, value: condition.value },
  });
  const { hidden, count } = recomputeHiddenFromCriteria(state, criteria);
  store.setState((s) => ({
    ...s,
    layout: { ...s.layout, hiddenRows: hidden },
    ui: { ...s.ui, filterRange: { ...range }, filterCriteria: criteria },
  }));
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
): { hidden: Set<number>; count: number } {
  const baseline = new Set(state.layout.hiddenRows);
  for (const c of criteria) {
    for (let r = c.range.r0 + 1; r <= c.range.r1; r += 1) baseline.delete(r);
  }
  const next = new Set(baseline);
  for (const c of criteria) {
    const hiddenSet = c.condition ? null : new Set(c.hiddenValues);
    for (let r = c.range.r0 + 1; r <= c.range.r1; r += 1) {
      if (next.has(r)) continue;
      const cell = state.data.cells.get(addrKey({ sheet: c.range.sheet, row: r, col: c.byCol }));
      const hide = c.condition
        ? !conditionMatches(cell, c.condition)
        : (hiddenSet as Set<string>).has(filterValueKey(cell?.value));
      if (hide) next.add(r);
    }
  }
  return { hidden: next, count: next.size - baseline.size };
}

/** Apply an Excel-style "Filter by Selected Cell's Value" action. The active
 *  cell must be inside the active filter range (or current selection range)
 *  and below the header row. The resulting criteria is stored as a normal
 *  value filter so Reapply and sheet-view snapshots keep working. */
export function filterBySelectedCellValue(
  state: State,
  store: SpreadsheetStore,
  range: Range = state.ui.filterRange ?? inferAutoFilterRange(state),
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
  const selectedKey = filterValueKey(selected?.value);
  const hiddenValues = distinctValues(state, range, active.col).filter(
    (value) => value !== selectedKey,
  );
  if (hiddenValues.length === 0) {
    setAutoFilter(store, range);
    return 0;
  }
  return applyValueFilter(state, store, range, active.col, hiddenValues);
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

/** Infer the Excel Current Region for a Filter toggle. Explicit multi-cell
 *  selections are honored as-is; a single active cell expands to the
 *  surrounding contiguous non-blank block bounded by blank rows/columns. */
export function inferAutoFilterRange(state: State, range: Range = state.selection.range): Range {
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
    if (r1 < 1048575 && hasNonBlankInRow(state, sheet, r1 + 1, c0, c1)) {
      r1 += 1;
      changed = true;
    }
    if (c0 > 0 && hasNonBlankInCol(state, sheet, c0 - 1, r0, r1)) {
      c0 -= 1;
      changed = true;
    }
    if (c1 < 16383 && hasNonBlankInCol(state, sheet, c1 + 1, r0, r1)) {
      c1 += 1;
      changed = true;
    }
  }

  return { sheet, r0, c0, r1, c1 };
}

type AdvancedCriterion =
  | { kind: 'text'; value: string }
  | { kind: 'number'; op: '=' | '<>' | '>' | '>=' | '<' | '<='; value: number }
  | { kind: 'notText'; value: string };

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
): number {
  const criteriaRows = buildAdvancedCriteriaRows(state, listRange, criteriaRange);
  if (criteriaRows.length === 0) {
    setAutoFilter(store, listRange);
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
    ui: { ...s.ui, filterRange: { ...listRange } },
  }));
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
  const criteriaRows = buildAdvancedCriteriaRows(state, listRange, criteriaRange);
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
          ),
        ),
      );
    if (!rowMatches) continue;
    if (options.uniqueOnly) {
      const key = advancedRowKey(state, listRange, r);
      if (seen.has(key)) continue;
      seen.add(key);
    }
    rowIndexes.push(r);
  }

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

  return rowIndexes.length;
}

function buildAdvancedCriteriaRows(
  state: State,
  listRange: Range,
  criteriaRange: Range,
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
      criteria.push({ col: listCol, criterion: parseAdvancedCriterion(raw) });
    }
    if (criteria.length > 0) rows.push(criteria);
  }
  return rows;
}

function advancedRowKey(state: State, range: Range, row: number): string {
  const parts: string[] = [];
  for (let col = range.c0; col <= range.c1; col += 1) {
    parts.push(
      filterValueKey(state.data.cells.get(addrKey({ sheet: range.sheet, row, col }))?.value),
    );
  }
  return parts.join('\u001f');
}

function cloneCellRecord(cell: { value: CellValue; formula: string | null }): {
  value: CellValue;
  formula: string | null;
} {
  return { value: { ...cell.value } as CellValue, formula: cell.formula };
}

function writeCellRecord(
  wb: WorkbookHandle,
  addr: Addr,
  cell: { value: CellValue; formula: string | null } | undefined,
): void {
  if (!cell) {
    wb.setBlank(addr);
    return;
  }
  if (cell.formula) {
    wb.setFormula(addr, cell.formula);
    return;
  }
  const value = cell.value;
  if (value.kind === 'number') wb.setNumber(addr, value.value);
  else if (value.kind === 'text') wb.setText(addr, value.value);
  else if (value.kind === 'bool') wb.setBool(addr, value.value);
  else wb.setBlank(addr);
}

function parseAdvancedCriterion(raw: string): AdvancedCriterion {
  const match = /^(<=|>=|<>|=|<|>)(.*)$/.exec(raw);
  const op = (match?.[1] ?? '=') as AdvancedCriterion extends { kind: 'number'; op: infer Op }
    ? Op
    : never;
  const body = (match?.[2] ?? raw).trim();
  const numeric = Number(body);
  if (body !== '' && Number.isFinite(numeric)) return { kind: 'number', op, value: numeric };
  if (op === '<>') return { kind: 'notText', value: body };
  return { kind: 'text', value: body };
}

function advancedCriterionMatches(
  cell: { value: unknown; formula: string | null } | undefined,
  criterion: AdvancedCriterion,
): boolean {
  const text = filterValueKey(cell?.value);
  if (criterion.kind === 'text') return textMatchesAdvancedCriterion(text, criterion.value);
  if (criterion.kind === 'notText') return !textMatchesAdvancedCriterion(text, criterion.value);
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

function textMatchesAdvancedCriterion(text: string, pattern: string): boolean {
  if (!pattern.includes('*') && !pattern.includes('?')) return text === pattern;
  const escaped = pattern.replace(/[.+^${}()|[\]\\]/g, '\\$&');
  const regex = new RegExp(`^${escaped.replace(/\*/g, '.*').replace(/\?/g, '.')}$`, 'i');
  return regex.test(text);
}

function cellText(state: State, sheet: number, row: number, col: number): string {
  const cell = state.data.cells.get(addrKey({ sheet, row, col }));
  return cell ? formatCell(cell.value) : '';
}

function cellNumber(value: unknown): number | null {
  if (!value || typeof value !== 'object') return null;
  const cv = value as { kind: string; value?: unknown };
  if (cv.kind !== 'number' || typeof cv.value !== 'number' || !Number.isFinite(cv.value)) {
    return null;
  }
  return cv.value;
}

function conditionMatches(
  cell: { value: unknown; formula: string | null } | undefined,
  condition: ConditionFilterOptions,
): boolean {
  const needle = condition.value.trim();
  const text = filterValueKey(cell?.value);
  switch (condition.op) {
    case 'notEquals':
      return text !== needle;
    case 'contains':
      return text.toLowerCase().includes(needle.toLowerCase());
    case 'notContains':
      return !text.toLowerCase().includes(needle.toLowerCase());
    case 'greaterThan':
    case 'greaterThanOrEqual':
    case 'lessThan':
    case 'lessThanOrEqual': {
      const target = Number(needle);
      const value = cellNumber(cell?.value);
      if (!Number.isFinite(target) || value == null) return false;
      if (condition.op === 'greaterThan') return value > target;
      if (condition.op === 'greaterThanOrEqual') return value >= target;
      if (condition.op === 'lessThan') return value < target;
      return value <= target;
    }
    default:
      return text === needle;
  }
}

export function reapplyFilters(state: State, store: SpreadsheetStore): number {
  const criteria = state.ui.filterCriteria;
  if (criteria.length === 0) return 0;
  let hiddenCount = 0;
  store.setState((s) => ({
    ...s,
    layout: {
      ...s.layout,
      hiddenRows: rowsAfterReapply(s, criteria, (count) => {
        hiddenCount = count;
      }),
    },
  }));
  return hiddenCount;
}

function rowsAfterReapply(
  state: State,
  criteria: readonly ValueFilterCriteria[],
  setHiddenCount: (count: number) => void,
): Set<number> {
  const { hidden, count } = recomputeHiddenFromCriteria(state, criteria);
  setHiddenCount(count);
  return hidden;
}

function upsertCriteria(
  existing: readonly ValueFilterCriteria[],
  next: ValueFilterCriteria,
): ValueFilterCriteria[] {
  const out = existing.filter((c) => !(sameRange(c.range, next.range) && c.byCol === next.byCol));
  // A value filter that hides nothing clears the column; a condition filter is
  // always meaningful, so it is kept even though hiddenValues is empty.
  if (!next.condition && next.hiddenValues.length === 0) return out;
  return [
    ...out,
    {
      range: { ...next.range },
      byCol: next.byCol,
      hiddenValues: [...next.hiddenValues],
      ...(next.condition ? { condition: { ...next.condition } } : {}),
    },
  ];
}

/** Mark `range` as the active autofilter region without applying any predicate.
 *  Headers in the range paint the chevron so the user can open the dropdown
 *  per column. Equivalent to the "Filter" toggle (Ctrl+Shift+L). */
export function setAutoFilter(store: SpreadsheetStore, range: Range | null): void {
  store.setState((s) => ({ ...s, ui: { ...s.ui, filterRange: range ? { ...range } : null } }));
}

/** Reveal all rows in `range` (or all rows when range is omitted). Also clears
 *  `ui.filterRange` when `range` matches the active autofilter region (or when
 *  no range is supplied). */
export function clearFilter(state: State, store: SpreadsheetStore, range?: Range): void {
  const next = new Set<number>(state.layout.hiddenRows);
  if (!range) {
    next.clear();
  } else {
    for (let r = range.r0; r <= range.r1; r += 1) next.delete(r);
  }
  store.setState((s) => {
    const fr = s.ui.filterRange;
    const clearFr =
      !range ||
      (fr != null &&
        fr.sheet === range.sheet &&
        fr.r0 === range.r0 &&
        fr.r1 === range.r1 &&
        fr.c0 === range.c0 &&
        fr.c1 === range.c1);
    return {
      ...s,
      layout: { ...s.layout, hiddenRows: next },
      ui: {
        ...s.ui,
        ...(clearFr ? { filterRange: null } : {}),
        filterCriteria: filterCriteriaAfterClear(s.ui.filterCriteria, range),
      },
    };
  });
}

function filterCriteriaAfterClear(
  criteria: readonly ValueFilterCriteria[],
  range?: Range,
): ValueFilterCriteria[] {
  if (!range) return [];
  return criteria.filter((c) => !sameRange(c.range, range));
}

/** Distinct (string-coerced) values found in `byCol` of `range`, excluding
 *  the header row. Useful for building a dropdown of filter options. */
export function distinctValues(state: State, range: Range, byCol: number): string[] {
  const out = new Set<string>();
  for (let r = range.r0 + 1; r <= range.r1; r += 1) {
    const cell = state.data.cells.get(addrKey({ sheet: range.sheet, row: r, col: byCol }));
    if (!cell) {
      out.add('');
      continue;
    }
    out.add(filterValueKey(cell.value));
  }
  return Array.from(out).sort();
}

export const filterValueKey = (v: unknown): string => {
  if (!v || typeof v !== 'object') return '';
  const cv = v as { kind: string; value?: unknown };
  if (cv.kind === 'number') return String(cv.value);
  if (cv.kind === 'text') return String(cv.value ?? '');
  if (cv.kind === 'bool') return cv.value ? 'TRUE' : 'FALSE';
  return '';
};
