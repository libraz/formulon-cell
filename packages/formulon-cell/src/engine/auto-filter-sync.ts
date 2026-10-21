import type { AutoFilter, AutoFilterColumn, AutoFilterEvalResult } from '@libraz/formulon';
import { inputLocaleForWorkbook, parseLocalizedNumericInput } from '../commands/coerce-input.js';
import { distinctValues } from '../commands/filter.js';
import { sameRange } from '../store/selection-geometry.js';
import type { SpreadsheetStore, State, ValueFilterCriteria } from '../store/store.js';
import { formatA1Range } from './address.js';
import { profileTextKey } from './profile-text-semantics.js';
import { parseRangeRef } from './range-resolver.js';
import type { Range } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

const escapeXml = (value: string): string =>
  value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const conditionOperator: Record<NonNullable<ValueFilterCriteria['condition']>['op'], string> = {
  equals: 'equal',
  notEquals: 'notEqual',
  contains: 'equal',
  notContains: 'notEqual',
  greaterThan: 'greaterThan',
  greaterThanOrEqual: 'greaterThanOrEqual',
  lessThan: 'lessThan',
  lessThanOrEqual: 'lessThanOrEqual',
};

const FILTER_KIND_VALUES = 1;
const FILTER_KIND_CUSTOM = 2;
const FILTER_OPERATOR_EQUAL = 0;
const FILTER_OPERATOR_LESS_THAN = 1;
const FILTER_OPERATOR_LESS_THAN_OR_EQUAL = 2;
const FILTER_OPERATOR_NOT_EQUAL = 3;
const FILTER_OPERATOR_GREATER_THAN_OR_EQUAL = 4;
const FILTER_OPERATOR_GREATER_THAN = 5;

export type ImportedTypedFilter = {
  filter: AutoFilter;
  /** Columns whose criterion the cell UI cannot represent. Their typed
   * definition is carried forward until the user changes that filter. */
  unsupportedColumns: Set<number>;
};

/** One engine-side AutoFilter definition: the worksheet's own, or a table's.
 * Both are driven through the same operations so the two paths stay shaped
 * the same. */
interface FilterTarget {
  key: string;
  get(): AutoFilter | null;
  set(filter: AutoFilter): boolean;
  remove(): boolean;
  apply(): boolean;
  clear(): boolean;
  evaluate(): AutoFilterEvalResult | null;
}

const sheetFilterTarget = (wb: WorkbookHandle, sheet: number): FilterTarget => ({
  key: `sheet:${sheet}`,
  get: () => wb.getAutoFilter(sheet),
  set: (filter) => wb.setAutoFilter(sheet, filter),
  remove: () => wb.removeAutoFilter(sheet),
  apply: () => wb.applyAutoFilter(sheet),
  clear: () => wb.clearAutoFilter(sheet),
  evaluate: () =>
    typeof wb.evaluateAutoFilter === 'function' ? wb.evaluateAutoFilter(sheet) : null,
});

const tableFilterTarget = (wb: WorkbookHandle, table: number): FilterTarget => ({
  key: `table:${table}`,
  get: () => wb.getTableAutoFilter(table),
  set: (filter) => wb.setTableAutoFilter(table, filter),
  remove: () => wb.removeTableAutoFilter(table),
  apply: () => wb.applyTableAutoFilter(table),
  clear: () => wb.clearTableAutoFilter(table),
  evaluate: () => wb.evaluateTableAutoFilter(table),
});

/** One engine AutoFilter slot on a sheet and the store range driving it, or
 * null when the store has no filter there. Excel keeps a table's AutoFilter
 * separate from the worksheet's, so a range on a table's header row and
 * columns belongs to that table (its filter skips a totals row); the remaining
 * range is the worksheet's own. */
interface FilterSlot {
  target: FilterTarget;
  table: boolean;
  range: Range | null;
}

function filterSlotsForSheet(wb: WorkbookHandle, state: State, sheet: number): FilterSlot[] {
  const ranges = state.ui.filterRanges.filter((range) => range.sheet === sheet);
  const claimed = new Set<Range>();
  const slots: FilterSlot[] = [];
  if (wb.capabilities.typedTableAutoFilter) {
    for (const table of wb.getTables()) {
      if (table.sheetIndex !== sheet) continue;
      const parsed = parseRangeRef(table.ref);
      const range = parsed
        ? (ranges.find(
            (r) =>
              r.r0 === parsed.r0 && r.c0 === parsed.c0 && r.c1 === parsed.c1 && r.r1 <= parsed.r1,
          ) ?? null)
        : null;
      if (range) claimed.add(range);
      slots.push({ target: tableFilterTarget(wb, table.index), table: true, range });
    }
  }
  slots.push({
    target: sheetFilterTarget(wb, sheet),
    table: false,
    range: ranges.find((range) => !claimed.has(range)) ?? null,
  });
  return slots;
}

const importedTypedFilters = new WeakMap<WorkbookHandle, Map<string, ImportedTypedFilter>>();

const emptyAutoFilterColumn = (colId: number): AutoFilterColumn => ({
  colId,
  hiddenButton: false,
  showButton: true,
  kind: 0,
  filterBlank: false,
  values: [],
  dateGroups: [],
  customAnd: false,
  customCount: 0,
  op1: FILTER_OPERATOR_EQUAL,
  val1: '',
  op2: FILTER_OPERATOR_EQUAL,
  val2: '',
  top: false,
  percent: false,
  hasFilterVal: false,
  topVal: 0,
  filterVal: 0,
  dynamicType: 0,
  hasDynVal: false,
  hasDynMaxVal: false,
  dynVal: 0,
  dynMaxVal: 0,
  valIso: '',
  maxValIso: '',
  dxfId: 0xffffffff,
  cellColor: false,
  iconSet: 0,
  iconId: 0,
  hasIconId: false,
});

const rangeFromAutoFilter = (filter: AutoFilter, sheet: number): Range => ({
  sheet,
  r0: filter.range.firstRow,
  c0: filter.range.firstCol,
  r1: filter.range.lastRow,
  c1: filter.range.lastCol,
});

const sameAutoFilterRange = (filter: AutoFilter, range: Range): boolean =>
  filter.range.firstRow === range.r0 &&
  filter.range.firstCol === range.c0 &&
  filter.range.lastRow === range.r1 &&
  filter.range.lastCol === range.c1;

function importedFiltersFor(wb: WorkbookHandle): Map<string, ImportedTypedFilter> {
  let filters = importedTypedFilters.get(wb);
  if (!filters) {
    filters = new Map();
    importedTypedFilters.set(wb, filters);
  }
  return filters;
}

/** Drop the imported definition of the removed table `index` and re-key the
 * later tables, whose workbook-wide indexes shift down by one. */
export function forgetTableAutoFilterImport(wb: WorkbookHandle, index: number): void {
  const filters = importedFiltersFor(wb);
  const shifted: [string, ImportedTypedFilter][] = [];
  for (const [key, imported] of filters) {
    if (!key.startsWith('table:')) continue;
    const table = Number(key.slice('table:'.length));
    if (table < index) continue;
    filters.delete(key);
    if (table > index) shifted.push([`table:${table - 1}`, imported]);
  }
  for (const [key, imported] of shifted) filters.set(key, imported);
}

/** Escape a UI contains value for formulon's Excel wildcard matcher. The
 * tilde must be escaped first so the escapes added for `*` and `?` remain
 * intact when the pattern is evaluated by the engine. */
function escapeAutoFilterWildcardLiteral(value: string): string {
  return value.replace(/~/g, '~~').replace(/\*/g, '~*').replace(/\?/g, '~?');
}

function containsPattern(value: string): string {
  return `*${escapeAutoFilterWildcardLiteral(value)}*`;
}

const filterMatchKey = (value: string, locale = 'en-US'): string => profileTextKey(value, locale);

/** Typed AutoFilter values are serialized by the engine in its canonical
 * decimal grammar even when the active profile uses a comma decimal mark.
 * Prefer the active profile for imported/localized payloads, then accept that
 * canonical wire form when rehydrating the engine-owned definition. UI input
 * still goes through the strict locale parser at its entry point. */
const parseStoredNumeric = (value: string, locale: string): number | null =>
  parseLocalizedNumericInput(value, locale) ?? parseLocalizedNumericInput(value, 'en-US');

function decodeWildcardLiteral(value: string): string | null {
  let decoded = '';
  for (let i = 0; i < value.length; i += 1) {
    const ch = value[i] ?? '';
    if (ch === '~') {
      const next = value[i + 1];
      if (next !== '~' && next !== '*' && next !== '?') return null;
      decoded += next;
      i += 1;
      continue;
    }
    if (ch === '*' || ch === '?') return null;
    decoded += ch;
  }
  return decoded;
}

/** Decode the single outer-wildcard pattern emitted by this adapter. A
 * pattern containing another unescaped wildcard is left engine-owned because
 * the cell-layer `contains` condition has no equivalent for it. */
function decodeContainsPattern(value: string): string | null {
  if (value.length < 2 || !value.startsWith('*') || !value.endsWith('*')) return null;

  // The last star is the outer wildcard only when it is not escaped. Count
  // the immediately preceding tildes: an odd run escapes the star, while an
  // even run consists of complete `~~` literal-tilde pairs.
  let tildes = 0;
  for (let i = value.length - 2; i >= 0 && value[i] === '~'; i -= 1) tildes += 1;
  if (tildes % 2 === 1) return null;

  return decodeWildcardLiteral(value.slice(1, -1));
}

function customColumnFromCriteria(
  colId: number,
  condition: NonNullable<ValueFilterCriteria['condition']>,
): AutoFilterColumn {
  const base = emptyAutoFilterColumn(colId);
  const wildcard = condition.op === 'contains' || condition.op === 'notContains';
  const literalWildcard = condition.op === 'equals' || condition.op === 'notEquals';
  const op =
    condition.op === 'equals' || condition.op === 'contains'
      ? FILTER_OPERATOR_EQUAL
      : condition.op === 'notEquals' || condition.op === 'notContains'
        ? FILTER_OPERATOR_NOT_EQUAL
        : condition.op === 'greaterThan'
          ? FILTER_OPERATOR_GREATER_THAN
          : condition.op === 'greaterThanOrEqual'
            ? FILTER_OPERATOR_GREATER_THAN_OR_EQUAL
            : condition.op === 'lessThan'
              ? FILTER_OPERATOR_LESS_THAN
              : FILTER_OPERATOR_LESS_THAN_OR_EQUAL;
  return {
    ...base,
    kind: FILTER_KIND_CUSTOM,
    customCount: 1,
    op1: op,
    val1: wildcard
      ? containsPattern(condition.value)
      : literalWildcard
        ? escapeAutoFilterWildcardLiteral(condition.value)
        : condition.value,
  };
}

/** Convert the cell-layer's criteria to formulon's typed AutoFilter model.
 * Returns `null` for a color criterion when there is no imported typed
 * snapshot to carry forward; the cell layer cannot supply the engine's
 * differential-style id for a new color filter. */
export function typedAutoFilterFromState(
  state: State,
  range: Range | null,
  imported?: ImportedTypedFilter,
  locale = 'en-US',
  wb?: WorkbookHandle,
): AutoFilter | null {
  if (!range) return null;
  const criteria = state.ui.filterCriteria
    .filter((entry) => sameRange(entry.range, range))
    .sort((a, b) => a.byCol - b.byCol);
  const hasColorCriteria = criteria.some((entry) => entry.color);
  // A color criterion has no cell-layer differential-style id. When an
  // imported typed definition is available, retain its engine-owned columns
  // and simply omit the new color criterion from this typed write. Without a
  // typed snapshot the raw XML fallback below is the only safe path.
  if (hasColorCriteria && !imported) return null;

  const columns = new Map<number, AutoFilterColumn>();
  // Imported Dynamic/Top10/Icon/Color criteria are kept intact while the UI
  // edits a representable column. A subsequent clear of the whole filter
  // removes them along with the definition.
  if (imported && sameAutoFilterRange(imported.filter, range)) {
    for (const column of imported.filter.columns) {
      if (imported.unsupportedColumns.has(column.colId)) columns.set(column.colId, column);
    }
  }

  for (const entry of criteria) {
    if (entry.color) continue;
    const colId = entry.byCol - range.c0;
    if (colId < 0 || colId > range.c1 - range.c0) continue;
    if (entry.condition) {
      columns.set(colId, customColumnFromCriteria(colId, entry.condition));
      continue;
    }
    const hidden = new Set(entry.hiddenValues.map((value) => filterMatchKey(value, locale)));
    const column = emptyAutoFilterColumn(colId);
    column.kind = FILTER_KIND_VALUES;
    column.filterBlank = !hidden.has('');
    column.values = distinctValues(state, range, entry.byCol, locale, wb).filter(
      (value) => value !== '' && !hidden.has(filterMatchKey(value, locale)),
    );
    columns.set(colId, column);
  }

  return {
    range: {
      firstRow: range.r0,
      firstCol: range.c0,
      lastRow: range.r1,
      lastCol: range.c1,
    },
    columns: Array.from(columns.values()).sort((a, b) => a.colId - b.colId),
    sort: imported && sameAutoFilterRange(imported.filter, range) ? imported.filter.sort : null,
  };
}

function wildcardCondition(
  op: number,
  value: string,
  locale = 'en-US',
): NonNullable<ValueFilterCriteria['condition']> | null {
  const decoded = decodeContainsPattern(value);
  if (decoded !== null) {
    if (op === FILTER_OPERATOR_EQUAL) return { op: 'contains', value: decoded };
    if (op === FILTER_OPERATOR_NOT_EQUAL) return { op: 'notContains', value: decoded };
    return null;
  }
  const literal = decodeWildcardLiteral(value);
  if (literal !== null) {
    if (op === FILTER_OPERATOR_EQUAL || op === FILTER_OPERATOR_NOT_EQUAL) {
      const numeric = parseStoredNumeric(literal, locale);
      if (numeric !== null) {
        return {
          op: op === FILTER_OPERATOR_EQUAL ? 'equals' : 'notEquals',
          value: String(numeric),
        };
      }
    }
    if (op === FILTER_OPERATOR_EQUAL) return { op: 'equals', value: literal };
    if (op === FILTER_OPERATOR_NOT_EQUAL) return { op: 'notEquals', value: literal };
  }
  if (/[?*]/.test(value)) return null;
  const storedNumeric = parseStoredNumeric(value, locale);
  switch (op) {
    case FILTER_OPERATOR_EQUAL:
      return { op: 'equals', value };
    case FILTER_OPERATOR_NOT_EQUAL:
      return { op: 'notEquals', value };
    case FILTER_OPERATOR_GREATER_THAN:
      return storedNumeric === null ? null : { op: 'greaterThan', value: String(storedNumeric) };
    case FILTER_OPERATOR_GREATER_THAN_OR_EQUAL:
      return storedNumeric === null
        ? null
        : {
            op: 'greaterThanOrEqual',
            value: String(storedNumeric),
          };
    case FILTER_OPERATOR_LESS_THAN:
      return storedNumeric === null ? null : { op: 'lessThan', value: String(storedNumeric) };
    case FILTER_OPERATOR_LESS_THAN_OR_EQUAL:
      return storedNumeric === null
        ? null
        : {
            op: 'lessThanOrEqual',
            value: String(storedNumeric),
          };
    default:
      return null;
  }
}

function localCriteriaFromTypedFilter(
  state: State,
  sheet: number,
  filter: AutoFilter,
  locale = 'en-US',
  wb?: WorkbookHandle,
): { range: Range; criteria: ValueFilterCriteria[]; unsupportedColumns: Set<number> } {
  const range = rangeFromAutoFilter(filter, sheet);
  const criteria: ValueFilterCriteria[] = [];
  const unsupportedColumns = new Set<number>();
  for (const column of filter.columns) {
    const byCol = range.c0 + column.colId;
    if (byCol < range.c0 || byCol > range.c1) {
      unsupportedColumns.add(column.colId);
      continue;
    }
    // The dropdown affordance flags have no cell-layer representation. Keep
    // their native column intact until that column is explicitly edited.
    if (column.hiddenButton || !column.showButton) {
      unsupportedColumns.add(column.colId);
      continue;
    }
    if (column.kind === FILTER_KIND_VALUES && column.dateGroups.length === 0) {
      const visible = new Set(column.values.map((value) => filterMatchKey(value, locale)));
      const hiddenValues = distinctValues(state, range, byCol, locale, wb).filter((value) =>
        value === '' ? !column.filterBlank : !visible.has(filterMatchKey(value, locale)),
      );
      criteria.push({ range, byCol, hiddenValues });
      continue;
    }
    if (column.kind === FILTER_KIND_CUSTOM && column.customCount === 1 && !column.customAnd) {
      const condition = wildcardCondition(column.op1, column.val1, locale);
      if (condition) {
        criteria.push({ range, byCol, hiddenValues: [], condition });
        continue;
      }
    }
    unsupportedColumns.add(column.colId);
  }
  return { range, criteria, unsupportedColumns };
}

/** Serialises the application-owned subset of AutoFilter: its range, value
 * checklist and one custom comparison per column. Color criteria cannot be
 * represented without a differential-style id, so the caller deliberately
 * leaves them out rather than emitting a misleading OOXML definition. */
export function autoFilterXmlFromState(
  state: State,
  range: Range | null,
  locale = 'en-US',
  wb?: WorkbookHandle,
): string {
  if (!range) return '';
  const criteria = state.ui.filterCriteria
    .filter((entry) => sameRange(entry.range, range))
    .sort((a, b) => a.byCol - b.byCol);
  if (criteria.length === 0)
    return `<autoFilter ref="${formatA1Range(range, { collapse: false })}"/>`;

  const columns = criteria.flatMap((entry) => {
    const colId = entry.byCol - range.c0;
    if (colId < 0 || colId > range.c1 - range.c0 || entry.color) return [];
    if (entry.condition) {
      const wildcard = entry.condition.op === 'contains' || entry.condition.op === 'notContains';
      const literalWildcard = entry.condition.op === 'equals' || entry.condition.op === 'notEquals';
      const value = wildcard
        ? containsPattern(entry.condition.value)
        : literalWildcard
          ? escapeAutoFilterWildcardLiteral(entry.condition.value)
          : entry.condition.value;
      return [
        `<filterColumn colId="${colId}"><customFilters><customFilter operator="${conditionOperator[entry.condition.op]}" val="${escapeXml(value)}"/></customFilters></filterColumn>`,
      ];
    }
    const hidden = new Set(entry.hiddenValues.map((value) => filterMatchKey(value, locale)));
    const visible = distinctValues(state, range, entry.byCol, locale, wb).filter(
      (value) => !hidden.has(filterMatchKey(value, locale)),
    );
    const includeBlank = visible.includes('');
    const filters = visible
      .filter((value) => value !== '')
      .map((value) => `<filter val="${escapeXml(value)}"/>`)
      .join('');
    const blankAttr = includeBlank ? ' blank="1"' : '';
    return [
      `<filterColumn colId="${colId}"><filters${blankAttr}>${filters}</filters></filterColumn>`,
    ];
  });
  return `<autoFilter ref="${formatA1Range(range, { collapse: false })}">${columns.join('')}</autoFilter>`;
}

/** Parses only the range, which is enough to restore header affordances while
 * leaving every imported criterion in the engine's verbatim XML until the
 * user changes the filter. */
export function autoFilterRangeFromXml(xml: string, sheet: number): Range | null {
  const ref = /<autoFilter\b[^>]*\bref="([^"]+)"/i.exec(xml)?.[1];
  const parsed = ref ? parseRangeRef(ref) : null;
  return parsed ? { sheet, r0: parsed.r0, c0: parsed.c0, r1: parsed.r1, c1: parsed.c1 } : null;
}

/** Apply native row visibility to the store after the engine evaluates a
 * typed filter. `store` is optional so legacy callers can keep using the
 * save-only sync path while a mount with a store gets authoritative matches. */
export function syncNativeAutoFilterVisibility(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
  target: Range,
  filter: FilterTarget = sheetFilterTarget(wb, sheet),
): boolean {
  if (!wb.capabilities.typedAutoFilter) return false;
  if (typeof wb.evaluateAutoFilter !== 'function') return false;
  const evaluated = filter.evaluate();
  if (!evaluated) return false;
  const nextHidden = new Set(store.getState().layout.hiddenRows);
  for (let row = target.r0 + 1; row <= target.r1; row += 1) nextHidden.delete(row);
  for (let i = 0; i < evaluated.match.length; i += 1) {
    if (!evaluated.match[i]) nextHidden.add(evaluated.firstRow + i);
  }
  store.setState((state) => ({ ...state, layout: { ...state.layout, hiddenRows: nextHidden } }));
  return true;
}

/** Mirror an exact store row mask without asking the native AutoFilter to
 * re-evaluate its criteria. This is used by filter history and by repeated
 * JS-owned filters whose definition has not changed. */
export function syncHiddenRowsToEngine(
  wb: WorkbookHandle,
  sheet: number,
  before: ReadonlySet<number>,
  after: ReadonlySet<number>,
): boolean {
  if (!wb.capabilities.hiddenRowsCols) return false;
  const rows = new Set<number>();
  for (const row of before) rows.add(row);
  for (const row of after) rows.add(row);
  for (const row of rows) wb.setRowHidden(sheet, row, after.has(row));
  return true;
}

/**
 * A color criterion, an advanced filter, or a slicer can own visibility in
 * the cell layer without having an equivalent typed AutoFilter column. Keep
 * the native filter definition (and any imported columns) evaluated, then
 * merge that result with the store-owned row mask before saving. Calling
 * `applyAutoFilter` alone would show every row for an empty definition and
 * would drop the rows hidden by a color criterion from the workbook.
 */
function syncStoreOwnedAutoFilterVisibility(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
  range: Range,
  filter: FilterTarget = sheetFilterTarget(wb, sheet),
): boolean {
  if (!wb.capabilities.hiddenRowsCols) return false;

  const storeHidden = new Set(store.getState().layout.hiddenRows);
  const evaluated = filter.evaluate();
  const firstRow = evaluated?.firstRow ?? range.r0;

  for (let row = range.r0 + 1; row <= range.r1; row += 1) {
    const index = row - firstRow;
    const nativeHidden = evaluated ? evaluated.match[index] === false : false;
    wb.setRowHidden(sheet, row, nativeHidden || storeHidden.has(row));
  }
  return true;
}

const emptyTypedFilter = (range: Range): AutoFilter => ({
  range: { firstRow: range.r0, firstCol: range.c0, lastRow: range.r1, lastCol: range.c1 },
  columns: [],
  sort: null,
});

/** True when the engine still holds a definition in `slot`'s place. */
function engineHoldsFilter(wb: WorkbookHandle, sheet: number, slot: FilterSlot): boolean {
  if (slot.target.get()) return true;
  if (slot.table || !wb.capabilities.autoFilter) return false;
  return typeof wb.getSheetAutoFilterXml === 'function' && wb.getSheetAutoFilterXml(sheet) !== null;
}

/** Drop an engine filter the store no longer has. Clearing it exposes its
 * rows in the engine, so they are re-synced from the store mask, where other
 * filters over the same rows may still hide them. */
function removeFilterTarget(
  wb: WorkbookHandle,
  store: SpreadsheetStore | undefined,
  sheet: number,
  target: FilterTarget,
): boolean {
  const existing = target.get();
  const cleared = target.clear();
  const removed = target.remove();
  importedFiltersFor(wb).delete(target.key);
  if (store && existing && wb.capabilities.hiddenRowsCols) {
    const hidden = store.getState().layout.hiddenRows;
    for (let row = existing.range.firstRow + 1; row <= existing.range.lastRow; row += 1) {
      wb.setRowHidden(sheet, row, hidden.has(row));
    }
  }
  return cleared && removed;
}

/** Write one range through the typed API. Returns null when the criteria need
 * the worksheet's raw-XML fallback; a table has none, so criteria the typed
 * model cannot carry (color) stay with the cell layer over an empty filter. */
function writeTypedFilter(
  wb: WorkbookHandle,
  state: State,
  sheet: number,
  slot: FilterSlot,
  range: Range,
  store?: SpreadsheetStore,
): boolean | null {
  const imported = importedFiltersFor(wb).get(slot.target.key);
  const filter =
    typedAutoFilterFromState(state, range, imported, inputLocaleForWorkbook(wb) ?? 'en-US', wb) ??
    (slot.table ? emptyTypedFilter(range) : null);
  if (!filter) return null;
  const criteriaForRange = state.ui.filterCriteria.filter((entry) => sameRange(entry.range, range));
  const hasColorCriteria = criteriaForRange.some((entry) => entry.color !== undefined);
  const set = slot.target.set(filter);
  const applied = set && slot.target.apply();
  if (applied && store) {
    // Advanced filters and slicers own hiddenRows directly and intentionally
    // leave filterCriteria empty. Color criteria are also evaluated by the
    // cell layer. Apply imported/native columns first, then merge their
    // matches with the store mask before the workbook is saved.
    if (criteriaForRange.length === 0 || hasColorCriteria) {
      syncStoreOwnedAutoFilterVisibility(wb, store, sheet, range, slot.target);
    } else {
      syncNativeAutoFilterVisibility(wb, store, sheet, range, slot.target);
    }
  }
  return set && applied;
}

/** Write the worksheet filter through the raw XML seam; an empty definition
 * removes it. */
function writeXmlFilter(
  wb: WorkbookHandle,
  state: State,
  sheet: number,
  range: Range | null,
  store: SpreadsheetStore | undefined,
  locale: string,
): boolean {
  if (!wb.capabilities.autoFilter) return false;
  const set = wb.setSheetAutoFilterXml(sheet, autoFilterXmlFromState(state, range, locale, wb));
  if (set && store && range) {
    const criteriaForRange = state.ui.filterCriteria.filter((entry) =>
      sameRange(entry.range, range),
    );
    const hasColorCriteria = criteriaForRange.some((entry) => entry.color !== undefined);
    if (criteriaForRange.length === 0 || hasColorCriteria) {
      // Old engines have no typed evaluator; write the cell-layer mask
      // directly so an XML fallback cannot expose rows on save/reload.
      syncStoreOwnedAutoFilterVisibility(wb, store, sheet, range);
    }
  }
  return set;
}

/** Reconcile every AutoFilter on `sheet` (the worksheet's own and each
 * table's) with the store's filter ranges: a slot with a range is written and
 * applied, a slot without one is removed from the engine. `only` limits the
 * writes to that range; removals are reconciled either way. Representable
 * criteria go through the typed API; the raw XML seam remains the fallback
 * for worksheet color filters and engines that predate typed AutoFilter. */
export function syncAutoFilterToEngine(
  wb: WorkbookHandle,
  state: State,
  sheet: number,
  store?: SpreadsheetStore,
  locale = inputLocaleForWorkbook(wb) ?? 'en-US',
  only?: Range,
): boolean {
  const slots = filterSlotsForSheet(wb, state, sheet);
  if (!wb.capabilities.typedAutoFilter) {
    const range = slots.find((slot) => !slot.table)?.range ?? null;
    return writeXmlFilter(wb, state, sheet, range, store, locale);
  }
  let ok = true;
  // Removals first, so a later apply re-hides rows a removed filter exposed.
  for (const slot of slots) {
    if (slot.range || !engineHoldsFilter(wb, sheet, slot)) continue;
    ok = removeFilterTarget(wb, store, sheet, slot.target) && ok;
  }
  for (const slot of slots) {
    const range = slot.range;
    if (!range || (only && !sameRange(only, range))) continue;
    const written =
      writeTypedFilter(wb, state, sheet, slot, range, store) ??
      writeXmlFilter(wb, state, sheet, range, store, locale);
    ok = written && ok;
  }
  return ok;
}

/** Decode an engine-owned typed definition into store filter state. */
function decodeTypedFilter(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
  target: FilterTarget,
  typed: AutoFilter,
): { range: Range; criteria: ValueFilterCriteria[] } {
  const decoded = localCriteriaFromTypedFilter(
    store.getState(),
    sheet,
    typed,
    inputLocaleForWorkbook(wb) ?? 'en-US',
    wb,
  );
  importedFiltersFor(wb).set(target.key, {
    filter: typed,
    unsupportedColumns: decoded.unsupportedColumns,
  });
  return { range: decoded.range, criteria: decoded.criteria };
}

/** Load the worksheet's AutoFilter and every table AutoFilter on `sheet` into
 * `ui.filterRanges` / `ui.filterCriteria`. */
export function hydrateAutoFilterFromEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
): void {
  const ranges: Range[] = [];
  const criteria: ValueFilterCriteria[] = [];
  let worksheetTyped = false;
  if (wb.capabilities.typedAutoFilter) {
    const load = (target: FilterTarget): boolean => {
      const typed = target.get();
      if (!typed) {
        importedFiltersFor(wb).delete(target.key);
        return false;
      }
      const decoded = decodeTypedFilter(wb, store, sheet, target, typed);
      ranges.push(decoded.range);
      criteria.push(...decoded.criteria);
      return true;
    };
    worksheetTyped = load(sheetFilterTarget(wb, sheet));
    if (wb.capabilities.typedTableAutoFilter) {
      for (const table of wb.getTables()) {
        if (table.sheetIndex === sheet) load(tableFilterTarget(wb, table.index));
      }
    }
    // Row visibility is hydrated from the engine's saved row-layout flags by
    // hydrateLayoutFromEngine. Re-evaluating the definitions here would
    // overwrite an intentional saved snapshot, such as a reapply undo or a
    // color/advanced filter whose rows are owned by the cell layer.
  }
  // A typed read can fail for an opaque imported criterion. The raw XML seam
  // keeps the worksheet filter's range visible, with its criteria left
  // engine-owned until a UI mutation replaces them.
  let xml: string | null = null;
  if (
    !worksheetTyped &&
    wb.capabilities.autoFilter &&
    typeof wb.getSheetAutoFilterXml === 'function'
  ) {
    xml = wb.getSheetAutoFilterXml(sheet);
    const range = xml === null ? null : autoFilterRangeFromXml(xml, sheet);
    if (range) ranges.push(range);
  }
  if (!wb.capabilities.typedAutoFilter && xml === null) return;
  wb.withEngineSyncMuted(() => {
    store.setState((state) => ({
      ...state,
      ui: { ...state.ui, filterRanges: ranges, filterCriteria: criteria },
    }));
  });
}
