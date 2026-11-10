import {
  type AutoFilter,
  type AutoFilterColumnInput,
  type AutoFilterEval,
  type AutoFilterInput,
  type AutoFilterSortStateInput,
  ComparisonOperator,
  DynamicFilterType,
  FilterKind,
  SortBy,
  SortMethod,
} from '@libraz/formulon';
import { inputLocaleForWorkbook, parseLocalizedNumericInput } from '../commands/coerce-input.js';
import { distinctValues, normalizeColorKey } from '../commands/filter.js';
import type { HiddenIndices } from '../store/row-visibility.js';
import { sameRange } from '../store/selection-geometry.js';
import type {
  FilterSortSpec,
  SpreadsheetStore,
  State,
  ValueFilterCriteria,
} from '../store/store.js';
import { formatA1Range } from './address.js';
import { dxfToApply } from './cf-sync.js';
import { dxfRecordFromApply } from './cf-writeback.js';
import { profileTextKey } from './profile-text-semantics.js';
import { parseRangeRef } from './range-resolver.js';
import type { Range } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

const escapeXml = (value: string): string =>
  value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

type ConditionOp = NonNullable<ValueFilterCriteria['condition']>['op'];

const conditionOperator: Record<
  Exclude<ConditionOp, 'between' | 'top10' | 'bottom10' | 'aboveAverage' | 'belowAverage'>,
  string
> = {
  equals: 'equal',
  notEquals: 'notEqual',
  contains: 'equal',
  notContains: 'notEqual',
  beginsWith: 'equal',
  notBeginsWith: 'notEqual',
  endsWith: 'equal',
  notEndsWith: 'notEqual',
  greaterThan: 'greaterThan',
  greaterThanOrEqual: 'greaterThanOrEqual',
  lessThan: 'lessThan',
  lessThanOrEqual: 'lessThanOrEqual',
};

const FILTER_KIND_VALUES = FilterKind.Values;
const FILTER_KIND_CUSTOM = FilterKind.Custom;
const FILTER_KIND_TOP10 = FilterKind.Top10;
const FILTER_KIND_DYNAMIC = FilterKind.Dynamic;
const FILTER_KIND_COLOR = FilterKind.Color;

const DYNAMIC_ABOVE_AVERAGE = DynamicFilterType.AboveAverage;
const DYNAMIC_BELOW_AVERAGE = DynamicFilterType.BelowAverage;
const SORT_BY_VALUE = SortBy.Value;
const FILTER_OPERATOR_EQUAL = ComparisonOperator.Equal;
const FILTER_OPERATOR_LESS_THAN = ComparisonOperator.LessThan;
const FILTER_OPERATOR_LESS_THAN_OR_EQUAL = ComparisonOperator.LessThanOrEqual;
const FILTER_OPERATOR_NOT_EQUAL = ComparisonOperator.NotEqual;
const FILTER_OPERATOR_GREATER_THAN_OR_EQUAL = ComparisonOperator.GreaterThanOrEqual;
const FILTER_OPERATOR_GREATER_THAN = ComparisonOperator.GreaterThan;

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
  set(filter: AutoFilterInput): boolean;
  remove(): boolean;
  apply(): boolean;
  clear(): boolean;
  evaluate(): AutoFilterEval | null;
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

const emptyAutoFilterColumn = (columnOffset: number): AutoFilterColumnInput => ({
  columnOffset,
  hiddenButton: false,
  showButton: true,
  kind: FilterKind.None,
  filterBlank: false,
  values: [],
  dateGroups: [],
  customAnd: false,
  conditions: [],
  top: false,
  percent: false,
  topValue: 0,
  filterValue: null,
  dynamicType: DynamicFilterType.Null,
  dynamicValue: null,
  dynamicMaxValue: null,
  valueIso: null,
  maxValueIso: null,
  dxfIndex: null,
  cellColor: false,
  iconSet: 0,
  iconIndex: null,
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

type PatternOp =
  | 'contains'
  | 'notContains'
  | 'beginsWith'
  | 'notBeginsWith'
  | 'endsWith'
  | 'notEndsWith';

const PATTERN_AFFIX: Record<PatternOp, readonly [string, string]> = {
  contains: ['*', '*'],
  notContains: ['*', '*'],
  beginsWith: ['', '*'],
  notBeginsWith: ['', '*'],
  endsWith: ['*', ''],
  notEndsWith: ['*', ''],
};

const isPatternOp = (op: string): op is PatternOp => Object.hasOwn(PATTERN_AFFIX, op);

/** Wildcard text for a contains / begins-with / ends-with operand. */
function wildcardPattern(op: PatternOp, value: string): string {
  const [head, tail] = PATTERN_AFFIX[op];
  return `${head}${escapeAutoFilterWildcardLiteral(value)}${tail}`;
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

/** Decode the single outer-wildcard pattern emitted by this adapter into the
 * positive operator it stands for. A pattern containing another unescaped
 * wildcard is left engine-owned because the cell layer cannot represent it. */
function decodeAffixPattern(
  value: string,
): { op: 'contains' | 'beginsWith' | 'endsWith'; text: string } | null {
  const leading = value.startsWith('*');
  // The last star is the outer wildcard only when it is not escaped: an odd
  // run of preceding tildes escapes it, an even run is complete `~~` pairs.
  let tildes = 0;
  for (let i = value.length - 2; i >= 0 && value[i] === '~'; i -= 1) tildes += 1;
  const trailing = value.length >= 2 && value.endsWith('*') && tildes % 2 === 0;
  if (value.length < 2 || (!leading && !trailing)) return null;
  const op = leading && trailing ? 'contains' : leading ? 'endsWith' : 'beginsWith';
  const text = decodeWildcardLiteral(value.slice(leading ? 1 : 0, trailing ? -1 : undefined));
  return text === null ? null : { op, text };
}

function customColumnFromCriteria(
  colId: number,
  condition: NonNullable<ValueFilterCriteria['condition']>,
): AutoFilterColumnInput {
  const base = emptyAutoFilterColumn(colId);
  if (condition.op === 'top10' || condition.op === 'bottom10') {
    return {
      ...base,
      kind: FILTER_KIND_TOP10,
      top: condition.op === 'top10',
      percent: condition.unit === 'percent',
      topValue: Number(condition.value),
    };
  }
  if (condition.op === 'aboveAverage' || condition.op === 'belowAverage') {
    return {
      ...base,
      kind: FILTER_KIND_DYNAMIC,
      dynamicType: condition.op === 'aboveAverage' ? DYNAMIC_ABOVE_AVERAGE : DYNAMIC_BELOW_AVERAGE,
    };
  }
  if (condition.op === 'between') {
    return {
      ...base,
      kind: FILTER_KIND_CUSTOM,
      customAnd: true,
      conditions: [
        { operator: FILTER_OPERATOR_GREATER_THAN_OR_EQUAL, value: condition.value },
        { operator: FILTER_OPERATOR_LESS_THAN_OR_EQUAL, value: condition.value2 ?? '' },
      ],
    };
  }
  const literalWildcard = condition.op === 'equals' || condition.op === 'notEquals';
  const op =
    condition.op === 'equals' ||
    condition.op === 'contains' ||
    condition.op === 'beginsWith' ||
    condition.op === 'endsWith'
      ? FILTER_OPERATOR_EQUAL
      : condition.op === 'notEquals' ||
          condition.op === 'notContains' ||
          condition.op === 'notBeginsWith' ||
          condition.op === 'notEndsWith'
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
    conditions: [
      {
        operator: op,
        value: isPatternOp(condition.op)
          ? wildcardPattern(condition.op, condition.value)
          : literalWildcard
            ? escapeAutoFilterWildcardLiteral(condition.value)
            : condition.value,
      },
    ],
  };
}

const dxfIdCache = new WeakMap<WorkbookHandle, Map<string, number>>();

/** Color column backed by a differential style holding the filter color; null
 * when the engine cannot author differential styles. */
function colorColumnFromCriteria(
  wb: WorkbookHandle | undefined,
  colId: number,
  color: NonNullable<ValueFilterCriteria['color']>,
): AutoFilterColumnInput | null {
  if (!wb?.capabilities.conditionalFormatDxf) return null;
  const cell = color.kind === 'cellColor';
  const dxf = dxfRecordFromApply(wb, cell ? { fill: color.color } : { color: color.color });
  if (!dxf) return null;
  // addDxf appends, so reuse the id already issued for this color.
  const key = JSON.stringify(dxf);
  let ids = dxfIdCache.get(wb);
  if (!ids) {
    ids = new Map();
    dxfIdCache.set(wb, ids);
  }
  let dxfId = ids.get(key);
  if (dxfId === undefined || !wb.getDxf(dxfId)) {
    dxfId = wb.addDxf(dxf);
    if (dxfId < 0) return null;
    ids.set(key, dxfId);
  }
  return {
    ...emptyAutoFilterColumn(colId),
    kind: FILTER_KIND_COLOR,
    dxfIndex: dxfId,
    cellColor: cell,
  };
}

/** The `<sortState>` for a sort of `spec.range`'s data rows by value. */
function sortStateFromSpec(spec: FilterSortSpec): AutoFilterSortStateInput {
  const { range } = spec;
  return {
    range: { firstRow: range.r0 + 1, firstCol: range.c0, lastRow: range.r1, lastCol: range.c1 },
    columnSort: false,
    caseSensitive: false,
    sortMethod: SortMethod.None,
    conditions: spec.keys.map((key) => ({
      range: {
        firstRow: range.r0 + 1,
        firstCol: key.byCol,
        lastRow: range.r1,
        lastCol: key.byCol,
      },
      descending: key.direction === 'desc',
      sortBy: SORT_BY_VALUE,
      customList: null,
    })),
  };
}

/** Value sort levels of a typed `<sortState>`; color, icon and custom-list
 *  levels have no cell-layer representation and are left out. */
function sortSpecFromTyped(typed: AutoFilter, range: Range): FilterSortSpec | null {
  const sort = typed.sort;
  if (!sort || sort.columnSort) return null;
  const keys = sort.conditions
    .filter(
      (condition) =>
        condition.sortBy === SORT_BY_VALUE &&
        condition.customList === null &&
        condition.range.firstCol >= range.c0 &&
        condition.range.firstCol <= range.c1,
    )
    .map((condition) => ({
      byCol: condition.range.firstCol,
      direction: condition.descending ? ('desc' as const) : ('asc' as const),
    }));
  return keys.length > 0 ? { range, keys } : null;
}

/** Convert the cell-layer's criteria to formulon's typed AutoFilter model.
 * Returns `null` for a color criterion when the engine cannot author its
 * differential style and there is no imported typed snapshot to carry
 * forward. */
export function typedAutoFilterFromState(
  state: State,
  range: Range | null,
  imported?: ImportedTypedFilter,
  locale = 'en-US',
  wb?: WorkbookHandle,
): AutoFilterInput | null {
  if (!range) return null;
  const criteria = state.ui.filterCriteria
    .filter((entry) => sameRange(entry.range, range))
    .sort((a, b) => a.byCol - b.byCol);
  const colorColumns = new Map<number, AutoFilterColumnInput | null>();
  for (const entry of criteria) {
    if (!entry.color) continue;
    const colId = entry.byCol - range.c0;
    colorColumns.set(colId, colorColumnFromCriteria(wb, colId, entry.color));
  }
  // A color criterion the engine cannot author keeps an imported typed
  // definition's columns and omits the new one; without an imported snapshot
  // the raw XML fallback is the only safe path.
  if ([...colorColumns.values()].includes(null) && !imported) return null;

  const columns = new Map<number, AutoFilterColumnInput>();
  // Imported Dynamic/Top10/Icon/Color criteria are kept intact while the UI
  // edits a representable column. A subsequent clear of the whole filter
  // removes them along with the definition.
  if (imported && sameAutoFilterRange(imported.filter, range)) {
    for (const column of imported.filter.columns) {
      if (imported.unsupportedColumns.has(column.columnOffset)) {
        columns.set(column.columnOffset, column);
      }
    }
  }

  for (const entry of criteria) {
    const colId = entry.byCol - range.c0;
    if (colId < 0 || colId > range.c1 - range.c0) continue;
    if (entry.color) {
      const column = colorColumns.get(colId);
      if (column) columns.set(colId, column);
      continue;
    }
    if (entry.condition) {
      columns.set(colId, customColumnFromCriteria(colId, entry.condition));
      continue;
    }
    const hidden = new Set(entry.hiddenValues.map((value) => filterMatchKey(value, locale)));
    columns.set(colId, {
      ...emptyAutoFilterColumn(colId),
      kind: FILTER_KIND_VALUES,
      filterBlank: !hidden.has(''),
      values: distinctValues(state, range, entry.byCol, locale, wb).filter(
        (value) => value !== '' && !hidden.has(filterMatchKey(value, locale)),
      ),
    });
  }

  return {
    range: {
      firstRow: range.r0,
      firstCol: range.c0,
      lastRow: range.r1,
      lastCol: range.c1,
    },
    columns: Array.from(columns.values()).sort((a, b) => a.columnOffset - b.columnOffset),
    sort: sortForRange(state, range, imported),
  };
}

function sortForRange(
  state: State,
  range: Range,
  imported: ImportedTypedFilter | undefined,
): AutoFilterSortStateInput | null {
  const spec = state.ui.filterSorts?.find((entry) => sameRange(entry.range, range));
  if (spec) return sortStateFromSpec(spec);
  // Levels the cell layer cannot show stay engine-owned until a sort replaces them.
  const carried =
    imported && sameAutoFilterRange(imported.filter, range) ? imported.filter.sort : null;
  return carried?.conditions.some((condition) => condition.sortBy !== SORT_BY_VALUE)
    ? carried
    : null;
}

function wildcardCondition(
  op: ComparisonOperator,
  value: string,
  locale = 'en-US',
): NonNullable<ValueFilterCriteria['condition']> | null {
  const decoded = decodeAffixPattern(value);
  if (decoded !== null) {
    if (op === FILTER_OPERATOR_EQUAL) return { op: decoded.op, value: decoded.text };
    if (op === FILTER_OPERATOR_NOT_EQUAL) {
      const negated = {
        contains: 'notContains',
        beginsWith: 'notBeginsWith',
        endsWith: 'notEndsWith',
      } as const;
      return { op: negated[decoded.op], value: decoded.text };
    }
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
    const byCol = range.c0 + column.columnOffset;
    if (byCol < range.c0 || byCol > range.c1) {
      unsupportedColumns.add(column.columnOffset);
      continue;
    }
    // The dropdown affordance flags have no cell-layer representation. Keep
    // their native column intact until that column is explicitly edited.
    if (column.hiddenButton || !column.showButton) {
      unsupportedColumns.add(column.columnOffset);
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
    if (
      column.kind === FILTER_KIND_COLOR &&
      column.dxfIndex !== null &&
      wb?.capabilities.conditionalFormatDxf
    ) {
      const apply = dxfToApply(wb, column.dxfIndex);
      const color = column.cellColor ? apply.fill : apply.color;
      if (color) {
        criteria.push({
          range,
          byCol,
          hiddenValues: [],
          color: {
            kind: column.cellColor ? 'cellColor' : 'fontColor',
            color: normalizeColorKey(color),
          },
        });
        continue;
      }
    }
    if (column.kind === FILTER_KIND_TOP10 && column.topValue > 0) {
      criteria.push({
        range,
        byCol,
        hiddenValues: [],
        condition: {
          op: column.top ? 'top10' : 'bottom10',
          value: String(column.topValue),
          unit: column.percent ? 'percent' : 'items',
        },
      });
      continue;
    }
    if (
      column.kind === FILTER_KIND_DYNAMIC &&
      (column.dynamicType === DYNAMIC_ABOVE_AVERAGE || column.dynamicType === DYNAMIC_BELOW_AVERAGE)
    ) {
      criteria.push({
        range,
        byCol,
        hiddenValues: [],
        condition: {
          op: column.dynamicType === DYNAMIC_ABOVE_AVERAGE ? 'aboveAverage' : 'belowAverage',
          value: '',
        },
      });
      continue;
    }
    const [first, second] = column.conditions;
    if (
      column.kind === FILTER_KIND_CUSTOM &&
      column.conditions.length === 2 &&
      column.customAnd &&
      first?.operator === FILTER_OPERATOR_GREATER_THAN_OR_EQUAL &&
      second?.operator === FILTER_OPERATOR_LESS_THAN_OR_EQUAL
    ) {
      const low = parseStoredNumeric(first.value, locale);
      const high = parseStoredNumeric(second.value, locale);
      if (low !== null && high !== null) {
        criteria.push({
          range,
          byCol,
          hiddenValues: [],
          condition: { op: 'between', value: String(low), value2: String(high) },
        });
        continue;
      }
    }
    if (
      column.kind === FILTER_KIND_CUSTOM &&
      column.conditions.length === 1 &&
      first &&
      !column.customAnd
    ) {
      const condition = wildcardCondition(first.operator, first.value, locale);
      if (condition) {
        criteria.push({ range, byCol, hiddenValues: [], condition });
        continue;
      }
    }
    unsupportedColumns.add(column.columnOffset);
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
      const { op, value: operand } = entry.condition;
      if (op === 'top10' || op === 'bottom10') {
        const percent = entry.condition.unit === 'percent' ? ' percent="1"' : '';
        const top = op === 'bottom10' ? ' top="0"' : '';
        return [
          `<filterColumn colId="${colId}"><top10${top}${percent} val="${escapeXml(operand)}"/></filterColumn>`,
        ];
      }
      if (op === 'aboveAverage' || op === 'belowAverage') {
        return [`<filterColumn colId="${colId}"><dynamicFilter type="${op}"/></filterColumn>`];
      }
      if (op === 'between') {
        return [
          `<filterColumn colId="${colId}"><customFilters and="1"><customFilter operator="greaterThanOrEqual" val="${escapeXml(operand)}"/><customFilter operator="lessThanOrEqual" val="${escapeXml(entry.condition.value2 ?? '')}"/></customFilters></filterColumn>`,
        ];
      }
      const literalWildcard = entry.condition.op === 'equals' || entry.condition.op === 'notEquals';
      const value = isPatternOp(entry.condition.op)
        ? wildcardPattern(entry.condition.op, entry.condition.value)
        : literalWildcard
          ? escapeAutoFilterWildcardLiteral(entry.condition.value)
          : entry.condition.value;
      return [
        `<filterColumn colId="${colId}"><customFilters><customFilter operator="${conditionOperator[op]}" val="${escapeXml(value)}"/></customFilters></filterColumn>`,
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
  hiddenAfter: HiddenIndices = after,
): boolean {
  if (!wb.capabilities.hiddenRowsCols) return false;
  const rows = new Set<number>();
  for (const row of before) rows.add(row);
  for (const row of after) rows.add(row);
  wb.withBatchedRecalc(() => {
    for (const row of rows) wb.setRowHidden(sheet, row, hiddenAfter.has(row));
  });
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

  wb.withBatchedRecalc(() => {
    for (let row = range.r0 + 1; row <= range.r1; row += 1) {
      const index = row - firstRow;
      const nativeHidden = evaluated ? evaluated.match[index] === false : false;
      wb.setRowHidden(sheet, row, nativeHidden || storeHidden.has(row));
    }
  });
  return true;
}

const emptyTypedFilter = (range: Range): AutoFilterInput => ({
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
    wb.withBatchedRecalc(() => {
      for (let row = existing.range.firstRow + 1; row <= existing.range.lastRow; row += 1) {
        wb.setRowHidden(sheet, row, hidden.has(row));
      }
    });
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
): { range: Range; criteria: ValueFilterCriteria[]; sort: FilterSortSpec | null } {
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
  return {
    range: decoded.range,
    criteria: decoded.criteria,
    sort: sortSpecFromTyped(typed, decoded.range),
  };
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
  const sorts: FilterSortSpec[] = [];
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
      if (decoded.sort) sorts.push(decoded.sort);
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
      ui: { ...state.ui, filterRanges: ranges, filterCriteria: criteria, filterSorts: sorts },
    }));
  });
}
