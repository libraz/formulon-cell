import {
  type AutoFilter,
  type AutoFilterColumnInput,
  type AutoFilterConditionInput,
  type AutoFilterEval,
  type AutoFilterInput,
  type AutoFilterSortConditionInput,
  type AutoFilterSortStateInput,
  ColorKind,
  ComparisonOperator,
  DynamicFilterType,
  FilterKind,
  SortBy,
  SortMethod,
} from '@libraz/formulon';
import { inputLocaleForWorkbook, parseLocalizedNumericInput } from '../commands/coerce-input.js';
import {
  datePeriodDynamicType,
  datePeriodOfDynamicType,
  datePeriodXmlType,
} from '../commands/date-periods.js';
import { NO_COLOR, normalizeColorKey } from '../commands/filter-colors.js';
import { isFilterComparisonOp } from '../commands/filter-conditions.js';
import {
  columnDateSerials,
  dateGroupItemsXml,
  dateGroupsForVisible,
  serialInDateGroups,
} from '../commands/filter-date-groups.js';
import { distinctValues } from '../commands/filter-items.js';
import type { HiddenIndices } from '../store/row-visibility.js';
import { sameRange } from '../store/selection-geometry.js';
import type {
  FilterComparisonOp,
  FilterHiddenButtons,
  FilterSortKey,
  FilterSortSpec,
  SpreadsheetStore,
  State,
  ValueFilterCriteria,
} from '../store/store.js';
import { formatA1Range, rangeOfRect, rectOfRange } from './address.js';
import { customIconFromEngine, customIconToEngine } from './cf-icon-sets.js';
import { dxfToApply } from './cf-sync.js';
import { dxfRecordFromApply } from './cf-writeback.js';
import { fillRecordFromFormat, fontRecordFromFormat } from './format-writeback.js';
import { profileTextKey } from './profile-text-semantics.js';
import { parseRangeRef } from './range-resolver.js';
import { noColor } from './style-records.js';
import type { DxfRecord, Range } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

const escapeXml = (value: string): string =>
  value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

type ConditionOp = NonNullable<ValueFilterCriteria['condition']>['op'];

const conditionOperator: Record<
  Exclude<
    ConditionOp,
    'between' | 'top10' | 'bottom10' | 'aboveAverage' | 'belowAverage' | 'datePeriod'
  >,
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

const FILTER_KIND_NONE = FilterKind.None;
const FILTER_KIND_VALUES = FilterKind.Values;
const FILTER_KIND_CUSTOM = FilterKind.Custom;
const FILTER_KIND_TOP10 = FilterKind.Top10;
const FILTER_KIND_DYNAMIC = FilterKind.Dynamic;
const FILTER_KIND_COLOR = FilterKind.Color;
const FILTER_KIND_ICON = FilterKind.Icon;

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
  evaluate: () => wb.evaluateAutoFilter(sheet),
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

const rangeFromAutoFilter = (filter: AutoFilter, sheet: number): Range =>
  rangeOfRect(sheet, filter.range);

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

/** The kind of the file criterion on `col` of the filter over `range` that the cell UI cannot
 *  show (it stays in the engine's definition), or null when the column holds none. */
export function unshownFilterKind(
  wb: WorkbookHandle,
  state: State,
  range: Range,
  col: number,
): FilterKind | null {
  const filters = importedTypedFilters.get(wb);
  if (!filters?.size) return null;
  const slot = filterSlotsForSheet(wb, state, range.sheet).find(
    (candidate) => candidate.range && sameRange(candidate.range, range),
  );
  const imported = slot ? filters.get(slot.target.key) : undefined;
  const offset = col - range.c0;
  if (!imported?.unsupportedColumns.has(offset) || !sameAutoFilterRange(imported.filter, range)) {
    return null;
  }
  return imported.filter.columns.find((column) => column.columnOffset === offset)?.kind ?? null;
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

/** A checklist column's kept items, with the column's dated numbers, which the file stores as
 *  date-group items rather than as values. */
function checklistVisible(
  state: State,
  range: Range,
  entry: ValueFilterCriteria,
  locale: string,
  wb?: WorkbookHandle,
): { visible: string[]; dates: Map<string, number>; filterBlank: boolean } {
  const hidden = new Set(entry.hiddenValues.map((value) => filterMatchKey(value, locale)));
  const visible = distinctValues(state, range, entry.byCol, locale, wb).filter(
    (value) => !hidden.has(filterMatchKey(value, locale)),
  );
  return {
    visible,
    dates: columnDateSerials(state, range, entry.byCol),
    filterBlank: !hidden.has(''),
  };
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
  if (condition.op === 'datePeriod') {
    return {
      ...base,
      kind: FILTER_KIND_DYNAMIC,
      dynamicType: condition.period
        ? datePeriodDynamicType(condition.period)
        : DynamicFilterType.Null,
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
  const first = customConditionInput(condition.op as FilterComparisonOp, condition.value);
  return {
    ...base,
    kind: FILTER_KIND_CUSTOM,
    ...(condition.second
      ? {
          customAnd: condition.join !== 'or',
          conditions: [first, customConditionInput(condition.second.op, condition.second.value)],
        }
      : { conditions: [first] }),
  };
}

/** One comparison as a native custom condition. */
function customConditionInput(op: FilterComparisonOp, value: string): AutoFilterConditionInput {
  const operator =
    op === 'equals' || op === 'contains' || op === 'beginsWith' || op === 'endsWith'
      ? FILTER_OPERATOR_EQUAL
      : op === 'notEquals' || op === 'notContains' || op === 'notBeginsWith' || op === 'notEndsWith'
        ? FILTER_OPERATOR_NOT_EQUAL
        : op === 'greaterThan'
          ? FILTER_OPERATOR_GREATER_THAN
          : op === 'greaterThanOrEqual'
            ? FILTER_OPERATOR_GREATER_THAN_OR_EQUAL
            : op === 'lessThan'
              ? FILTER_OPERATOR_LESS_THAN
              : FILTER_OPERATOR_LESS_THAN_OR_EQUAL;
  return { operator, value: encodeComparisonValue(op, value) };
}

/** A comparison's operand as a custom condition stores it: text patterns become wildcards, and a
 *  literal wildcard character in an equality is escaped. */
const encodeComparisonValue = (op: FilterComparisonOp, value: string): string =>
  isPatternOp(op)
    ? wildcardPattern(op, value)
    : op === 'equals' || op === 'notEquals'
      ? escapeAutoFilterWildcardLiteral(value)
      : value;

const AUTOMATIC = { ...noColor(), kind: ColorKind.Auto };

/** The differential format of No Fill (a fill with no pattern) or of an Automatic font colour. */
const noColorDxf = (kind: 'cellColor' | 'fontColor'): DxfRecord =>
  kind === 'cellColor'
    ? { fill: fillRecordFromFormat({}) }
    : { font: { ...fontRecordFromFormat({}), sizePt: null, vertAlign: 0, color: AUTOMATIC } };

/** The colour key a colour filter or sort level's differential format holds: its fill or font
 *  colour, `NO_COLOR` for No Fill or Automatic, null for one that states neither. */
function colorKeyOfDxf(wb: WorkbookHandle, dxfIndex: number, cell: boolean): string | null {
  const dxf = wb.getDxf(dxfIndex);
  if (cell && dxf?.fill && dxf.fill.pattern === 0 && !dxf.fill.gradient) return NO_COLOR;
  if (!cell && dxf?.font?.color?.kind === ColorKind.Auto) return NO_COLOR;
  const apply = dxfToApply(wb, dxfIndex);
  const color = cell ? apply.fill : apply.color;
  return color ? normalizeColorKey(color) : null;
}

const dxfIdCache = new WeakMap<WorkbookHandle, Map<string, number>>();

/** Color column backed by a differential style holding the filter color; null
 * when the engine cannot author differential styles. */
/** The differential format whose fill (`cellColor`) or font colour carries `color`, issued once
 *  per workbook; null when the engine cannot author one. */
function dxfIdForColor(
  wb: WorkbookHandle | undefined,
  kind: 'cellColor' | 'fontColor',
  color: string,
): number | null {
  if (!wb?.capabilities.conditionalFormatDxf) return null;
  const dxf =
    color === NO_COLOR
      ? noColorDxf(kind)
      : dxfRecordFromApply(wb, kind === 'cellColor' ? { fill: color } : { color });
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
  return dxfId;
}

function colorColumnFromCriteria(
  wb: WorkbookHandle | undefined,
  colId: number,
  color: NonNullable<ValueFilterCriteria['color']>,
): AutoFilterColumnInput | null {
  if (color.kind === 'cellIcon') {
    const icon = customIconToEngine(color.icon);
    return icon
      ? {
          ...emptyAutoFilterColumn(colId),
          kind: FILTER_KIND_ICON,
          iconSet: icon.iconSet,
          iconIndex: icon.iconIndex,
        }
      : null;
  }
  const cell = color.kind === 'cellColor';
  const dxfId = dxfIdForColor(wb, color.kind, color.color);
  if (dxfId === null) return null;
  return {
    ...emptyAutoFilterColumn(colId),
    kind: FILTER_KIND_COLOR,
    dxfIndex: dxfId,
    cellColor: cell,
  };
}

/** The `<sortState>` for a sort of `spec.range`'s data rows by value. */
function sortStateFromSpec(spec: FilterSortSpec, wb?: WorkbookHandle): AutoFilterSortStateInput {
  const { range } = spec;
  return {
    range: { firstRow: range.r0 + 1, firstCol: range.c0, lastRow: range.r1, lastCol: range.c1 },
    columnSort: false,
    caseSensitive: false,
    sortMethod: SortMethod.None,
    conditions: spec.keys.flatMap((key): AutoFilterSortConditionInput[] => {
      const base = {
        range: {
          firstRow: range.r0 + 1,
          firstCol: key.byCol,
          lastRow: range.r1,
          lastCol: key.byCol,
        },
        descending: key.direction === 'desc',
        customList: null,
      };
      if ((key.sortOn === 'cellColor' || key.sortOn === 'fontColor') && key.color) {
        const dxfIndex = dxfIdForColor(wb, key.sortOn, key.color);
        if (dxfIndex === null) return [];
        return [
          {
            ...base,
            sortBy: key.sortOn === 'cellColor' ? SortBy.CellColor : SortBy.FontColor,
            dxfIndex,
          },
        ];
      }
      if (key.sortOn === 'customList' && key.customList?.length) {
        return [{ ...base, sortBy: SORT_BY_VALUE, customList: key.customList.join(',') }];
      }
      if (key.sortOn === 'cellIcon' && key.icon) {
        const icon = customIconToEngine(key.icon);
        return icon
          ? [{ ...base, sortBy: SortBy.Icon, iconSet: icon.iconSet, iconIndex: icon.iconIndex }]
          : [];
      }
      return [{ ...base, sortBy: SORT_BY_VALUE }];
    }),
  };
}

/** The value, custom-list, colour and icon sort levels of a typed `<sortState>`. */
function sortSpecFromTyped(
  typed: AutoFilter,
  range: Range,
  wb?: WorkbookHandle,
): FilterSortSpec | null {
  const sort = typed.sort;
  if (!sort || sort.columnSort) return null;
  const keys = sort.conditions.flatMap((condition): FilterSortKey[] => {
    const byCol = condition.range.firstCol;
    if (byCol < range.c0 || byCol > range.c1) return [];
    const base = { byCol, direction: condition.descending ? ('desc' as const) : ('asc' as const) };
    if (condition.sortBy === SORT_BY_VALUE) {
      return condition.customList
        ? [{ ...base, sortOn: 'customList', customList: condition.customList.split(',') }]
        : [base];
    }
    if (condition.sortBy === SortBy.Icon) {
      const icon =
        condition.iconIndex === null
          ? null
          : customIconFromEngine({ iconSet: condition.iconSet, iconIndex: condition.iconIndex });
      return icon ? [{ ...base, sortOn: 'cellIcon', icon }] : [];
    }
    if (condition.dxfIndex === null || !wb?.capabilities.conditionalFormatDxf) return [];
    const cell = condition.sortBy === SortBy.CellColor;
    const color = colorKeyOfDxf(wb, condition.dxfIndex, cell);
    return color ? [{ ...base, sortOn: cell ? 'cellColor' : 'fontColor', color }] : [];
  });
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
    const { visible, dates, filterBlank } = checklistVisible(state, range, entry, locale, wb);
    columns.set(colId, {
      ...emptyAutoFilterColumn(colId),
      kind: FILTER_KIND_VALUES,
      filterBlank,
      values: visible.filter((value) => value !== '' && !dates.has(value)),
      dateGroups: dateGroupsForVisible(
        visible.flatMap((value) => {
          const serial = dates.get(value);
          return serial === undefined ? [] : [serial];
        }),
        dates.values(),
        state.ui.date1904 === true,
      ),
    });
  }
  for (const col of hiddenButtonCols(state, range)) {
    const colId = col - range.c0;
    columns.set(colId, {
      ...(columns.get(colId) ?? emptyAutoFilterColumn(colId)),
      hiddenButton: true,
      showButton: false,
    });
  }

  return {
    range: rectOfRange(range),
    columns: Array.from(columns.values()).sort((a, b) => a.columnOffset - b.columnOffset),
    sort: sortForRange(state, range, wb),
  };
}

/** The sheet columns of `range` whose dropdown button is hidden. */
const hiddenButtonCols = (state: State, range: Range): number[] =>
  (state.ui.filterHiddenButtons ?? [])
    .filter((spec) => sameRange(spec.range, range))
    .flatMap((spec) => spec.cols)
    .filter((col) => col >= range.c0 && col <= range.c1);

function sortForRange(
  state: State,
  range: Range,
  wb?: WorkbookHandle,
): AutoFilterSortStateInput | null {
  const spec = state.ui.filterSorts?.find((entry) => sameRange(entry.range, range));
  return spec ? sortStateFromSpec(spec, wb) : null;
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
): {
  range: Range;
  criteria: ValueFilterCriteria[];
  unsupportedColumns: Set<number>;
  hiddenButtons: number[];
} {
  const range = rangeFromAutoFilter(filter, sheet);
  const criteria: ValueFilterCriteria[] = [];
  const unsupportedColumns = new Set<number>();
  // A column whose button the file hides keeps its filter; only the button goes.
  const hiddenButtons = filter.columns
    .filter((column) => column.hiddenButton || !column.showButton)
    .map((column) => range.c0 + column.columnOffset);
  for (const column of filter.columns) {
    const byCol = range.c0 + column.columnOffset;
    if (column.kind === FILTER_KIND_NONE) continue;
    if (column.kind === FILTER_KIND_VALUES) {
      const visible = new Set(column.values.map((value) => filterMatchKey(value, locale)));
      const dates =
        column.dateGroups.length > 0 ? columnDateSerials(state, range, byCol) : new Map();
      const date1904 = state.ui.date1904 === true;
      const hiddenValues = distinctValues(state, range, byCol, locale, wb).filter((value) => {
        if (value === '') return !column.filterBlank;
        const serial = dates.get(value);
        if (serial !== undefined) return !serialInDateGroups(serial, column.dateGroups, date1904);
        return !visible.has(filterMatchKey(value, locale));
      });
      criteria.push({ range, byCol, hiddenValues });
      continue;
    }
    if (
      column.kind === FILTER_KIND_COLOR &&
      column.dxfIndex !== null &&
      wb?.capabilities.conditionalFormatDxf
    ) {
      const color = colorKeyOfDxf(wb, column.dxfIndex, column.cellColor);
      if (color) {
        criteria.push({
          range,
          byCol,
          hiddenValues: [],
          color: { kind: column.cellColor ? 'cellColor' : 'fontColor', color },
        });
        continue;
      }
    }
    const icon =
      column.kind === FILTER_KIND_ICON && column.iconIndex !== null
        ? customIconFromEngine({ iconSet: column.iconSet, iconIndex: column.iconIndex })
        : null;
    if (icon) {
      criteria.push({ range, byCol, hiddenValues: [], color: { kind: 'cellIcon', icon } });
      continue;
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
    const period =
      column.kind === FILTER_KIND_DYNAMIC ? datePeriodOfDynamicType(column.dynamicType) : null;
    if (period) {
      criteria.push({
        range,
        byCol,
        hiddenValues: [],
        condition: { op: 'datePeriod', value: '', period },
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
    if (column.kind === FILTER_KIND_CUSTOM && column.conditions.length === 2 && first && second) {
      const one = wildcardCondition(first.operator, first.value, locale);
      const two = wildcardCondition(second.operator, second.value, locale);
      if (one && two && isFilterComparisonOp(one.op) && isFilterComparisonOp(two.op)) {
        criteria.push({
          range,
          byCol,
          hiddenValues: [],
          condition: {
            op: one.op,
            value: one.value,
            second: { op: two.op, value: two.value },
            join: column.customAnd ? 'and' : 'or',
          },
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
  return { range, criteria, unsupportedColumns, hiddenButtons };
}

/** One comparison as a `<customFilter>`. */
const customFilterXml = (op: FilterComparisonOp, value: string): string =>
  `<customFilter operator="${conditionOperator[op]}" val="${escapeXml(encodeComparisonValue(op, value))}"/>`;

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
  if (criteria.length === 0 && hiddenButtonCols(state, range).length === 0)
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
      if (op === 'datePeriod') {
        const type = entry.condition.period ? datePeriodXmlType(entry.condition.period) : '';
        return type
          ? [`<filterColumn colId="${colId}"><dynamicFilter type="${type}"/></filterColumn>`]
          : [];
      }
      if (op === 'between') {
        return [
          `<filterColumn colId="${colId}"><customFilters and="1"><customFilter operator="greaterThanOrEqual" val="${escapeXml(operand)}"/><customFilter operator="lessThanOrEqual" val="${escapeXml(entry.condition.value2 ?? '')}"/></customFilters></filterColumn>`,
        ];
      }
      const { second } = entry.condition;
      const filters = [
        customFilterXml(op as FilterComparisonOp, operand),
        ...(second ? [customFilterXml(second.op, second.value)] : []),
      ].join('');
      const and = second && entry.condition.join !== 'or' ? ' and="1"' : '';
      return [
        `<filterColumn colId="${colId}"><customFilters${and}>${filters}</customFilters></filterColumn>`,
      ];
    }
    const hidden = new Set(entry.hiddenValues.map((value) => filterMatchKey(value, locale)));
    const visible = distinctValues(state, range, entry.byCol, locale, wb).filter(
      (value) => !hidden.has(filterMatchKey(value, locale)),
    );
    const includeBlank = visible.includes('');
    const dates = columnDateSerials(state, range, entry.byCol);
    const filters =
      visible
        .filter((value) => value !== '' && !dates.has(value))
        .map((value) => `<filter val="${escapeXml(value)}"/>`)
        .join('') +
      dateGroupItemsXml(
        dateGroupsForVisible(
          visible.flatMap((value) => {
            const serial = dates.get(value);
            return serial === undefined ? [] : [serial];
          }),
          dates.values(),
          state.ui.date1904 === true,
        ),
      );
    const blankAttr = includeBlank ? ' blank="1"' : '';
    return [
      `<filterColumn colId="${colId}"><filters${blankAttr}>${filters}</filters></filterColumn>`,
    ];
  });
  return `<autoFilter ref="${formatA1Range(range, { collapse: false })}">${withHiddenButtons(
    columns,
    hiddenButtonCols(state, range).map((col) => col - range.c0),
  ).join('')}</autoFilter>`;
}

/** `<filterColumn>` elements with `hiddenButton="1"` on each of `hidden`, adding an empty
 *  element for a hidden column without a criterion, in column order. */
function withHiddenButtons(columns: readonly string[], hidden: readonly number[]): string[] {
  const byId = new Map<number, string>();
  for (const column of columns) {
    const id = Number(/^<filterColumn colId="(\d+)"/.exec(column)?.[1]);
    byId.set(id, column);
  }
  for (const id of hidden) {
    const column = byId.get(id);
    byId.set(
      id,
      column
        ? column.replace(
            `<filterColumn colId="${id}"`,
            `<filterColumn colId="${id}" hiddenButton="1"`,
          )
        : `<filterColumn colId="${id}" hiddenButton="1"/>`,
    );
  }
  return [...byId].sort(([a], [b]) => a - b).map(([, column]) => column);
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
  range: rectOfRange(range),
  columns: [],
  sort: null,
});

/** True when the engine still holds a definition in `slot`'s place. */
function engineHoldsFilter(wb: WorkbookHandle, sheet: number, slot: FilterSlot): boolean {
  if (slot.target.get()) return true;
  if (slot.table || !wb.capabilities.autoFilter) return false;
  return wb.getSheetAutoFilterXml(sheet) !== null;
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
): {
  range: Range;
  criteria: ValueFilterCriteria[];
  sort: FilterSortSpec | null;
  hiddenButtons: FilterHiddenButtons | null;
} {
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
    sort: sortSpecFromTyped(typed, decoded.range, wb),
    hiddenButtons:
      decoded.hiddenButtons.length > 0
        ? { range: decoded.range, cols: decoded.hiddenButtons }
        : null,
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
  const hiddenButtons: FilterHiddenButtons[] = [];
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
      if (decoded.hiddenButtons) hiddenButtons.push(decoded.hiddenButtons);
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
  if (!worksheetTyped && wb.capabilities.autoFilter) {
    xml = wb.getSheetAutoFilterXml(sheet);
    const range = xml === null ? null : autoFilterRangeFromXml(xml, sheet);
    if (range) ranges.push(range);
  }
  if (!wb.capabilities.typedAutoFilter && xml === null) return;
  wb.withEngineSyncMuted(() => {
    store.setState((state) => ({
      ...state,
      ui: {
        ...state.ui,
        filterRanges: ranges,
        filterCriteria: criteria,
        filterSorts: sorts,
        filterHiddenButtons: hiddenButtons,
      },
    }));
  });
}
