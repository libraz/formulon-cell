import {
  inputLocaleForWorkbook,
  parseLocalizedNumericInput,
  workbookDate1904,
} from '../commands/coerce-input.js';
import { parseDateTimeInput } from '../commands/input-grammar.js';
import type { PivotSourceField } from '../commands/pivot-table.js';
import {
  PIVOT_SHOW_AS_BASE_NEXT,
  PIVOT_SHOW_AS_BASE_PREVIOUS,
  type PivotAggregation,
  PivotAxis,
  PivotFilterKind,
  type PivotFilterSpec,
  PivotShowValuesAs,
  type PivotShowValuesAs as PivotShowValuesAsValue,
} from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { isMacHost } from '../extensions/ui-options.js';
import { excelSerialToIsoDate } from '../format/date-serial.js';
import type { DatePeriodKey } from '../i18n/strings/data-objects/types.js';
import type { Strings } from '../i18n/strings.js';
import { appendDialogSelectOptions, createDialogSelect } from '../toolbar/dialogs/form-controls.js';
import { createDialogButton } from './dialog-shell.js';
import { showPivotFilterDialog } from './pivot-filter-dialog.js';
import {
  PIVOT_FILTER_COMPARISONS,
  PIVOT_FILTER_PERIODS,
  type PivotFilterCategory,
  type PivotTop10Basis,
  type PivotTop10Direction,
  pivotFilterCategory,
  pivotFilterOperand,
  pivotTop10Kind,
  pivotTop10Parts,
} from './pivot-filter-kinds.js';

export type PivotAreaKind = 'filters' | 'columns' | 'rows' | 'values';

export interface PivotFieldSettingsActive {
  kind: PivotAreaKind;
  fieldName: string;
}

export type PivotValueFieldSetting = {
  aggregation?: PivotAggregation;
  numberFormat?: string;
  showValuesAs?: PivotShowValuesAsValue;
  /** Row or column field name the show-as kind is measured along. */
  showAsBaseField?: string;
  /** Position in the base field's items, or a base-item sentinel. */
  showAsBaseItem?: number;
};

const BASE_ITEM_KINDS: ReadonlySet<PivotShowValuesAsValue> = new Set([
  PivotShowValuesAs.DifferenceFrom,
  PivotShowValuesAs.PercentDifferenceFrom,
  PivotShowValuesAs.PercentOf,
]);

const BASE_FIELD_KINDS: ReadonlySet<PivotShowValuesAsValue> = new Set([
  ...BASE_ITEM_KINDS,
  PivotShowValuesAs.PercentOfRunningTotal,
  PivotShowValuesAs.PercentOfParent,
  PivotShowValuesAs.RankAscending,
  PivotShowValuesAs.RankDescending,
]);

/** Whether the engine reads a base field for this show-values-as kind. */
export const pivotShowAsNeedsBaseField = (kind: PivotShowValuesAsValue | undefined): boolean =>
  kind !== undefined && BASE_FIELD_KINDS.has(kind);

/** Whether the engine also reads a base item for this show-values-as kind. */
export const pivotShowAsNeedsBaseItem = (kind: PivotShowValuesAsValue | undefined): boolean =>
  kind !== undefined && BASE_ITEM_KINDS.has(kind);

/**
 * Completes or clears the base field/item of a value-field setting so it matches what
 * its kind needs: a missing or stale base defaults to the first eligible field and to
 * "(previous)". `itemCount` (optional) invalidates a specific item past the field's items.
 */
export const resolvePivotShowAsBase = (
  setting: PivotValueFieldSetting,
  eligibleFields: readonly string[],
  itemCount?: (fieldName: string) => number,
): PivotValueFieldSetting => {
  const { showAsBaseField: _field, showAsBaseItem: _item, ...rest } = setting;
  if (!pivotShowAsNeedsBaseField(setting.showValuesAs)) return rest;
  const field =
    setting.showAsBaseField !== undefined && eligibleFields.includes(setting.showAsBaseField)
      ? setting.showAsBaseField
      : eligibleFields[0];
  if (field === undefined) return rest;
  if (!pivotShowAsNeedsBaseItem(setting.showValuesAs)) return { ...rest, showAsBaseField: field };
  const item = setting.showAsBaseItem;
  const keep =
    field === setting.showAsBaseField &&
    item !== undefined &&
    (item === PIVOT_SHOW_AS_BASE_PREVIOUS ||
      item === PIVOT_SHOW_AS_BASE_NEXT ||
      (itemCount === undefined ? item >= 0 : item >= 0 && item < itemCount(field)));
  return {
    ...rest,
    showAsBaseField: field,
    showAsBaseItem: keep ? item : PIVOT_SHOW_AS_BASE_PREVIOUS,
  };
};

export const createPivotAreaSettingsButton = (
  label: string,
  ariaLabel = label,
): HTMLButtonElement => {
  const button = createDialogButton({ label, baseClass: 'fc-pivotdlg__area-settings' });
  button.setAttribute('aria-label', ariaLabel);
  return button;
};

/** The filter a settings or filter dialog edits: an engine filter kind, null for none, and its
 *  operand as typed; a range kind keeps both bounds as `low..high`. */
export interface PivotFilterConditionState {
  kind: PivotFilterKind | null;
  value: string;
  /** The data field a value or Top 10 filter scores; absent for the first. */
  dataField?: number;
}

/** How typed operands read: the workbook's input locale and date system. */
export interface PivotFilterInputContext {
  locale?: string;
  date1904?: boolean;
}

/** True when `condition` sets no filter: no kind, or a kind that takes an operand with none typed. */
export const isEmptyPivotFilterCondition = (condition: PivotFilterConditionState): boolean =>
  condition.kind === null ||
  (pivotFilterOperand(condition.kind) !== 'none' && condition.value.trim() === '');

/** The operand reading of `wb`. */
export const pivotFilterInputFor = (wb: WorkbookHandle): PivotFilterInputContext => ({
  locale: inputLocaleForWorkbook(wb),
  date1904: workbookDate1904(wb),
});

interface PivotFieldSettingsControls {
  rowSelect: HTMLSelectElement;
  colSelect: HTMLSelectElement;
  rowSortSelect: HTMLSelectElement;
  colSortSelect: HTMLSelectElement;
  aggSelect: HTMLSelectElement;
  numberFormatInput: HTMLInputElement;
  rowSubtotalTop: HTMLInputElement;
  colSubtotalTop: HTMLInputElement;
}

export interface PivotFieldSettingsPanelOptions {
  host?: HTMLElement;
  panelEl: HTMLDivElement;
  active: PivotFieldSettingsActive | null;
  strings: Strings['pivotTableDialog'];
  /** Relative date period names for the Date Filters list. */
  datePeriods: Record<DatePeriodKey, string>;
  okLabel?: string;
  cancelLabel?: string;
  fields: readonly PivotSourceField[];
  controls: PivotFieldSettingsControls;
  selectedValueFields: readonly string[];
  selectedFilterFields: readonly string[];
  selectedValueSetting(fieldName: string): PivotValueFieldSetting;
  setValueFieldSetting(fieldName: string, setting: PivotValueFieldSetting): void;
  fieldCanBeValue(fieldName: string): boolean;
  replaceFilterField(previous: string, next: string): void;
  normalizeSelectedFilters(): void;
  refreshFieldList(): void;
  filterItemVisibility(fieldName: string, itemName: string): boolean;
  setFilterItemVisibility(fieldName: string, itemName: string, visible: boolean): void;
  selectedFilterCondition(fieldName: string): PivotFilterConditionState | undefined;
  setFilterCondition(fieldName: string, condition: PivotFilterConditionState): void;
  inferFilterItems(fieldName: string): readonly string[];
}

const cloneSelectOptions = (source: HTMLSelectElement, target: HTMLSelectElement): void => {
  target.replaceChildren();
  appendDialogSelectOptions(
    target,
    Array.from(source.options).map((option) => ({
      value: option.value,
      label: option.textContent ?? '',
    })),
  );
  target.value = source.value;
};

export const fillPivotFieldSelect = (
  select: HTMLSelectElement,
  fields: readonly PivotSourceField[],
): void => {
  select.replaceChildren();
  appendDialogSelectOptions(
    select,
    fields.map((field) => ({ value: field.name, label: field.name })),
  );
};

const settingsFieldRow = (label: string, control: HTMLElement): HTMLLabelElement => {
  const row = document.createElement('label');
  row.className = 'fc-pivotdlg__settings-field';
  const text = document.createElement('span');
  text.textContent = label;
  row.append(text, control);
  return row;
};

const settingsCheckRow = (label: string, control: HTMLInputElement): HTMLLabelElement => {
  const row = document.createElement('label');
  row.className = 'fc-pivotdlg__settings-check';
  row.append(control, document.createTextNode(label));
  return row;
};

const FILTER_NUMBER_SOURCE = String.raw`[+-]?(?:(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?)`;
const NUMBER_PAIR_PATTERN = new RegExp(
  String.raw`^\s*(${FILTER_NUMBER_SOURCE})\s*(?:\.\.|,|–|-)\s*(${FILTER_NUMBER_SOURCE})\s*$`,
  'i',
);

const helpFor = (kind: PivotAreaKind, strings: Strings['pivotTableDialog']): string => {
  if (kind === 'filters') return strings.fieldSettingsFiltersHelp;
  if (kind === 'columns') return strings.fieldSettingsColumnsHelp;
  if (kind === 'rows') return strings.fieldSettingsRowsHelp;
  return strings.fieldSettingsValuesHelp;
};

export const splitFilterConditionRange = (value: string): [string, string] => {
  const match = NUMBER_PAIR_PATTERN.exec(value);
  if (match?.[1] !== undefined && match[2] !== undefined) return [match[1], match[2]];
  const parts = value.split(/\s*(?:\.\.|,|–|-)\s*/);
  return [parts[0]?.trim() ?? '', parts[1]?.trim() ?? ''];
};

/** Fills `select` with a category's filter list: a none entry, the desktop's comparisons in
 *  order, then Top 10 (value) or a separator and the relative periods (date). */
export const appendFilterConditionOptions = (
  select: HTMLSelectElement,
  category: PivotFilterCategory,
  strings: Strings['pivotTableDialog'],
  datePeriods: Record<DatePeriodKey, string>,
): void => {
  const dialogs = strings.filterDialogs;
  select.replaceChildren();
  appendDialogSelectOptions(select, [
    { value: '', label: strings.filterConditionNone },
    ...PIVOT_FILTER_COMPARISONS[category].map(([kind, key]) => ({
      value: String(kind),
      label: dialogs.operators[key],
    })),
  ]);
  if (category === 'value') {
    appendDialogSelectOptions(select, [
      { value: String(PivotFilterKind.ValueTopCount), label: dialogs.menu.top10 },
    ]);
  } else if (category === 'date') {
    appendDialogSelectOptions(select, [
      { value: '-', label: '──────────', disabled: true },
      ...PIVOT_FILTER_PERIODS.map(([kind, key]) => ({
        value: String(kind),
        label: datePeriods[key],
      })),
    ]);
  }
};

const splitFilterDateRange = (value: string): [string, string] => {
  const parts = value.split(/\s*(?:\.\.|,|–)\s*/);
  return [parts[0]?.trim() ?? '', parts[1]?.trim() ?? ''];
};

/** One typed operand in the engine's form; null when it does not read as `operand` asks. */
const readOperand = (
  text: string,
  operand: 'text' | 'number' | 'date',
  input: PivotFilterInputContext,
): number | string | null => {
  if (operand === 'text') return text === '' ? null : text;
  if (operand === 'number') return parseLocalizedNumericInput(text, input.locale);
  return parseDateTimeInput(text, input.locale, input.date1904 === true)?.value ?? null;
};

export const pivotFilterConditionToSpec = (
  fieldName: string,
  condition: PivotFilterConditionState | undefined,
  axis: PivotAxis = PivotAxis.Page,
  input: PivotFilterInputContext = {},
): PivotFilterSpec | null => {
  if (!condition || condition.kind === null) return null;
  const { kind } = condition;
  const scored =
    condition.dataField !== undefined && condition.dataField > 0
      ? { dataFieldIndex: condition.dataField }
      : {};
  const text = condition.value.trim();
  const operand = pivotFilterOperand(kind);
  if (operand === 'none') return { axis, fieldName, kind, ...scored };
  if (operand === 'count') {
    const count = Number(text);
    return {
      axis,
      fieldName,
      kind,
      value: Number.isFinite(count) && count > 0 ? Math.floor(count) : 10,
      ...scored,
    };
  }
  if (operand === 'textRange' || operand === 'numberRange' || operand === 'dateRange') {
    const single = operand === 'textRange' ? 'text' : operand === 'numberRange' ? 'number' : 'date';
    const [lowText, highText] =
      single === 'number' ? splitFilterConditionRange(text) : splitFilterDateRange(text);
    const value = readOperand(lowText, single, input);
    const valueHigh = readOperand(highText, single, input);
    if (value === null || valueHigh === null) return null;
    return { axis, fieldName, kind, value, valueHigh, ...scored };
  }
  const value = readOperand(text, operand, input);
  return value === null ? null : { axis, fieldName, kind, value, ...scored };
};

/** One engine operand as the text a dialog shows. */
const operandText = (
  value: number | string | null | undefined,
  date: boolean,
  date1904: boolean,
) =>
  value === null || value === undefined
    ? ''
    : typeof value === 'number' && date
      ? (excelSerialToIsoDate(value, date1904) ?? String(value))
      : String(value);

export const pivotFilterSpecToCondition = (
  spec: PivotFilterSpec | undefined,
  input: PivotFilterInputContext = {},
): PivotFilterConditionState | null => {
  if (!spec) return null;
  const operand = pivotFilterOperand(spec.kind);
  const date = operand === 'date' || operand === 'dateRange';
  const date1904 = input.date1904 === true;
  const low = operandText(spec.value, date, date1904);
  const scored =
    spec.dataFieldIndex !== undefined && spec.dataFieldIndex > 0
      ? { dataField: spec.dataFieldIndex }
      : {};
  if (operand === 'textRange' || operand === 'numberRange' || operand === 'dateRange') {
    return {
      kind: spec.kind,
      value: `${low}..${operandText(spec.valueHigh, date, date1904)}`,
      ...scored,
    };
  }
  return { kind: spec.kind, value: low, ...scored };
};

export interface PivotFilterConditionControlsOptions {
  strings: Strings['pivotTableDialog'];
  /** Relative date period names for the Date Filters list. */
  datePeriods: Record<DatePeriodKey, string>;
  condition?: PivotFilterConditionState;
  /** Data field captions a value or Top 10 filter can score; the index is the data field index. */
  dataFields?: readonly string[];
  selectClassName: string;
  valueClassName: string;
  valuesContainerClassName: string;
  categoryDataset?: Record<string, string>;
  conditionDataset?: Record<string, string>;
  fieldRow(label: string, control: HTMLElement): HTMLElement;
  onChange(condition: PivotFilterConditionState): void;
  onUserChange?(): void;
}

const applyDataset = (element: HTMLElement, dataset?: Record<string, string>): void => {
  for (const [key, value] of Object.entries(dataset ?? {})) element.dataset[key] = value;
};

const TOP10_OPTION_VALUE = String(PivotFilterKind.ValueTopCount);

export const createPivotFilterConditionControls = (
  options: PivotFilterConditionControlsOptions,
): HTMLElement[] => {
  const { strings: t, datePeriods } = options;
  const dialogs = t.filterDialogs;
  const dataFields = options.dataFields ?? [];
  const condition = options.condition ?? { kind: null, value: '' };
  let currentValue = condition.value;
  let dataField = condition.dataField ?? 0;
  const initialTop10 = condition.kind === null ? null : pivotTop10Parts(condition.kind);
  let top10Direction: PivotTop10Direction = initialTop10?.direction ?? 'top';
  let top10Basis: PivotTop10Basis = initialTop10?.basis ?? 'items';
  const conditionCategory = createDialogSelect(
    [
      { value: 'label', label: t.filterConditionCategoryLabel },
      { value: 'value', label: t.filterConditionCategoryValue },
      { value: 'date', label: t.filterConditionCategoryDate },
    ],
    condition.kind === null ? 'label' : pivotFilterCategory(condition.kind),
    { className: options.selectClassName },
  );
  applyDataset(conditionCategory, options.categoryDataset);

  const conditionSelect = createDialogSelect([], '', { className: options.selectClassName });
  applyDataset(conditionSelect, options.conditionDataset);
  const fillConditionOptions = (): void =>
    appendFilterConditionOptions(
      conditionSelect,
      conditionCategory.value as PivotFilterCategory,
      t,
      datePeriods,
    );
  fillConditionOptions();
  const initialKind =
    condition.kind === null ? '' : initialTop10 ? TOP10_OPTION_VALUE : String(condition.kind);
  conditionSelect.value = initialKind;
  if (conditionSelect.value !== initialKind) conditionSelect.value = '';
  const isTop10 = (): boolean => conditionSelect.value === TOP10_OPTION_VALUE;
  const selectedKind = (): PivotFilterKind | null =>
    conditionSelect.value === ''
      ? null
      : isTop10()
        ? pivotTop10Kind(top10Direction, top10Basis)
        : (Number(conditionSelect.value) as PivotFilterKind);
  const conditionValueControls = document.createElement('div');
  conditionValueControls.className = options.valuesContainerClassName;
  const syncCondition = (value: string): void => {
    currentValue = value;
    const kind = selectedKind();
    const scored = kind !== null && conditionCategory.value === 'value' && dataField > 0;
    options.onChange({ kind, value, ...(scored ? { dataField } : {}) });
  };
  const renderDataFieldRow = (): void => {
    if (conditionCategory.value !== 'value' || dataFields.length === 0) return;
    const select = createDialogSelect(
      dataFields.map((name, index) => ({ value: String(index), label: name })),
      String(dataField),
      { className: options.selectClassName },
    );
    select.dataset.pivotFilterDataField = 'true';
    select.addEventListener('change', () => {
      options.onUserChange?.();
      dataField = Number(select.value);
      syncCondition(currentValue);
    });
    conditionValueControls.append(options.fieldRow(dialogs.top10.by, select));
  };
  const renderTop10Controls = (): void => {
    const direction = createDialogSelect(
      [
        { value: 'top', label: dialogs.top10.top },
        { value: 'bottom', label: dialogs.top10.bottom },
      ],
      top10Direction,
      { className: options.selectClassName },
    );
    direction.dataset.pivotFilterTop10Direction = 'true';
    const count = document.createElement('input');
    count.type = 'number';
    count.min = '1';
    count.step = '1';
    count.className = options.valueClassName;
    count.value = currentValue || '10';
    count.placeholder = t.filterConditionTopCountPlaceholder;
    count.dataset.pivotFilterTop10Count = 'true';
    const basis = createDialogSelect(
      [
        { value: 'items', label: dialogs.top10.items },
        { value: 'percent', label: dialogs.top10.percent },
        { value: 'sum', label: dialogs.top10.sum },
      ],
      top10Basis,
      { className: options.selectClassName },
    );
    basis.dataset.pivotFilterTop10Basis = 'true';
    direction.addEventListener('change', () => {
      options.onUserChange?.();
      top10Direction = direction.value as PivotTop10Direction;
      syncCondition(count.value);
    });
    basis.addEventListener('change', () => {
      options.onUserChange?.();
      top10Basis = basis.value as PivotTop10Basis;
      syncCondition(count.value);
    });
    count.addEventListener('input', () => {
      options.onUserChange?.();
      syncCondition(count.value);
    });
    const group = document.createElement('div');
    group.append(direction, count, basis);
    conditionValueControls.append(options.fieldRow(dialogs.top10.show, group));
    renderDataFieldRow();
    syncCondition(count.value);
  };
  const renderConditionValueControls = (): void => {
    conditionValueControls.replaceChildren();
    if (isTop10()) {
      renderTop10Controls();
      return;
    }
    const kind = selectedKind();
    const operand = kind === null ? 'none' : pivotFilterOperand(kind);
    if (kind !== null) renderDataFieldRow();
    if (operand === 'none') {
      syncCondition('');
      return;
    }
    if (operand === 'textRange' || operand === 'numberRange' || operand === 'dateRange') {
      const [lowValue, highValue] =
        operand === 'numberRange'
          ? splitFilterConditionRange(currentValue)
          : splitFilterDateRange(currentValue);
      const bound = (text: string): HTMLInputElement => {
        const input = document.createElement('input');
        input.type = operand === 'numberRange' ? 'number' : 'text';
        input.className = options.valueClassName;
        input.value = text;
        return input;
      };
      const low = bound(lowValue);
      const high = bound(highValue);
      const syncRange = (): void => syncCondition(`${low.value}..${high.value}`);
      const syncRangeFromUser = (): void => {
        options.onUserChange?.();
        syncRange();
      };
      low.addEventListener('input', syncRangeFromUser);
      high.addEventListener('input', syncRangeFromUser);
      conditionValueControls.append(
        options.fieldRow(t.filterConditionLowValue, low),
        options.fieldRow(t.filterConditionHighValue, high),
      );
      syncRange();
      return;
    }
    const value = document.createElement('input');
    value.className = options.valueClassName;
    value.value = currentValue;
    if (operand === 'number') {
      value.type = 'number';
      value.placeholder = t.filterConditionNumberPlaceholder;
    } else if (operand === 'count') {
      value.type = 'number';
      value.min = '1';
      value.step = '1';
      value.value = currentValue || '10';
      value.placeholder = t.filterConditionTopCountPlaceholder;
    } else {
      // Dates are typed as text and read through the locale's date grammar, as cells read them.
      value.type = 'text';
      value.placeholder = operand === 'text' ? t.filterConditionTextPlaceholder : '';
    }
    value.addEventListener('input', () => {
      options.onUserChange?.();
      syncCondition(value.value);
    });
    conditionValueControls.append(options.fieldRow(t.filterConditionValue, value));
    syncCondition(value.value);
  };
  conditionCategory.addEventListener('change', () => {
    options.onUserChange?.();
    fillConditionOptions();
    conditionSelect.value = '';
    renderConditionValueControls();
  });
  conditionSelect.addEventListener('change', () => {
    options.onUserChange?.();
    renderConditionValueControls();
  });
  renderConditionValueControls();
  return [
    options.fieldRow(t.filterConditionCategory, conditionCategory),
    options.fieldRow(t.filterCondition, conditionSelect),
    conditionValueControls,
  ];
};

export const renderPivotFieldSettingsPanel = (options: PivotFieldSettingsPanelOptions): void => {
  const { panelEl, active, strings: t, fields, controls } = options;
  if (!active) {
    panelEl.hidden = true;
    panelEl.replaceChildren();
    return;
  }

  const title = document.createElement('strong');
  title.textContent = t.fieldSettingsFor.replace('{field}', active.fieldName);
  const detail = document.createElement('span');
  detail.textContent = helpFor(active.kind, t);
  const settingsControls = document.createElement('div');
  settingsControls.className = 'fc-pivotdlg__settings-grid';

  if (active.kind === 'rows') {
    const sort = createDialogSelect([], '', { className: 'fc-fmtdlg__select' });
    cloneSelectOptions(controls.rowSortSelect, sort);
    sort.addEventListener('change', () => {
      controls.rowSortSelect.value = sort.value;
    });
    const subtotal = document.createElement('input');
    subtotal.type = 'checkbox';
    subtotal.checked = controls.rowSubtotalTop.checked;
    subtotal.addEventListener('change', () => {
      controls.rowSubtotalTop.checked = subtotal.checked;
    });
    settingsControls.append(
      settingsFieldRow(t.rowSort, sort),
      settingsCheckRow(t.rowSubtotalTop, subtotal),
    );
  } else if (active.kind === 'columns') {
    const sort = createDialogSelect([], '', { className: 'fc-fmtdlg__select' });
    cloneSelectOptions(controls.colSortSelect, sort);
    sort.addEventListener('change', () => {
      controls.colSortSelect.value = sort.value;
    });
    const subtotal = document.createElement('input');
    subtotal.type = 'checkbox';
    subtotal.checked = controls.colSubtotalTop.checked;
    subtotal.addEventListener('change', () => {
      controls.colSubtotalTop.checked = subtotal.checked;
    });
    settingsControls.append(
      settingsFieldRow(t.columnSort, sort),
      settingsCheckRow(t.columnSubtotalTop, subtotal),
    );
  } else if (active.kind === 'values') {
    const setting = options.selectedValueSetting(active.fieldName);
    const aggregation = createDialogSelect([], '', { className: 'fc-fmtdlg__select' });
    cloneSelectOptions(controls.aggSelect, aggregation);
    aggregation.value =
      setting.aggregation === undefined ? controls.aggSelect.value : String(setting.aggregation);
    aggregation.addEventListener('change', () => {
      options.setValueFieldSetting(active.fieldName, {
        ...options.selectedValueSetting(active.fieldName),
        aggregation: Number(aggregation.value) as PivotAggregation,
      });
    });
    const eligibleBaseFields = [controls.rowSelect.value, controls.colSelect.value].filter(
      (name) => name !== '',
    );
    const baseItemCount = (fieldName: string): number => options.inferFilterItems(fieldName).length;
    const resolved = resolvePivotShowAsBase(setting, eligibleBaseFields, baseItemCount);
    // The desktop order. Running Total is one entry; it runs along the row
    // field when there is one, else along the column field.
    const showValuesAs = createDialogSelect(
      [
        { value: String(PivotShowValuesAs.Normal), label: t.showValuesAsNormal },
        { value: String(PivotShowValuesAs.DifferenceFrom), label: t.showValuesAsDifferenceFrom },
        { value: String(PivotShowValuesAs.PercentOf), label: t.showValuesAsPercentOf },
        {
          value: String(PivotShowValuesAs.PercentDifferenceFrom),
          label: t.showValuesAsPercentDifferenceFrom,
        },
        { value: String(PivotShowValuesAs.RunningTotalInRow), label: t.showValuesAsRunningTotal },
        { value: String(PivotShowValuesAs.PercentOfRow), label: t.showValuesAsPercentOfRow },
        { value: String(PivotShowValuesAs.PercentOfCol), label: t.showValuesAsPercentOfColumn },
        { value: String(PivotShowValuesAs.PercentOfTotal), label: t.showValuesAsPercentOfTotal },
        { value: String(PivotShowValuesAs.Index), label: t.showValuesAsIndex },
        {
          value: String(PivotShowValuesAs.PercentOfParentRow),
          label: t.showValuesAsPercentOfParentRow,
        },
        {
          value: String(PivotShowValuesAs.PercentOfParentCol),
          label: t.showValuesAsPercentOfParentColumn,
        },
        {
          value: String(PivotShowValuesAs.PercentOfParent),
          label: t.showValuesAsPercentOfParentTotal,
        },
        {
          value: String(PivotShowValuesAs.PercentOfRunningTotal),
          label: t.showValuesAsPercentRunningTotal,
        },
        { value: String(PivotShowValuesAs.RankAscending), label: t.showValuesAsRankAscending },
        { value: String(PivotShowValuesAs.RankDescending), label: t.showValuesAsRankDescending },
      ],
      String(
        setting.showValuesAs === PivotShowValuesAs.RunningTotalInCol
          ? PivotShowValuesAs.RunningTotalInRow
          : (setting.showValuesAs ?? PivotShowValuesAs.Normal),
      ),
      { className: 'fc-fmtdlg__select' },
    );
    const storeShowAs = (next: PivotValueFieldSetting): void => {
      const after = resolvePivotShowAsBase(next, eligibleBaseFields, baseItemCount);
      options.setValueFieldSetting(active.fieldName, after);
      // Re-render only when the set of base controls changes, to keep focus otherwise.
      if (
        after.showAsBaseField !== resolved.showAsBaseField ||
        (after.showAsBaseItem === undefined) !== (resolved.showAsBaseItem === undefined)
      ) {
        renderPivotFieldSettingsPanel(options);
      }
    };
    showValuesAs.addEventListener('change', () => {
      storeShowAs({
        ...options.selectedValueSetting(active.fieldName),
        showValuesAs:
          Number(showValuesAs.value) === PivotShowValuesAs.RunningTotalInRow &&
          !controls.rowSelect.value &&
          controls.colSelect.value
            ? PivotShowValuesAs.RunningTotalInCol
            : (Number(showValuesAs.value) as PivotShowValuesAsValue),
      });
    });
    const baseControls: HTMLElement[] = [];
    if (resolved.showAsBaseField !== undefined) {
      const baseField = createDialogSelect(
        eligibleBaseFields.map((name) => ({ value: name, label: name })),
        resolved.showAsBaseField,
        { className: 'fc-fmtdlg__select' },
      );
      baseField.addEventListener('change', () => {
        storeShowAs({
          ...options.selectedValueSetting(active.fieldName),
          showAsBaseField: baseField.value,
          showAsBaseItem: undefined,
        });
      });
      baseControls.push(settingsFieldRow(t.showValuesAsBaseField, baseField));
    }
    if (resolved.showAsBaseField !== undefined && resolved.showAsBaseItem !== undefined) {
      const baseItem = createDialogSelect(
        [
          ...options
            .inferFilterItems(resolved.showAsBaseField)
            .map((name, index) => ({ value: String(index), label: name })),
          {
            value: String(PIVOT_SHOW_AS_BASE_PREVIOUS),
            label: t.showValuesAsBaseItemPrevious,
          },
          { value: String(PIVOT_SHOW_AS_BASE_NEXT), label: t.showValuesAsBaseItemNext },
        ],
        String(resolved.showAsBaseItem),
        { className: 'fc-fmtdlg__select' },
      );
      baseItem.addEventListener('change', () => {
        storeShowAs({
          ...options.selectedValueSetting(active.fieldName),
          showAsBaseField: resolved.showAsBaseField,
          showAsBaseItem: Number(baseItem.value),
        });
      });
      baseControls.push(settingsFieldRow(t.showValuesAsBaseItem, baseItem));
    }
    const format = document.createElement('input');
    format.type = 'text';
    format.className = 'fc-namedlg__input';
    format.placeholder = t.numberFormatPlaceholder;
    format.value = setting.numberFormat ?? controls.numberFormatInput.value;
    format.addEventListener('input', () => {
      options.setValueFieldSetting(active.fieldName, {
        ...options.selectedValueSetting(active.fieldName),
        numberFormat: format.value,
      });
    });
    settingsControls.append(
      settingsFieldRow(t.aggregation, aggregation),
      settingsFieldRow(t.showValuesAs, showValuesAs),
      ...baseControls,
      settingsFieldRow(t.numberFormat, format),
    );
  } else {
    const filter = createDialogSelect([], '', { className: 'fc-fmtdlg__select' });
    fillPivotFieldSelect(
      filter,
      fields.filter(
        (field) =>
          field.name === active.fieldName ||
          (field.name !== controls.rowSelect.value &&
            field.name !== controls.colSelect.value &&
            !options.selectedValueFields.includes(field.name) &&
            !options.selectedFilterFields.includes(field.name)),
      ),
    );
    filter.value = active.fieldName;
    filter.addEventListener('change', () => {
      options.replaceFilterField(active.fieldName, filter.value);
      options.normalizeSelectedFilters();
      options.refreshFieldList();
    });
    settingsControls.append(settingsFieldRow(t.filterField, filter));

    settingsControls.append(
      ...createPivotFilterConditionControls({
        strings: t,
        datePeriods: options.datePeriods,
        dataFields: options.selectedValueFields,
        condition: options.selectedFilterCondition(active.fieldName),
        selectClassName: 'fc-fmtdlg__select',
        valueClassName: 'fc-namedlg__input',
        valuesContainerClassName: 'fc-pivotdlg__settings-condition-values',
        categoryDataset: { pivotFilterCategory: 'true' },
        conditionDataset: { pivotFilterCondition: 'true' },
        fieldRow: settingsFieldRow,
        onChange: (condition) => options.setFilterCondition(active.fieldName, condition),
      }),
    );
    if (options.host && options.okLabel && options.cancelLabel) {
      const filterDialogButton = createPivotAreaSettingsButton(t.filterDialog);
      filterDialogButton.addEventListener('click', () => {
        void showPivotFilterDialog({
          host: options.host as HTMLElement,
          strings: t,
          fieldName: active.fieldName,
          condition: options.selectedFilterCondition(active.fieldName),
          okLabel: options.okLabel as string,
          cancelLabel: options.cancelLabel as string,
          dataFields: options.selectedValueFields,
          mac: isMacHost(options.host as HTMLElement),
        }).then((condition) => {
          if (!condition) return;
          options.setFilterCondition(active.fieldName, condition);
          renderPivotFieldSettingsPanel(options);
        });
      });
      const dialogRow = document.createElement('div');
      dialogRow.className = 'fc-pivotdlg__settings-field';
      const dialogLabel = document.createElement('span');
      dialogLabel.textContent = t.filterCondition;
      dialogRow.append(dialogLabel, filterDialogButton);
      settingsControls.append(dialogRow);
    }

    const items = options.inferFilterItems(active.fieldName);
    if (items.length > 0) {
      const itemList = document.createElement('div');
      itemList.className = 'fc-pivotdlg__settings-items';
      const itemLabel = document.createElement('span');
      itemLabel.textContent = t.filterItems;
      const itemGrid = document.createElement('div');
      itemGrid.className = 'fc-pivotdlg__settings-item-grid';
      for (const item of items) {
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.checked = options.filterItemVisibility(active.fieldName, item);
        checkbox.addEventListener('change', () => {
          options.setFilterItemVisibility(active.fieldName, item, checkbox.checked);
        });
        const label = document.createElement('label');
        label.className = 'fc-pivotdlg__settings-check';
        label.append(checkbox, document.createTextNode(item));
        itemGrid.appendChild(label);
      }
      itemList.append(itemLabel, itemGrid);
      settingsControls.append(itemList);
    }
  }

  panelEl.replaceChildren(title, detail, settingsControls);
  panelEl.hidden = false;
};
