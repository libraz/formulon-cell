import {
  inputLocaleForWorkbook,
  parseLocalizedNumericInput,
  workbookDate1904,
} from '../commands/coerce-input.js';
import { parseDateTimeInput } from '../commands/input-grammar.js';
import type { PivotSourceField } from '../commands/pivot-cache-source.js';
import {
  PIVOT_SHOW_AS_BASE_NEXT,
  PIVOT_SHOW_AS_BASE_PREVIOUS,
  type PivotAggregation,
  PivotAxis,
  type PivotFilterKind,
  type PivotFilterSpec,
  PivotShowValuesAs,
  type PivotShowValuesAs as PivotShowValuesAsValue,
} from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { excelSerialToIsoDate } from '../format/date-serial.js';
import type { Strings } from '../i18n/strings.js';
import { appendDialogSelectOptions, createDialogSelect } from '../toolbar/dialogs/form-controls.js';
import { createDialogButton } from './dialog-shell.js';
import { pivotFilterOperand } from './pivot-filter-kinds.js';

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
  panelEl: HTMLDivElement;
  active: PivotFieldSettingsActive | null;
  strings: Strings['pivotTableDialog'];
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
