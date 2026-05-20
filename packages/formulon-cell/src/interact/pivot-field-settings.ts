import type { PivotSourceField } from '../commands/pivot-table.js';
import {
  PivotAxis,
  type PivotFilterSpec,
  PivotFilterType,
  PivotFilterValueKind,
} from '../engine/types.js';
import type { Strings } from '../i18n/strings.js';

export type PivotAreaKind = 'filters' | 'columns' | 'rows' | 'values';

export interface PivotFieldSettingsActive {
  kind: PivotAreaKind;
  fieldName: string;
}

export type PivotFilterConditionKind =
  | 'none'
  | 'label-contains'
  | 'label-begins-with'
  | 'label-date'
  | 'value-greater-than'
  | 'value-between'
  | 'value-top-10';

export type PivotFilterConditionCategory = 'label' | 'value' | 'date';

export interface PivotFilterConditionState {
  kind: PivotFilterConditionKind;
  value: string;
}

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

const appendOption = (select: HTMLSelectElement, value: string, label: string): void => {
  const opt = document.createElement('option');
  opt.value = value;
  opt.textContent = label;
  select.appendChild(opt);
};

const appendUnsupportedFilterOption = (
  select: HTMLSelectElement,
  value: string,
  label: string,
  strings: Strings['pivotTableDialog'],
): void => {
  const opt = document.createElement('option');
  opt.value = value;
  opt.textContent = `${label} (${strings.filterConditionUnsupportedSuffix})`;
  opt.disabled = true;
  opt.dataset.unsupportedPivotFilter = 'true';
  select.appendChild(opt);
};

const cloneSelectOptions = (source: HTMLSelectElement, target: HTMLSelectElement): void => {
  target.replaceChildren();
  for (const option of source.options) appendOption(target, option.value, option.textContent ?? '');
  target.value = source.value;
};

const fieldSelect = (select: HTMLSelectElement, fields: readonly PivotSourceField[]): void => {
  select.replaceChildren();
  for (const f of fields) appendOption(select, f.name, f.name);
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

export const categoryForFilterCondition = (
  kind: PivotFilterConditionKind,
): PivotFilterConditionCategory => {
  if (kind === 'label-date') return 'date';
  if (kind === 'value-greater-than' || kind === 'value-between' || kind === 'value-top-10') {
    return 'value';
  }
  return 'label';
};

export const appendFilterConditionOptions = (
  select: HTMLSelectElement,
  category: PivotFilterConditionCategory,
  strings: Strings['pivotTableDialog'],
): void => {
  select.replaceChildren();
  appendOption(select, 'none', strings.filterConditionNone);
  if (category === 'label') {
    appendUnsupportedFilterOption(
      select,
      'unsupported-label-equals',
      strings.filterConditionLabelEquals,
      strings,
    );
    appendUnsupportedFilterOption(
      select,
      'unsupported-label-does-not-equal',
      strings.filterConditionLabelDoesNotEqual,
      strings,
    );
    appendOption(select, 'label-contains', strings.filterConditionLabelContains);
    appendUnsupportedFilterOption(
      select,
      'unsupported-label-does-not-contain',
      strings.filterConditionLabelDoesNotContain,
      strings,
    );
    appendOption(select, 'label-begins-with', strings.filterConditionLabelBeginsWith);
    appendUnsupportedFilterOption(
      select,
      'unsupported-label-ends-with',
      strings.filterConditionLabelEndsWith,
      strings,
    );
  } else if (category === 'date') {
    appendOption(select, 'label-date', strings.filterConditionLabelDate);
    appendUnsupportedFilterOption(
      select,
      'unsupported-date-before',
      strings.filterConditionDateBefore,
      strings,
    );
    appendUnsupportedFilterOption(
      select,
      'unsupported-date-after',
      strings.filterConditionDateAfter,
      strings,
    );
    appendUnsupportedFilterOption(
      select,
      'unsupported-date-between',
      strings.filterConditionDateBetween,
      strings,
    );
  } else {
    appendOption(select, 'value-greater-than', strings.filterConditionValueGreaterThan);
    appendUnsupportedFilterOption(
      select,
      'unsupported-value-less-than',
      strings.filterConditionValueLessThan,
      strings,
    );
    appendUnsupportedFilterOption(
      select,
      'unsupported-value-equals',
      strings.filterConditionValueEquals,
      strings,
    );
    appendOption(select, 'value-between', strings.filterConditionValueBetween);
    appendUnsupportedFilterOption(
      select,
      'unsupported-value-not-between',
      strings.filterConditionValueNotBetween,
      strings,
    );
    appendOption(select, 'value-top-10', strings.filterConditionValueTop10);
  }
};

const parseFilterNumberRange = (value: string): [number, number] | null => {
  const [lowText, highText] = splitFilterConditionRange(value);
  if (!lowText || !highText) return null;
  const low = Number(lowText);
  const high = Number(highText);
  if (!Number.isFinite(low) || !Number.isFinite(high)) return null;
  return [low, high];
};

export const pivotFilterConditionToSpec = (
  fieldName: string,
  condition: PivotFilterConditionState | undefined,
  axis: PivotAxis = PivotAxis.Page,
): PivotFilterSpec | null => {
  if (!condition || condition.kind === 'none') return null;
  const valueText = condition.value.trim();
  if (!valueText) return null;
  if (condition.kind === 'label-contains' || condition.kind === 'label-begins-with') {
    return {
      axis,
      fieldName,
      type:
        condition.kind === 'label-contains'
          ? PivotFilterType.LabelContains
          : PivotFilterType.LabelBeginsWith,
      valueKind: PivotFilterValueKind.Text,
      valueText,
    };
  }
  if (condition.kind === 'label-date') {
    return {
      axis,
      fieldName,
      type: PivotFilterType.LabelDate,
      valueKind: PivotFilterValueKind.Text,
      valueText,
    };
  }
  if (condition.kind === 'value-between') {
    const range = parseFilterNumberRange(valueText);
    if (!range) return null;
    return {
      axis,
      fieldName,
      type: PivotFilterType.ValueBetween,
      valueKind: PivotFilterValueKind.Double,
      valueDouble: range[0],
      valueHighKind: PivotFilterValueKind.Double,
      valueHighDouble: range[1],
    };
  }
  if (condition.kind === 'value-top-10') {
    const count = Number(valueText);
    return {
      axis,
      fieldName,
      type: PivotFilterType.ValueTop10,
      valueKind: PivotFilterValueKind.Int,
      valueInt: Number.isFinite(count) && count > 0 ? Math.floor(count) : 10,
    };
  }
  const value = Number(valueText);
  if (!Number.isFinite(value)) return null;
  return {
    axis,
    fieldName,
    type: PivotFilterType.ValueGreaterThan,
    valueKind: PivotFilterValueKind.Double,
    valueDouble: value,
  };
};

export const pivotFilterSpecToCondition = (
  spec: PivotFilterSpec | undefined,
): PivotFilterConditionState | null => {
  if (!spec) return null;
  if (spec.type === PivotFilterType.LabelContains && spec.valueText) {
    return { kind: 'label-contains', value: spec.valueText };
  }
  if (spec.type === PivotFilterType.LabelBeginsWith && spec.valueText) {
    return { kind: 'label-begins-with', value: spec.valueText };
  }
  if (spec.type === PivotFilterType.LabelDate && spec.valueText) {
    return { kind: 'label-date', value: spec.valueText };
  }
  if (spec.type === PivotFilterType.ValueBetween) {
    if (
      typeof spec.valueDouble !== 'number' ||
      typeof spec.valueHighDouble !== 'number' ||
      !Number.isFinite(spec.valueDouble) ||
      !Number.isFinite(spec.valueHighDouble)
    )
      return null;
    return { kind: 'value-between', value: `${spec.valueDouble}..${spec.valueHighDouble}` };
  }
  if (spec.type === PivotFilterType.ValueTop10) {
    const value =
      typeof spec.valueInt === 'number' && Number.isFinite(spec.valueInt) && spec.valueInt > 0
        ? spec.valueInt
        : 10;
    return { kind: 'value-top-10', value: String(value) };
  }
  if (spec.type === PivotFilterType.ValueGreaterThan) {
    if (typeof spec.valueDouble !== 'number' || !Number.isFinite(spec.valueDouble)) return null;
    return { kind: 'value-greater-than', value: String(spec.valueDouble) };
  }
  return null;
};

export interface PivotFilterConditionControlsOptions {
  strings: Strings['pivotTableDialog'];
  condition?: PivotFilterConditionState;
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

export const createPivotFilterConditionControls = (
  options: PivotFilterConditionControlsOptions,
): HTMLElement[] => {
  const { strings: t } = options;
  const condition = options.condition ?? { kind: 'none', value: '' };
  let currentValue = condition.value;
  const conditionCategory = document.createElement('select');
  conditionCategory.className = options.selectClassName;
  applyDataset(conditionCategory, options.categoryDataset);
  appendOption(conditionCategory, 'label', t.filterConditionCategoryLabel);
  appendOption(conditionCategory, 'value', t.filterConditionCategoryValue);
  appendOption(conditionCategory, 'date', t.filterConditionCategoryDate);
  conditionCategory.value = categoryForFilterCondition(condition.kind);

  const conditionSelect = document.createElement('select');
  conditionSelect.className = options.selectClassName;
  applyDataset(conditionSelect, options.conditionDataset);
  appendFilterConditionOptions(
    conditionSelect,
    conditionCategory.value as PivotFilterConditionCategory,
    t,
  );
  conditionSelect.value = condition.kind;
  if (conditionSelect.value !== condition.kind) conditionSelect.value = 'none';
  const conditionValueControls = document.createElement('div');
  conditionValueControls.className = options.valuesContainerClassName;
  const syncCondition = (value: string): void => {
    currentValue = value;
    options.onChange({
      kind: conditionSelect.value as PivotFilterConditionKind,
      value,
    });
  };
  const renderConditionValueControls = (): void => {
    conditionValueControls.replaceChildren();
    const kind = conditionSelect.value as PivotFilterConditionKind;
    if (kind === 'none') {
      syncCondition('');
      return;
    }
    if (kind === 'value-between') {
      const [lowValue, highValue] = splitFilterConditionRange(currentValue);
      const low = document.createElement('input');
      low.type = 'number';
      low.className = options.valueClassName;
      low.value = lowValue;
      const high = document.createElement('input');
      high.type = 'number';
      high.className = options.valueClassName;
      high.value = highValue;
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
    if (kind === 'label-date') {
      value.type = 'date';
      value.placeholder = 'yyyy-mm-dd';
    } else if (kind === 'value-greater-than') {
      value.type = 'number';
      value.placeholder = t.filterConditionNumberPlaceholder;
    } else if (kind === 'value-top-10') {
      value.type = 'number';
      value.min = '1';
      value.step = '1';
      value.value = currentValue || '10';
      value.placeholder = t.filterConditionTopCountPlaceholder;
    } else {
      value.type = 'text';
      value.placeholder = t.filterConditionTextPlaceholder;
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
    appendFilterConditionOptions(
      conditionSelect,
      conditionCategory.value as PivotFilterConditionCategory,
      t,
    );
    conditionSelect.value = 'none';
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
    const sort = document.createElement('select');
    sort.className = 'fc-fmtdlg__select';
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
    const sort = document.createElement('select');
    sort.className = 'fc-fmtdlg__select';
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
    const aggregation = document.createElement('select');
    aggregation.className = 'fc-fmtdlg__select';
    cloneSelectOptions(controls.aggSelect, aggregation);
    aggregation.addEventListener('change', () => {
      controls.aggSelect.value = aggregation.value;
    });
    const format = document.createElement('input');
    format.type = 'text';
    format.className = 'fc-namedlg__input';
    format.placeholder = t.numberFormatPlaceholder;
    format.value = controls.numberFormatInput.value;
    format.addEventListener('input', () => {
      controls.numberFormatInput.value = format.value;
    });
    settingsControls.append(
      settingsFieldRow(t.aggregation, aggregation),
      settingsFieldRow(t.numberFormat, format),
    );
  } else {
    const filter = document.createElement('select');
    filter.className = 'fc-fmtdlg__select';
    fieldSelect(
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
