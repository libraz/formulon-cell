import type { PivotSourceField } from '../commands/pivot-table.js';
import {
  PivotAxis,
  type PivotFilterSpec,
  PivotFilterType,
  PivotFilterValueKind,
} from '../engine/types.js';
import type { Strings } from '../i18n/strings.js';
import { appendDialogSelectOptions, createDialogSelect } from '../toolbar/dialogs/form-controls.js';
import {
  appendDialogActions,
  appendDialogFrame,
  createDialogButton,
  createDialogShell,
} from './dialog-shell.js';

export type PivotAreaKind = 'filters' | 'columns' | 'rows' | 'values';

export interface PivotFieldSettingsActive {
  kind: PivotAreaKind;
  fieldName: string;
}

export const createPivotAreaSettingsButton = (
  label: string,
  ariaLabel = label,
): HTMLButtonElement => {
  const button = createDialogButton({ label, baseClass: 'fc-pivotdlg__area-settings' });
  button.setAttribute('aria-label', ariaLabel);
  return button;
};

export type PivotFilterConditionKind =
  | 'none'
  | 'label-equals'
  | 'label-does-not-equal'
  | 'label-contains'
  | 'label-does-not-contain'
  | 'label-begins-with'
  | 'label-ends-with'
  | 'label-date'
  | 'date-before'
  | 'date-after'
  | 'date-between'
  | 'value-less-than'
  | 'value-equals'
  | 'value-greater-than'
  | 'value-between'
  | 'value-not-between'
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
  host?: HTMLElement;
  panelEl: HTMLDivElement;
  active: PivotFieldSettingsActive | null;
  strings: Strings['pivotTableDialog'];
  okLabel?: string;
  cancelLabel?: string;
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

const fieldSelect = (select: HTMLSelectElement, fields: readonly PivotSourceField[]): void => {
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

export const categoryForFilterCondition = (
  kind: PivotFilterConditionKind,
): PivotFilterConditionCategory => {
  if (
    kind === 'label-date' ||
    kind === 'date-before' ||
    kind === 'date-after' ||
    kind === 'date-between'
  )
    return 'date';
  if (
    kind === 'value-greater-than' ||
    kind === 'value-less-than' ||
    kind === 'value-equals' ||
    kind === 'value-between' ||
    kind === 'value-not-between' ||
    kind === 'value-top-10'
  ) {
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
  const options: Array<{ value: PivotFilterConditionKind; label: string }> = [
    { value: 'none', label: strings.filterConditionNone },
  ];
  if (category === 'label') {
    options.push(
      { value: 'label-equals', label: strings.filterConditionLabelEquals },
      { value: 'label-does-not-equal', label: strings.filterConditionLabelDoesNotEqual },
      { value: 'label-contains', label: strings.filterConditionLabelContains },
      { value: 'label-does-not-contain', label: strings.filterConditionLabelDoesNotContain },
      { value: 'label-begins-with', label: strings.filterConditionLabelBeginsWith },
      { value: 'label-ends-with', label: strings.filterConditionLabelEndsWith },
    );
  } else if (category === 'date') {
    options.push(
      { value: 'label-date', label: strings.filterConditionLabelDate },
      { value: 'date-before', label: strings.filterConditionDateBefore },
      { value: 'date-after', label: strings.filterConditionDateAfter },
      { value: 'date-between', label: strings.filterConditionDateBetween },
    );
  } else {
    options.push(
      { value: 'value-greater-than', label: strings.filterConditionValueGreaterThan },
      { value: 'value-less-than', label: strings.filterConditionValueLessThan },
      { value: 'value-equals', label: strings.filterConditionValueEquals },
      { value: 'value-between', label: strings.filterConditionValueBetween },
      { value: 'value-not-between', label: strings.filterConditionValueNotBetween },
      { value: 'value-top-10', label: strings.filterConditionValueTop10 },
    );
  }
  appendDialogSelectOptions(select, options);
};

const parseFilterNumberRange = (value: string): [number, number] | null => {
  const [lowText, highText] = splitFilterConditionRange(value);
  if (!lowText || !highText) return null;
  const low = Number(lowText);
  const high = Number(highText);
  if (!Number.isFinite(low) || !Number.isFinite(high)) return null;
  return [low, high];
};

const splitFilterDateRange = (value: string): [string, string] => {
  const parts = value.split(/\s*(?:\.\.|,|–)\s*/);
  return [parts[0]?.trim() ?? '', parts[1]?.trim() ?? ''];
};

export const pivotFilterConditionToSpec = (
  fieldName: string,
  condition: PivotFilterConditionState | undefined,
  axis: PivotAxis = PivotAxis.Page,
): PivotFilterSpec | null => {
  if (!condition || condition.kind === 'none') return null;
  const valueText = condition.value.trim();
  if (!valueText) return null;
  if (
    condition.kind === 'label-equals' ||
    condition.kind === 'label-does-not-equal' ||
    condition.kind === 'label-contains' ||
    condition.kind === 'label-does-not-contain' ||
    condition.kind === 'label-begins-with' ||
    condition.kind === 'label-ends-with'
  ) {
    const typeByKind: Record<
      | 'label-equals'
      | 'label-does-not-equal'
      | 'label-contains'
      | 'label-does-not-contain'
      | 'label-begins-with'
      | 'label-ends-with',
      PivotFilterType
    > = {
      'label-equals': PivotFilterType.LabelEquals,
      'label-does-not-equal': PivotFilterType.LabelDoesNotEqual,
      'label-contains': PivotFilterType.LabelContains,
      'label-does-not-contain': PivotFilterType.LabelDoesNotContain,
      'label-begins-with': PivotFilterType.LabelBeginsWith,
      'label-ends-with': PivotFilterType.LabelEndsWith,
    };
    return {
      axis,
      fieldName,
      type: typeByKind[condition.kind],
      valueKind: PivotFilterValueKind.Text,
      valueText,
    };
  }
  if (
    condition.kind === 'label-date' ||
    condition.kind === 'date-before' ||
    condition.kind === 'date-after'
  ) {
    const type =
      condition.kind === 'date-before'
        ? PivotFilterType.DateBefore
        : condition.kind === 'date-after'
          ? PivotFilterType.DateAfter
          : PivotFilterType.LabelDate;
    return {
      axis,
      fieldName,
      type,
      valueKind: PivotFilterValueKind.Text,
      valueText,
    };
  }
  if (condition.kind === 'date-between') {
    const [lowText, highText] = splitFilterDateRange(valueText);
    if (!lowText || !highText) return null;
    return {
      axis,
      fieldName,
      type: PivotFilterType.DateBetween,
      valueKind: PivotFilterValueKind.Text,
      valueText: lowText,
      valueHighKind: PivotFilterValueKind.Text,
      valueHighText: highText,
    };
  }
  if (condition.kind === 'value-between' || condition.kind === 'value-not-between') {
    const range = parseFilterNumberRange(valueText);
    if (!range) return null;
    return {
      axis,
      fieldName,
      type:
        condition.kind === 'value-between'
          ? PivotFilterType.ValueBetween
          : PivotFilterType.ValueNotBetween,
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
  const singleValueType =
    condition.kind === 'value-less-than'
      ? PivotFilterType.ValueLessThan
      : condition.kind === 'value-equals'
        ? PivotFilterType.ValueEquals
        : PivotFilterType.ValueGreaterThan;
  return {
    axis,
    fieldName,
    type: singleValueType,
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
  if (spec.type === PivotFilterType.LabelEquals && spec.valueText) {
    return { kind: 'label-equals', value: spec.valueText };
  }
  if (spec.type === PivotFilterType.LabelDoesNotEqual && spec.valueText) {
    return { kind: 'label-does-not-equal', value: spec.valueText };
  }
  if (spec.type === PivotFilterType.LabelDoesNotContain && spec.valueText) {
    return { kind: 'label-does-not-contain', value: spec.valueText };
  }
  if (spec.type === PivotFilterType.LabelEndsWith && spec.valueText) {
    return { kind: 'label-ends-with', value: spec.valueText };
  }
  if (spec.type === PivotFilterType.LabelDate && spec.valueText) {
    return { kind: 'label-date', value: spec.valueText };
  }
  if (spec.type === PivotFilterType.DateBefore && spec.valueText) {
    return { kind: 'date-before', value: spec.valueText };
  }
  if (spec.type === PivotFilterType.DateAfter && spec.valueText) {
    return { kind: 'date-after', value: spec.valueText };
  }
  if (spec.type === PivotFilterType.DateBetween) {
    if (!spec.valueText || !spec.valueHighText) return null;
    return { kind: 'date-between', value: `${spec.valueText}..${spec.valueHighText}` };
  }
  if (spec.type === PivotFilterType.ValueBetween || spec.type === PivotFilterType.ValueNotBetween) {
    if (
      typeof spec.valueDouble !== 'number' ||
      typeof spec.valueHighDouble !== 'number' ||
      !Number.isFinite(spec.valueDouble) ||
      !Number.isFinite(spec.valueHighDouble)
    )
      return null;
    return {
      kind: spec.type === PivotFilterType.ValueBetween ? 'value-between' : 'value-not-between',
      value: `${spec.valueDouble}..${spec.valueHighDouble}`,
    };
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
  if (spec.type === PivotFilterType.ValueLessThan) {
    if (typeof spec.valueDouble !== 'number' || !Number.isFinite(spec.valueDouble)) return null;
    return { kind: 'value-less-than', value: String(spec.valueDouble) };
  }
  if (spec.type === PivotFilterType.ValueEquals) {
    if (typeof spec.valueDouble !== 'number' || !Number.isFinite(spec.valueDouble)) return null;
    return { kind: 'value-equals', value: String(spec.valueDouble) };
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

export interface PivotFilterDialogOptions {
  host: HTMLElement;
  strings: Strings['pivotTableDialog'];
  fieldName: string;
  condition?: PivotFilterConditionState;
  okLabel: string;
  cancelLabel: string;
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
  const conditionCategory = createDialogSelect(
    [
      { value: 'label', label: t.filterConditionCategoryLabel },
      { value: 'value', label: t.filterConditionCategoryValue },
      { value: 'date', label: t.filterConditionCategoryDate },
    ],
    categoryForFilterCondition(condition.kind),
    { className: options.selectClassName },
  );
  applyDataset(conditionCategory, options.categoryDataset);

  const conditionSelect = createDialogSelect([], '', { className: options.selectClassName });
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
    if (kind === 'value-between' || kind === 'value-not-between' || kind === 'date-between') {
      const [lowValue, highValue] =
        kind === 'date-between'
          ? splitFilterDateRange(currentValue)
          : splitFilterConditionRange(currentValue);
      const low = document.createElement('input');
      low.type = kind === 'date-between' ? 'date' : 'number';
      low.className = options.valueClassName;
      low.value = lowValue;
      const high = document.createElement('input');
      high.type = kind === 'date-between' ? 'date' : 'number';
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
    if (kind === 'label-date' || kind === 'date-before' || kind === 'date-after') {
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

export const showPivotFilterDialog = (
  options: PivotFilterDialogOptions,
): Promise<PivotFilterConditionState | null> =>
  new Promise<PivotFilterConditionState | null>((resolve) => {
    const title = options.strings.filterDialogTitle.replace('{field}', options.fieldName);
    let draft: PivotFilterConditionState = options.condition ?? { kind: 'none', value: '' };
    let done = false;
    const finish = (result: PivotFilterConditionState | null): void => {
      if (done) return;
      done = true;
      shell.dispose();
      resolve(result);
    };
    const shell = createDialogShell({
      host: options.host,
      className: 'fc-pivotdlg',
      ariaLabel: title,
      onDismiss: () => finish(null),
    });
    shell.overlay.classList.add('fc-fmtdlg');
    shell.overlay.classList.add('fc-pivotdlg--filter');
    const { body, footer } = appendDialogFrame(shell, {
      title,
      panelClasses: ['fc-fmtdlg__panel', 'fc-pivotdlg__panel'],
      bodyClass: 'fc-fmtdlg__body fc-pivotdlg__body',
    });
    const grid = document.createElement('div');
    grid.className = 'fc-pivotdlg__settings-grid';
    grid.append(
      ...createPivotFilterConditionControls({
        strings: options.strings,
        condition: draft,
        selectClassName: 'fc-fmtdlg__select',
        valueClassName: 'fc-namedlg__input',
        valuesContainerClassName: 'fc-pivotdlg__settings-condition-values',
        categoryDataset: { pivotFilterCategory: 'true' },
        conditionDataset: { pivotFilterCondition: 'true' },
        fieldRow: settingsFieldRow,
        onChange: (condition) => {
          draft = condition;
        },
      }),
    );
    body.appendChild(grid);
    const { cancelBtn: cancel, okBtn: ok } = appendDialogActions(footer, {
      cancelLabel: options.cancelLabel,
      okLabel: options.okLabel,
    });
    shell.on(cancel, 'click', () => finish(null));
    shell.on(ok, 'click', () => finish(draft));
    shell.open();
    requestAnimationFrame(() => {
      const category = grid.querySelector<HTMLSelectElement>(
        'select[data-pivot-filter-category="true"]',
      );
      const categoryButton =
        category?.closest('.fc-select')?.querySelector<HTMLButtonElement>('.fc-select__button') ??
        category;
      categoryButton?.focus({ preventScroll: true });
    });
  });

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
    const aggregation = createDialogSelect([], '', { className: 'fc-fmtdlg__select' });
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
    const filter = createDialogSelect([], '', { className: 'fc-fmtdlg__select' });
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
