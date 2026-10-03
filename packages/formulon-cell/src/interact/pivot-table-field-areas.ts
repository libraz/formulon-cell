import type { PivotSourceField } from '../commands/pivot-table.js';
import type { PivotAggregation, PivotFilterSpec } from '../engine/types.js';
import { appendDialogSelectOptions } from '../toolbar/dialogs/form-controls.js';
import {
  fillPivotFieldSelect,
  type PivotAreaKind,
  type PivotFieldSettingsActive,
  type PivotFilterConditionState,
  type PivotValueFieldSetting,
  pivotFilterConditionToSpec,
} from './pivot-field-settings.js';

export interface PivotFieldAreaSelects {
  rowSelect: HTMLSelectElement;
  colSelect: HTMLSelectElement;
  filterSelect: HTMLSelectElement;
  valueSelect: HTMLSelectElement;
}

export interface PivotCreateSpecDefaults {
  aggregation: PivotAggregation;
  numberFormat: string;
}

const retainFieldEntries = <V>(map: Map<string, V>, keep: readonly string[]): Map<string, V> =>
  new Map(Array.from(map.entries()).filter(([fieldName]) => keep.includes(fieldName)));

const optionValues = (select: HTMLSelectElement): Set<string> =>
  new Set(Array.from(select.options).map((option) => option.value));

const setFirstDifferent = (select: HTMLSelectElement, fieldName: string): void => {
  const next = Array.from(select.options).find(
    (option) => option.value && option.value !== fieldName,
  );
  select.value = next?.value ?? '';
};

/** Field-to-area assignment state shared by the dialog's selects and field list. */
export function createPivotFieldAreas(selects: PivotFieldAreaSelects) {
  const { rowSelect, colSelect, filterSelect, valueSelect } = selects;
  let selectedFilterFields: string[] = [];
  let selectedValueFields: string[] = [];
  let selectedFilterItemVisibility = new Map<string, Map<string, boolean>>();
  let selectedFilterConditions = new Map<string, PivotFilterConditionState>();
  let selectedValueSettings = new Map<string, PivotValueFieldSetting>();
  let activeFieldSettings: PivotFieldSettingsActive | null = null;

  const fieldCanBeValue = (fieldName: string): boolean => optionValues(valueSelect).has(fieldName);

  const normalizeSelectedFilters = (): void => {
    const filterNames = optionValues(filterSelect);
    selectedFilterFields = Array.from(
      new Set(
        selectedFilterFields.filter(
          (name) =>
            name &&
            filterNames.has(name) &&
            name !== rowSelect.value &&
            name !== colSelect.value &&
            !selectedValueFields.includes(name),
        ),
      ),
    );
    if (
      filterSelect.value &&
      filterSelect.value !== rowSelect.value &&
      filterSelect.value !== colSelect.value &&
      !selectedValueFields.includes(filterSelect.value) &&
      !selectedFilterFields.includes(filterSelect.value)
    ) {
      selectedFilterFields.unshift(filterSelect.value);
    }
    filterSelect.value = selectedFilterFields[0] ?? '';
  };

  const normalizeSelectedValues = (fields: readonly PivotSourceField[]): void => {
    const valueNames = optionValues(valueSelect);
    selectedValueFields = selectedValueFields.filter((name) => valueNames.has(name));
    if (valueSelect.value && !selectedValueFields.includes(valueSelect.value)) {
      selectedValueFields.unshift(valueSelect.value);
    }
    if (selectedValueFields.length === 0) {
      const fallback =
        fields.find((field) => field.numericCount > 0 && valueNames.has(field.name)) ??
        fields.find((field) => valueNames.has(field.name));
      if (fallback) selectedValueFields = [fallback.name];
    }
    valueSelect.value = selectedValueFields[0] ?? '';
  };

  const addSelectedFilter = (fieldName: string): void => {
    if (
      !fieldName ||
      fieldName === rowSelect.value ||
      fieldName === colSelect.value ||
      selectedValueFields.includes(fieldName) ||
      selectedFilterFields.includes(fieldName)
    )
      return;
    selectedFilterFields = [...selectedFilterFields, fieldName];
    filterSelect.value = selectedFilterFields[0] ?? fieldName;
  };

  const removeSelectedFilter = (fieldName: string): void => {
    selectedFilterFields = selectedFilterFields.filter((name) => name !== fieldName);
    selectedFilterItemVisibility.delete(fieldName);
    selectedFilterConditions.delete(fieldName);
    filterSelect.value = selectedFilterFields[0] ?? '';
  };

  const addSelectedValue = (fieldName: string): void => {
    if (!fieldCanBeValue(fieldName) || selectedValueFields.includes(fieldName)) return;
    selectedValueFields = [...selectedValueFields, fieldName];
    valueSelect.value = selectedValueFields[0] ?? fieldName;
  };

  const removeSelectedValue = (fieldName: string): void => {
    selectedValueFields = selectedValueFields.filter((name) => name !== fieldName);
    selectedValueSettings.delete(fieldName);
    valueSelect.value = selectedValueFields[0] ?? '';
  };

  const replaceFilterField = (previous: string, next: string): void => {
    if (
      !next ||
      next === rowSelect.value ||
      next === colSelect.value ||
      selectedValueFields.includes(next)
    ) {
      return;
    }
    selectedFilterFields = selectedFilterFields.map((field) => (field === previous ? next : field));
    selectedFilterFields = Array.from(new Set(selectedFilterFields));
    if (previous !== next) {
      selectedFilterItemVisibility.delete(previous);
      selectedFilterItemVisibility.delete(next);
      selectedFilterConditions.delete(previous);
      selectedFilterConditions.delete(next);
    }
    filterSelect.value = selectedFilterFields[0] ?? '';
    activeFieldSettings = { kind: 'filters', fieldName: next };
  };

  const filterItemVisibility = (fieldName: string, itemName: string): boolean =>
    selectedFilterItemVisibility.get(fieldName)?.get(itemName) ?? true;

  const setFilterItemVisibility = (fieldName: string, itemName: string, visible: boolean): void => {
    const byField = new Map(selectedFilterItemVisibility);
    const items = new Map(byField.get(fieldName) ?? []);
    items.set(itemName, visible);
    byField.set(fieldName, items);
    selectedFilterItemVisibility = byField;
  };

  const setFilterCondition = (fieldName: string, condition: PivotFilterConditionState): void => {
    const byField = new Map(selectedFilterConditions);
    if (condition.kind === 'none' || !condition.value.trim()) byField.delete(fieldName);
    else byField.set(fieldName, condition);
    selectedFilterConditions = byField;
  };

  const removeFieldAssignment = (fieldName: string): void => {
    if (rowSelect.value === fieldName) rowSelect.value = '';
    if (colSelect.value === fieldName) colSelect.value = '';
    selectedFilterFields = selectedFilterFields.filter((name) => name !== fieldName);
    selectedValueFields = selectedValueFields.filter((name) => name !== fieldName);
    selectedValueSettings.delete(fieldName);
    selectedFilterItemVisibility.delete(fieldName);
    selectedFilterConditions.delete(fieldName);
  };

  const selectedValueSetting = (fieldName: string): PivotValueFieldSetting =>
    selectedValueSettings.get(fieldName) ?? {};

  const setValueFieldSetting = (fieldName: string, setting: PivotValueFieldSetting): void => {
    selectedValueSettings = new Map(selectedValueSettings);
    const numberFormat = setting.numberFormat?.trim() ?? '';
    const next = {
      ...(setting.aggregation === undefined ? {} : { aggregation: setting.aggregation }),
      ...(numberFormat.length > 0 ? { numberFormat } : {}),
      ...(setting.showValuesAs === undefined ? {} : { showValuesAs: setting.showValuesAs }),
    };
    if (
      next.aggregation === undefined &&
      next.numberFormat === undefined &&
      next.showValuesAs === undefined
    ) {
      selectedValueSettings.delete(fieldName);
    } else {
      selectedValueSettings.set(fieldName, next);
    }
  };

  const assignFieldToArea = (
    fieldName: string,
    kind: PivotAreaKind,
    fields: readonly PivotSourceField[],
  ): void => {
    if (!fields.some((field) => field.name === fieldName)) return;
    removeFieldAssignment(fieldName);
    if (kind === 'filters') addSelectedFilter(fieldName);
    else if (kind === 'columns') colSelect.value = fieldName;
    else if (kind === 'rows') rowSelect.value = fieldName;
    else if (fieldCanBeValue(fieldName)) addSelectedValue(fieldName);
    normalizeSelectedValues(fields);
    normalizeSelectedFilters();
  };

  /** Applies a field-list checkbox change to the area assignment. */
  const toggleFromFieldList = (
    name: string,
    checked: boolean,
    fields: readonly PivotSourceField[],
  ): void => {
    if (checked) {
      if (fieldCanBeValue(name)) addSelectedValue(name);
      else if (!rowSelect.value) rowSelect.value = name;
      else if (!colSelect.value && rowSelect.value !== name && valueSelect.value !== name)
        colSelect.value = name;
      else if (
        !filterSelect.value &&
        rowSelect.value !== name &&
        colSelect.value !== name &&
        valueSelect.value !== name
      )
        addSelectedFilter(name);
      else if (!fieldCanBeValue(name) && rowSelect.value !== name && colSelect.value !== name)
        addSelectedFilter(name);
    } else if (selectedFilterFields.includes(name)) {
      removeSelectedFilter(name);
    } else if (colSelect.value === name) {
      colSelect.value = '';
    } else if (rowSelect.value === name) {
      setFirstDifferent(rowSelect, name);
    } else if (selectedValueFields.includes(name)) {
      removeSelectedValue(name);
      if (selectedValueFields.length === 0) setFirstDifferent(valueSelect, name);
    }
    normalizeSelectedValues(fields);
    normalizeSelectedFilters();
  };

  /** Rebuilds the select options for a new source and reconciles prior assignments. */
  const resetSource = (fields: readonly PivotSourceField[], noneLabel: string): void => {
    const numeric = fields.filter((f) => f.numericCount > 0);
    const prevRow = rowSelect.value;
    const prevCol = colSelect.value;
    const prevFilters =
      selectedFilterFields.length > 0 ? selectedFilterFields : [filterSelect.value];
    const prevValues = selectedValueFields.length > 0 ? selectedValueFields : [valueSelect.value];
    fillPivotFieldSelect(rowSelect, fields);
    fillPivotFieldSelect(colSelect, fields);
    fillPivotFieldSelect(filterSelect, fields);
    fillPivotFieldSelect(valueSelect, numeric.length > 0 ? numeric : fields);
    for (const select of [colSelect, filterSelect]) {
      const currentOptions = Array.from(select.options).map((option) => ({
        value: option.value,
        label: option.textContent ?? '',
      }));
      select.replaceChildren();
      appendDialogSelectOptions(select, [{ value: '', label: noneLabel }, ...currentOptions]);
    }
    const rowValues = optionValues(rowSelect);
    const colValues = optionValues(colSelect);
    const filterValues = optionValues(filterSelect);
    const valueValues = optionValues(valueSelect);
    rowSelect.value = rowValues.has(prevRow) ? prevRow : (fields[0]?.name ?? '');
    selectedValueFields = prevValues.filter((name) => valueValues.has(name));
    valueSelect.value = selectedValueFields[0]
      ? selectedValueFields[0]
      : ((numeric[0] ?? fields[fields.length - 1])?.name ?? '');
    normalizeSelectedValues(fields);
    selectedValueSettings = retainFieldEntries(selectedValueSettings, selectedValueFields);
    colSelect.value = colValues.has(prevCol)
      ? prevCol
      : fields[1]?.name === selectedValueFields[0]
        ? ''
        : (fields[1]?.name ?? '');
    selectedFilterFields = prevFilters.filter(
      (name) =>
        name &&
        filterValues.has(name) &&
        name !== rowSelect.value &&
        name !== colSelect.value &&
        !selectedValueFields.includes(name),
    );
    selectedFilterItemVisibility = retainFieldEntries(
      selectedFilterItemVisibility,
      selectedFilterFields,
    );
    selectedFilterConditions = retainFieldEntries(selectedFilterConditions, selectedFilterFields);
    filterSelect.value = selectedFilterFields[0] ?? '';
    normalizeSelectedFilters();
  };

  /** Drops the active settings target once its field leaves its area. */
  const pruneActiveSettings = (): void => {
    const active = activeFieldSettings;
    const present =
      active &&
      (active.kind === 'filters'
        ? selectedFilterFields.includes(active.fieldName)
        : active.kind === 'columns'
          ? colSelect.value === active.fieldName
          : active.kind === 'rows'
            ? rowSelect.value === active.fieldName
            : selectedValueFields.includes(active.fieldName));
    if (!present) activeFieldSettings = null;
  };

  const selectSingleValue = (): void => {
    selectedValueFields = valueSelect.value ? [valueSelect.value] : [];
    selectedValueSettings = retainFieldEntries(selectedValueSettings, selectedValueFields);
  };

  const selectSingleFilter = (): void => {
    selectedFilterFields = filterSelect.value ? [filterSelect.value] : [];
  };

  /** Flattens the assignment state into the create-command payload. */
  const toCreateSpec = (defaults: PivotCreateSpecDefaults) => ({
    filterItems: selectedFilterFields.flatMap((fieldName) =>
      Array.from(selectedFilterItemVisibility.get(fieldName)?.entries() ?? []).map(
        ([itemName, visible]) => ({
          fieldName,
          itemName,
          visible,
        }),
      ),
    ),
    pivotFilters: selectedFilterFields.flatMap<PivotFilterSpec>((fieldName) => {
      const spec = pivotFilterConditionToSpec(fieldName, selectedFilterConditions.get(fieldName));
      return spec ? [spec] : [];
    }),
    valueFieldSettings: selectedValueFields.map((fieldName) => {
      const setting = selectedValueSetting(fieldName);
      return {
        fieldName,
        aggregation: setting.aggregation === undefined ? defaults.aggregation : setting.aggregation,
        numberFormat: setting.numberFormat ?? defaults.numberFormat,
        showValuesAs: setting.showValuesAs,
      };
    }),
  });

  return {
    get filterFields(): readonly string[] {
      return selectedFilterFields;
    },
    get valueFields(): readonly string[] {
      return selectedValueFields;
    },
    get active(): PivotFieldSettingsActive | null {
      return activeFieldSettings;
    },
    setActive(next: PivotFieldSettingsActive | null): void {
      activeFieldSettings = next;
    },
    fieldCanBeValue,
    normalizeSelectedFilters,
    normalizeSelectedValues,
    replaceFilterField,
    filterItemVisibility,
    setFilterItemVisibility,
    selectedFilterCondition: (fieldName: string): PivotFilterConditionState | undefined =>
      selectedFilterConditions.get(fieldName),
    setFilterCondition,
    selectedValueSetting,
    setValueFieldSetting,
    assignFieldToArea,
    toggleFromFieldList,
    resetSource,
    pruneActiveSettings,
    selectSingleValue,
    selectSingleFilter,
    toCreateSpec,
  };
}

export type PivotFieldAreas = ReturnType<typeof createPivotFieldAreas>;
