import {
  createPivotTableFromRange,
  inferPivotFieldItems,
  inferPivotSourceFields,
  type PivotSourceField,
} from '../commands/pivot-table.js';
import { parseRangeRef } from '../engine/range-resolver.js';
import { PivotAggregation, type PivotFilterSpec, type PivotShowValuesAs } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { appendDialogSelectOptions, createDialogSelect } from '../toolbar/dialogs/form-controls.js';
import { projectDisabledReason, projectDisabledState } from '../toolbar/menu-a11y.js';
import { appendDialogActions, appendDialogFrame, createDialogShell } from './dialog-shell.js';
import {
  createPivotAreaSettingsButton,
  type PivotAreaKind,
  type PivotFieldSettingsActive,
  type PivotFilterConditionKind,
  type PivotFilterConditionState,
  pivotFilterConditionToSpec,
  renderPivotFieldSettingsPanel,
} from './pivot-field-settings.js';
import { attachRangePickerButton } from './range-picker-control.js';

export interface PivotTableDialogDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  wb: WorkbookHandle;
  strings?: Strings;
  onAfterCreate?: () => void;
  invalidate?: () => void;
}

export interface PivotTableDialogOpenOptions {
  placement?: 'new' | 'existing';
}

export interface PivotTableDialogHandle {
  open(opts?: PivotTableDialogOpenOptions): void;
  close(): void;
  setStrings(next: Strings): void;
  bindWorkbook(next: WorkbookHandle): void;
  detach(): void;
}

const colLetter = (n: number): string => {
  let v = n;
  let out = '';
  do {
    out = String.fromCharCode(65 + (v % 26)) + out;
    v = Math.floor(v / 26) - 1;
  } while (v >= 0);
  return out;
};

const rangeLabel = (range: { r0: number; c0: number; r1: number; c1: number }): string =>
  `${colLetter(range.c0)}${range.r0 + 1}:${colLetter(range.c1)}${range.r1 + 1}`;

const parseCellRef = (input: string): { row: number; col: number } | null => {
  const m = input
    .trim()
    .replace(/\$/g, '')
    .match(/^([A-Za-z]+)([1-9][0-9]*)$/);
  if (!m) return null;
  const letters = m[1];
  const rows = m[2];
  if (!letters || !rows) return null;
  let col = 0;
  for (const ch of letters.toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row: Number(rows) - 1, col: col - 1 };
};

export function attachPivotTableDialog(deps: PivotTableDialogDeps): PivotTableDialogHandle {
  const { host, store } = deps;
  let wb = deps.wb;
  let strings = deps.strings ?? defaultStrings;
  let open = false;
  let selectedFilterFields: string[] = [];
  let selectedValueFields: string[] = [];
  let selectedFilterItemVisibility = new Map<string, Map<string, boolean>>();
  let selectedFilterConditions = new Map<string, PivotFilterConditionState>();
  let selectedValueSettings = new Map<
    string,
    { aggregation?: PivotAggregation; numberFormat?: string; showValuesAs?: PivotShowValuesAs }
  >();
  let activeFieldSettings: PivotFieldSettingsActive | null = null;
  let draggedPivotField = '';

  const shell = createDialogShell({
    host,
    className: 'fc-pivotdlg',
    ariaLabel: strings.pivotTableDialog.title,
    onDismiss: () => close(),
  });
  shell.overlay.classList.add('fc-fmtdlg');
  const { overlay } = shell;
  const frame = appendDialogFrame(shell, {
    title: '',
    bodyTag: 'form',
    panelClasses: ['fc-fmtdlg__panel', 'fc-pivotdlg__panel'],
    bodyClass: 'fc-fmtdlg__body fc-pivotdlg__body',
  });
  const { header, footer } = frame;
  const body = frame.body as HTMLFormElement;

  const sourceInput = document.createElement('input');
  sourceInput.type = 'text';
  sourceInput.className = 'fc-namedlg__input';
  sourceInput.autocomplete = 'off';
  sourceInput.spellcheck = false;
  const tableRangeInput = document.createElement('input');
  tableRangeInput.type = 'radio';
  tableRangeInput.name = 'fc-pivotdlg-source-kind';
  tableRangeInput.value = 'range';
  tableRangeInput.checked = true;
  const externalSourceInput = document.createElement('input');
  externalSourceInput.type = 'radio';
  externalSourceInput.name = 'fc-pivotdlg-source-kind';
  externalSourceInput.value = 'external';
  projectDisabledState(
    externalSourceInput,
    true,
    strings.pivotTableDialog.externalSourceUnavailable,
    {
      describedById: 'fc-pivotdlg-external-unavailable',
      datasetKey: 'disabledReason',
    },
  );
  const nameInput = document.createElement('input');
  nameInput.type = 'text';
  nameInput.className = 'fc-namedlg__input';
  nameInput.autocomplete = 'off';
  nameInput.spellcheck = false;
  const destInput = document.createElement('input');
  destInput.type = 'text';
  destInput.className = 'fc-namedlg__input';
  destInput.autocomplete = 'off';
  destInput.spellcheck = false;
  const newWorksheetInput = document.createElement('input');
  newWorksheetInput.type = 'radio';
  newWorksheetInput.name = 'fc-pivotdlg-destination';
  newWorksheetInput.value = 'new';
  const existingWorksheetInput = document.createElement('input');
  existingWorksheetInput.type = 'radio';
  existingWorksheetInput.name = 'fc-pivotdlg-destination';
  existingWorksheetInput.value = 'existing';
  existingWorksheetInput.checked = true;
  const rowSelect = createDialogSelect([], '', { className: 'fc-fmtdlg__select' });
  const colSelect = createDialogSelect([], '', { className: 'fc-fmtdlg__select' });
  const filterSelect = createDialogSelect([], '', { className: 'fc-fmtdlg__select' });
  const valueSelect = createDialogSelect([], '', { className: 'fc-fmtdlg__select' });
  const aggSelect = createDialogSelect([], '', { className: 'fc-fmtdlg__select' });
  const rowSortSelect = createDialogSelect([], '', { className: 'fc-fmtdlg__select' });
  const colSortSelect = createDialogSelect([], '', { className: 'fc-fmtdlg__select' });
  const numberFormatInput = document.createElement('input');
  numberFormatInput.type = 'text';
  numberFormatInput.className = 'fc-namedlg__input';
  numberFormatInput.autocomplete = 'off';
  numberFormatInput.spellcheck = false;
  const rowSubtotalTop = document.createElement('input');
  rowSubtotalTop.type = 'checkbox';
  rowSubtotalTop.checked = true;
  const colSubtotalTop = document.createElement('input');
  colSubtotalTop.type = 'checkbox';
  colSubtotalTop.checked = true;
  const rowTotals = document.createElement('input');
  rowTotals.type = 'checkbox';
  rowTotals.checked = true;
  const colTotals = document.createElement('input');
  colTotals.type = 'checkbox';
  colTotals.checked = true;
  const fieldList = document.createElement('div');
  fieldList.className = 'fc-pivotdlg__field-list';
  const error = document.createElement('div');
  error.className = 'fc-namedlg__error';
  error.setAttribute('role', 'alert');
  error.hidden = true;

  const { cancelBtn, okBtn } = appendDialogActions(footer, {
    cancelLabel: '',
    okLabel: '',
  });

  const showError = (msg: string): void => {
    error.textContent = msg;
    error.hidden = false;
  };

  const setOkDisabled = (disabled: boolean, reason: string | null): void => {
    projectDisabledState(okBtn, disabled, reason, {
      datasetKey: 'disabledReason',
      titlePrefix: strings.pivotTableDialog.ok,
    });
  };

  const sheetIndexByName = (name: string): number => {
    const target = name.toLowerCase();
    for (let i = 0; i < wb.sheetCount; i += 1) {
      if (wb.sheetName(i).toLowerCase() === target) return i;
    }
    return -1;
  };

  const rangeFromSourceInput = () => {
    const parsed = parseRangeRef(sourceInput.value);
    if (!parsed) return null;
    const fallback = store.getState().selection.range.sheet;
    const sheet = parsed.sheetName == null ? fallback : sheetIndexByName(parsed.sheetName);
    if (sheet < 0) return null;
    return { sheet, r0: parsed.r0, c0: parsed.c0, r1: parsed.r1, c1: parsed.c1 };
  };

  const selectedRangeLabel = (): string => rangeLabel(store.getState().selection.range);

  const activeCellLabel = (): string => {
    const active = store.getState().selection.active;
    return `${colLetter(active.col)}${active.row + 1}`;
  };

  const fieldSelect = (select: HTMLSelectElement, fields: readonly PivotSourceField[]): void => {
    select.replaceChildren();
    appendDialogSelectOptions(
      select,
      fields.map((field) => ({ value: field.name, label: field.name })),
    );
  };

  const optionValues = (select: HTMLSelectElement): Set<string> =>
    new Set(Array.from(select.options).map((option) => option.value));

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

  const setFirstDifferent = (select: HTMLSelectElement, fieldName: string): void => {
    const next = Array.from(select.options).find(
      (option) => option.value && option.value !== fieldName,
    );
    select.value = next?.value ?? '';
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

  const flattenFilterItems = (): { fieldName: string; itemName: string; visible: boolean }[] =>
    selectedFilterFields.flatMap((fieldName) =>
      Array.from(selectedFilterItemVisibility.get(fieldName)?.entries() ?? []).map(
        ([itemName, visible]) => ({
          fieldName,
          itemName,
          visible,
        }),
      ),
    );

  const setFilterCondition = (fieldName: string, condition: PivotFilterConditionState): void => {
    const byField = new Map(selectedFilterConditions);
    if (condition.kind === 'none' || !condition.value.trim()) byField.delete(fieldName);
    else byField.set(fieldName, condition);
    selectedFilterConditions = byField;
  };

  const flattenPivotFilters = (): PivotFilterSpec[] =>
    selectedFilterFields.flatMap<PivotFilterSpec>((fieldName) => {
      const spec = pivotFilterConditionToSpec(fieldName, selectedFilterConditions.get(fieldName));
      return spec ? [spec] : [];
    });

  const removeFieldAssignment = (fieldName: string): void => {
    if (rowSelect.value === fieldName) rowSelect.value = '';
    if (colSelect.value === fieldName) colSelect.value = '';
    selectedFilterFields = selectedFilterFields.filter((name) => name !== fieldName);
    selectedValueFields = selectedValueFields.filter((name) => name !== fieldName);
    selectedValueSettings.delete(fieldName);
    selectedFilterItemVisibility.delete(fieldName);
    selectedFilterConditions.delete(fieldName);
  };

  const selectedValueSetting = (
    fieldName: string,
  ): { aggregation?: PivotAggregation; numberFormat?: string; showValuesAs?: PivotShowValuesAs } =>
    selectedValueSettings.get(fieldName) ?? {};

  const setValueFieldSetting = (
    fieldName: string,
    setting: {
      aggregation?: PivotAggregation;
      numberFormat?: string;
      showValuesAs?: PivotShowValuesAs;
    },
  ): void => {
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

  const flattenValueFieldSettings = (): {
    fieldName: string;
    aggregation?: PivotAggregation;
    numberFormat?: string;
    showValuesAs?: PivotShowValuesAs;
  }[] =>
    selectedValueFields.map((fieldName) => {
      const setting = selectedValueSetting(fieldName);
      return {
        fieldName,
        aggregation:
          setting.aggregation === undefined
            ? (Number(aggSelect.value) as PivotAggregation)
            : setting.aggregation,
        numberFormat: setting.numberFormat ?? numberFormatInput.value,
        showValuesAs: setting.showValuesAs,
      };
    });

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

  const pivotDragData = (event: DragEvent): string =>
    event.dataTransfer?.getData('text/plain') ||
    event.dataTransfer?.getData('application/x-fc-pivot-field') ||
    draggedPivotField;

  const setPivotDragData = (event: DragEvent, fieldName: string): void => {
    draggedPivotField = fieldName;
    event.dataTransfer?.setData('text/plain', fieldName);
    event.dataTransfer?.setData('application/x-fc-pivot-field', fieldName);
    if (event.dataTransfer) event.dataTransfer.effectAllowed = 'move';
  };

  const clearPivotDragData = (): void => {
    draggedPivotField = '';
  };

  const renderFieldSettingsPanel = (
    panelEl: HTMLDivElement,
    fields: readonly PivotSourceField[],
  ): void => {
    renderPivotFieldSettingsPanel({
      host,
      panelEl,
      active: activeFieldSettings,
      strings: strings.pivotTableDialog,
      okLabel: strings.pageSetup.ok,
      cancelLabel: strings.pageSetup.cancel,
      fields,
      controls: {
        rowSelect,
        colSelect,
        rowSortSelect,
        colSortSelect,
        aggSelect,
        numberFormatInput,
        rowSubtotalTop,
        colSubtotalTop,
      },
      selectedValueFields,
      selectedFilterFields,
      selectedValueSetting,
      setValueFieldSetting,
      fieldCanBeValue,
      replaceFilterField,
      normalizeSelectedFilters,
      refreshFieldList: () => updateFieldList(fields),
      filterItemVisibility,
      setFilterItemVisibility,
      selectedFilterCondition: (fieldName) => selectedFilterConditions.get(fieldName),
      setFilterCondition: (fieldName, condition) =>
        setFilterCondition(fieldName, {
          kind: condition.kind as PivotFilterConditionKind,
          value: condition.value,
        }),
      inferFilterItems: (fieldName) => {
        const range = rangeFromSourceInput();
        return range ? inferPivotFieldItems(wb, range, fieldName) : [];
      },
    });
  };

  const updateFieldList = (fields: readonly PivotSourceField[]): void => {
    const t = strings.pivotTableDialog;
    const assigned = new Set(
      [rowSelect.value, colSelect.value, ...selectedFilterFields, ...selectedValueFields].filter(
        Boolean,
      ),
    );
    fieldList.replaceChildren();

    const title = document.createElement('div');
    title.className = 'fc-pivotdlg__field-list-title';
    title.textContent = t.fieldList;
    const available = document.createElement('div');
    available.className = 'fc-pivotdlg__field-list-available';
    const availableLabel = document.createElement('div');
    availableLabel.className = 'fc-pivotdlg__field-list-label';
    availableLabel.textContent = t.availableFields;
    const fieldGrid = document.createElement('div');
    fieldGrid.className = 'fc-pivotdlg__field-list-grid';
    for (const field of fields) {
      const label = document.createElement('label');
      label.className = 'fc-pivotdlg__field-chip';
      const input = document.createElement('input');
      input.type = 'checkbox';
      input.checked = assigned.has(field.name);
      input.dataset.pivotFieldListField = field.name;
      input.addEventListener('change', () => {
        if (input.checked) {
          if (fieldCanBeValue(field.name)) addSelectedValue(field.name);
          else if (!rowSelect.value) rowSelect.value = field.name;
          else if (
            !colSelect.value &&
            rowSelect.value !== field.name &&
            valueSelect.value !== field.name
          )
            colSelect.value = field.name;
          else if (
            !filterSelect.value &&
            rowSelect.value !== field.name &&
            colSelect.value !== field.name &&
            valueSelect.value !== field.name
          )
            addSelectedFilter(field.name);
          else if (
            !fieldCanBeValue(field.name) &&
            rowSelect.value !== field.name &&
            colSelect.value !== field.name
          )
            addSelectedFilter(field.name);
        } else if (selectedFilterFields.includes(field.name)) {
          removeSelectedFilter(field.name);
        } else if (colSelect.value === field.name) {
          colSelect.value = '';
        } else if (rowSelect.value === field.name) {
          setFirstDifferent(rowSelect, field.name);
        } else if (selectedValueFields.includes(field.name)) {
          removeSelectedValue(field.name);
          if (selectedValueFields.length === 0) setFirstDifferent(valueSelect, field.name);
        }
        normalizeSelectedValues(fields);
        normalizeSelectedFilters();
        updateFieldList(fields);
      });
      const name = document.createElement('span');
      name.textContent = field.name;
      label.draggable = true;
      label.addEventListener('dragstart', (event) => setPivotDragData(event, field.name));
      label.addEventListener('dragend', clearPivotDragData);
      label.append(input, name);
      fieldGrid.appendChild(label);
    }
    available.append(availableLabel, fieldGrid);

    const areas = document.createElement('div');
    areas.className = 'fc-pivotdlg__areas';
    const areasLabel = document.createElement('div');
    areasLabel.className = 'fc-pivotdlg__field-list-label';
    areasLabel.textContent = t.fieldAreas;
    const areaGrid = document.createElement('div');
    areaGrid.className = 'fc-pivotdlg__area-grid';
    const settingsPanel = document.createElement('div');
    settingsPanel.className = 'fc-pivotdlg__area-settings-panel';
    settingsPanel.hidden = true;
    settingsPanel.setAttribute('role', 'status');
    settingsPanel.setAttribute('aria-live', 'polite');
    const showFieldSettings = (kind: PivotAreaKind, fieldName: string): void => {
      activeFieldSettings = { kind, fieldName };
      renderFieldSettingsPanel(settingsPanel, fields);
      settingsPanel.querySelector<HTMLElement>('select, input')?.focus();
    };
    const area = (
      label: string,
      values: readonly string[],
      kind: PivotAreaKind,
    ): HTMLDivElement => {
      const wrap = document.createElement('div');
      wrap.className = 'fc-pivotdlg__area';
      wrap.dataset.pivotArea = kind;
      wrap.addEventListener('dragover', (event) => {
        const fieldName = pivotDragData(event);
        if (!fieldName) return;
        if (kind === 'values' && !fieldCanBeValue(fieldName)) return;
        event.preventDefault();
        wrap.dataset.pivotDragOver = 'true';
      });
      wrap.addEventListener('dragleave', () => {
        delete wrap.dataset.pivotDragOver;
      });
      wrap.addEventListener('drop', (event) => {
        const fieldName = pivotDragData(event);
        if (!fieldName) return;
        event.preventDefault();
        delete wrap.dataset.pivotDragOver;
        assignFieldToArea(fieldName, kind, fields);
        updateFieldList(fields);
      });
      const heading = document.createElement('span');
      heading.textContent = label;
      const list = document.createElement('div');
      list.className = 'fc-pivotdlg__area-fields';
      const present = values.filter(Boolean);
      if (present.length === 0) {
        const none = document.createElement('strong');
        none.textContent = t.none;
        list.appendChild(none);
      } else {
        for (const value of present) {
          const chip = document.createElement('div');
          chip.className = 'fc-pivotdlg__area-field';
          chip.draggable = true;
          chip.addEventListener('dragstart', (event) => setPivotDragData(event, value));
          chip.addEventListener('dragend', clearPivotDragData);
          const name = document.createElement('strong');
          name.textContent = value;
          const settings = createPivotAreaSettingsButton(
            t.fieldSettings,
            t.fieldSettingsFor.replace('{field}', value),
          );
          settings.addEventListener('click', () => showFieldSettings(kind, value));
          chip.append(name, settings);
          list.appendChild(chip);
        }
      }
      wrap.append(heading, list);
      return wrap;
    };
    areaGrid.append(
      area(t.filtersArea, selectedFilterFields, 'filters'),
      area(t.columnField, [colSelect.value], 'columns'),
      area(t.rowField, [rowSelect.value], 'rows'),
      area(t.valueField, selectedValueFields, 'values'),
    );
    const activeIsPresent =
      activeFieldSettings &&
      (activeFieldSettings.kind === 'filters'
        ? selectedFilterFields.includes(activeFieldSettings.fieldName)
        : activeFieldSettings.kind === 'columns'
          ? colSelect.value === activeFieldSettings.fieldName
          : activeFieldSettings.kind === 'rows'
            ? rowSelect.value === activeFieldSettings.fieldName
            : selectedValueFields.includes(activeFieldSettings.fieldName));
    if (!activeIsPresent) activeFieldSettings = null;
    renderFieldSettingsPanel(settingsPanel, fields);
    areas.append(areasLabel, areaGrid, settingsPanel);
    fieldList.append(title, available, areas);
  };

  const labeled = (label: string, input: HTMLElement): HTMLLabelElement => {
    const row = document.createElement('label');
    row.className = 'fc-pivotdlg__field';
    const span = document.createElement('span');
    span.textContent = label;
    row.append(span, input);
    return row;
  };
  const checked = (label: string, input: HTMLInputElement): HTMLLabelElement => {
    const row = document.createElement('label');
    row.className = 'fc-pivotdlg__check';
    row.append(input, document.createTextNode(label));
    return row;
  };
  const sourceSelection = (): HTMLDivElement => {
    const wrap = document.createElement('div');
    wrap.className = 'fc-pivotdlg__source-choice';
    const legend = document.createElement('span');
    legend.className = 'fc-pivotdlg__placement-label';
    legend.textContent = strings.pivotTableDialog.sourceSection;
    const rangeChoice = checked(strings.pivotTableDialog.tableOrRange, tableRangeInput);
    const sourceField = labeled(strings.pivotTableDialog.source, sourceInput);
    sourceField.classList.add('fc-pivotdlg__source-field');
    const externalChoice = checked(strings.pivotTableDialog.externalSource, externalSourceInput);
    externalChoice.classList.add('fc-pivotdlg__check--disabled');
    const externalUnavailable = document.createElement('span');
    externalUnavailable.id = 'fc-pivotdlg-external-unavailable';
    externalUnavailable.className = 'fc-pivotdlg__disabled-note';
    externalUnavailable.textContent = strings.pivotTableDialog.externalSourceUnavailable;
    projectDisabledReason(externalChoice, strings.pivotTableDialog.externalSourceUnavailable, {
      ariaDescription: false,
    });
    wrap.append(legend, rangeChoice, sourceField, externalChoice, externalUnavailable);
    return wrap;
  };
  const destinationPlacement = (): HTMLDivElement => {
    const wrap = document.createElement('div');
    wrap.className = 'fc-pivotdlg__placement';
    const legend = document.createElement('span');
    legend.className = 'fc-pivotdlg__placement-label';
    legend.textContent = strings.pivotTableDialog.destinationSection;
    wrap.append(
      legend,
      checked(strings.pivotTableDialog.newWorksheet, newWorksheetInput),
      checked(strings.pivotTableDialog.existingWorksheet, existingWorksheetInput),
      labeled(strings.pivotTableDialog.destination, destInput),
    );
    return wrap;
  };
  const section = (...children: HTMLElement[]): HTMLDivElement => {
    const el = document.createElement('div');
    el.className = 'fc-pivotdlg__section';
    el.append(...children);
    return el;
  };
  const checkGrid = (...children: HTMLLabelElement[]): HTMLDivElement => {
    const el = document.createElement('div');
    el.className = 'fc-pivotdlg__checkgrid';
    el.append(...children);
    return el;
  };

  const configureForSource = (): void => {
    const t = strings.pivotTableDialog;
    error.hidden = true;
    error.textContent = '';
    const range = rangeFromSourceInput();
    if (!range) {
      showError(t.invalidRange);
      setOkDisabled(true, t.invalidRange);
      return;
    }
    const fields = inferPivotSourceFields(wb, range);
    const numeric = fields.filter((f) => f.numericCount > 0);
    if (!wb.capabilities.pivotTableMutate) {
      showError(t.unsupported);
      setOkDisabled(true, t.unsupported);
      return;
    }
    if (fields.length < 2) {
      showError(t.invalidRange);
      setOkDisabled(true, t.invalidRange);
      return;
    }
    const prevRow = rowSelect.value;
    const prevCol = colSelect.value;
    const prevFilters =
      selectedFilterFields.length > 0 ? selectedFilterFields : [filterSelect.value];
    const prevValues = selectedValueFields.length > 0 ? selectedValueFields : [valueSelect.value];
    setOkDisabled(false, null);
    fieldSelect(rowSelect, fields);
    fieldSelect(colSelect, fields);
    fieldSelect(filterSelect, fields);
    fieldSelect(valueSelect, numeric.length > 0 ? numeric : fields);
    for (const select of [colSelect, filterSelect]) {
      const currentOptions = Array.from(select.options).map((option) => ({
        value: option.value,
        label: option.textContent ?? '',
      }));
      select.replaceChildren();
      appendDialogSelectOptions(select, [{ value: '', label: t.none }, ...currentOptions]);
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
    selectedValueSettings = new Map(
      Array.from(selectedValueSettings.entries()).filter(([fieldName]) =>
        selectedValueFields.includes(fieldName),
      ),
    );
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
    selectedFilterItemVisibility = new Map(
      Array.from(selectedFilterItemVisibility.entries()).filter(([fieldName]) =>
        selectedFilterFields.includes(fieldName),
      ),
    );
    selectedFilterConditions = new Map(
      Array.from(selectedFilterConditions.entries()).filter(([fieldName]) =>
        selectedFilterFields.includes(fieldName),
      ),
    );
    filterSelect.value = selectedFilterFields[0] ?? '';
    normalizeSelectedFilters();
    updateFieldList(fields);
  };

  const render = (): void => {
    const t = strings.pivotTableDialog;
    header.textContent = t.title;
    shell.setAriaLabel(t.title);
    cancelBtn.textContent = t.cancel;
    okBtn.textContent = t.ok;
    sourceInput.placeholder = t.sourcePlaceholder;
    nameInput.placeholder = t.namePlaceholder;
    destInput.placeholder = t.destinationPlaceholder;
    numberFormatInput.placeholder = t.numberFormatPlaceholder;
    error.hidden = true;
    error.textContent = '';

    const range = store.getState().selection.range;
    body.replaceChildren();

    sourceInput.value = sourceInput.value || rangeLabel(range);
    nameInput.value = nameInput.value || `PivotTable${wb.getPivotTables().length + 1}`;
    const dest = `${colLetter(range.c0)}${range.r1 + 3}`;
    destInput.value = destInput.value || dest;
    aggSelect.replaceChildren();
    appendDialogSelectOptions(aggSelect, [
      { value: String(PivotAggregation.Sum), label: t.sum },
      { value: String(PivotAggregation.Count), label: t.count },
      { value: String(PivotAggregation.Average), label: t.average },
      { value: String(PivotAggregation.Max), label: t.max },
      { value: String(PivotAggregation.Min), label: t.min },
    ]);
    rowSortSelect.replaceChildren();
    colSortSelect.replaceChildren();
    for (const select of [rowSortSelect, colSortSelect]) {
      appendDialogSelectOptions(select, [
        { value: 'none', label: t.sortNone },
        { value: 'asc', label: t.sortAsc },
        { value: 'desc', label: t.sortDesc },
      ]);
    }

    body.append(
      section(sourceSelection(), labeled(t.name, nameInput)),
      section(destinationPlacement()),
      fieldList,
      section(
        labeled(t.filtersArea, filterSelect),
        labeled(t.rowField, rowSelect),
        labeled(t.columnField, colSelect),
        labeled(t.valueField, valueSelect),
        labeled(t.aggregation, aggSelect),
      ),
      section(
        labeled(t.rowSort, rowSortSelect),
        labeled(t.columnSort, colSortSelect),
        labeled(t.numberFormat, numberFormatInput),
      ),
      checkGrid(
        checked(t.rowSubtotalTop, rowSubtotalTop),
        checked(t.columnSubtotalTop, colSubtotalTop),
        checked(t.rowGrandTotals, rowTotals),
        checked(t.columnGrandTotals, colTotals),
      ),
      error,
    );
    attachRangePickerButton(sourceInput, {
      label: t.rangePickerSelect,
      getValue: selectedRangeLabel,
      subscribeToRangeChanges: (listener) => store.subscribe(listener),
      kind: 'pivot-source',
    });
    attachRangePickerButton(destInput, {
      label: t.rangePickerSelect,
      getValue: activeCellLabel,
      subscribeToRangeChanges: (listener) => store.subscribe(listener),
      kind: 'pivot-destination',
    });
    updateDestinationPlacementState();
    configureForSource();
  };

  const updateDestinationPlacementState = (): void => {
    const existing = existingWorksheetInput.checked;
    const reason = existing ? null : strings.pivotTableDialog.destinationRequiresExistingWorksheet;
    projectDisabledState(destInput, !existing, reason, { datasetKey: 'disabledReason' });
    const destinationPicker = destInput
      .closest('.fc-range-picker')
      ?.querySelector<HTMLButtonElement>('.fc-range-picker__btn');
    destinationPicker?.toggleAttribute('disabled', !existing);
    if (destinationPicker) {
      projectDisabledReason(destinationPicker, reason, {
        datasetKey: 'disabledReason',
        titlePrefix: strings.pivotTableDialog.rangePickerSelect,
      });
    }
  };

  const close = (): void => {
    open = false;
    shell.close();
  };

  const onSubmit = (e: SubmitEvent): void => {
    e.preventDefault();
    const range = rangeFromSourceInput();
    if (!range) {
      showError(strings.pivotTableDialog.invalidRange);
      sourceInput.focus();
      return;
    }
    const useNewWorksheet = newWorksheetInput.checked;
    const dest = useNewWorksheet ? { row: 0, col: 0 } : parseCellRef(destInput.value);
    if (!dest) {
      showError(strings.pivotTableDialog.invalidDestination);
      destInput.focus();
      return;
    }
    let destinationSheet = range.sheet;
    if (useNewWorksheet) {
      const added = wb.addSheet();
      if (added < 0) {
        showError(strings.pivotTableDialog.engineFailed);
        return;
      }
      destinationSheet = added;
    }
    const result = createPivotTableFromRange(wb, {
      source: range,
      destination: { sheet: destinationSheet, row: dest.row, col: dest.col },
      name: nameInput.value,
      rowField: rowSelect.value,
      columnField: colSelect.value || undefined,
      filterField: filterSelect.value || undefined,
      filterFields: selectedFilterFields,
      filterItems: flattenFilterItems(),
      pivotFilters: flattenPivotFilters(),
      valueField: valueSelect.value,
      valueFields: selectedValueFields,
      valueFieldSettings: flattenValueFieldSettings(),
      aggregation: Number(aggSelect.value) as PivotAggregation,
      rowSort: rowSortSelect.value as 'none' | 'asc' | 'desc',
      columnSort: colSortSelect.value as 'none' | 'asc' | 'desc',
      rowSubtotalTop: rowSubtotalTop.checked,
      columnSubtotalTop: colSubtotalTop.checked,
      valueNumberFormat: numberFormatInput.value,
      showRowGrandTotals: rowTotals.checked,
      showColumnGrandTotals: colTotals.checked,
    });
    if (!result.ok) {
      showError(strings.pivotTableDialog.engineFailed);
      return;
    }
    if (destinationSheet !== store.getState().data.sheetIndex) {
      mutators.replaceCells(store, wb.cells(destinationSheet));
      mutators.setSheetIndex(store, destinationSheet);
    }
    mutators.setActive(store, { sheet: destinationSheet, row: dest.row, col: dest.col });
    deps.onAfterCreate?.();
    deps.invalidate?.();
    close();
  };

  const onKey = (e: KeyboardEvent): void => {
    e.stopPropagation();
    if (e.key === 'Escape') {
      e.preventDefault();
      close();
    } else if (e.key === 'Enter' && !okBtn.disabled) {
      e.preventDefault();
      body.requestSubmit();
    }
  };
  const onOk = (): void => body.requestSubmit();

  shell.on(body, 'submit', onSubmit as EventListener);
  shell.on(sourceInput, 'input', configureForSource as EventListener);
  const onValueSelectChange = (): void => {
    selectedValueFields = valueSelect.value ? [valueSelect.value] : [];
    selectedValueSettings = new Map(
      Array.from(selectedValueSettings.entries()).filter(([fieldName]) =>
        selectedValueFields.includes(fieldName),
      ),
    );
    configureForSource();
  };

  const onFilterSelectChange = (): void => {
    selectedFilterFields = filterSelect.value ? [filterSelect.value] : [];
    configureForSource();
  };

  for (const select of [rowSelect, colSelect]) {
    shell.on(select, 'change', configureForSource as EventListener);
  }
  shell.on(filterSelect, 'change', onFilterSelectChange as EventListener);
  shell.on(valueSelect, 'change', onValueSelectChange as EventListener);
  shell.on(newWorksheetInput, 'change', updateDestinationPlacementState as EventListener);
  shell.on(existingWorksheetInput, 'change', updateDestinationPlacementState as EventListener);
  shell.on(okBtn, 'click', onOk);
  shell.on(cancelBtn, 'click', close);
  shell.on(overlay, 'keydown', onKey as EventListener);

  return {
    open(opts = {}) {
      sourceInput.value = rangeLabel(store.getState().selection.range);
      if (opts.placement === 'new') {
        newWorksheetInput.checked = true;
        existingWorksheetInput.checked = false;
      } else {
        newWorksheetInput.checked = false;
        existingWorksheetInput.checked = true;
      }
      render();
      shell.open();
      open = true;
      const initial =
        wb.capabilities.pivotTableMutate && sourceInput.isConnected ? sourceInput : cancelBtn;
      initial.focus({ preventScroll: true });
      if (initial === sourceInput) sourceInput.select();
    },
    close,
    setStrings(next) {
      strings = next;
      if (open) render();
    },
    bindWorkbook(next) {
      wb = next;
      if (open) render();
    },
    detach() {
      shell.dispose();
    },
  };
}
