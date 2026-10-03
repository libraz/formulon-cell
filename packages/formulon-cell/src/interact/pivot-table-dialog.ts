import {
  createPivotTableFromRange,
  inferPivotFieldItems,
  inferPivotSourceFields,
  type PivotSourceField,
} from '../commands/pivot-table.js';
import { formatA1Cell, parseA1Atom } from '../engine/address.js';
import { parseRangeRef } from '../engine/range-resolver.js';
import { PivotAggregation } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { appendDialogSelectOptions, createDialogSelect } from '../toolbar/dialogs/form-controls.js';
import { projectDisabledReason, projectDisabledState } from '../toolbar/menu-a11y.js';
import { formatSheetAbsoluteRange } from '../wrappers/toolbar-a1.js';
import type { SheetRange } from '../wrappers/toolbar-types.js';
import { appendDialogActions, appendDialogFrame, createDialogShell } from './dialog-shell.js';
import { createPivotFieldAreas } from './pivot-table-field-areas.js';
import { renderPivotFieldList } from './pivot-table-field-list.js';
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

export function attachPivotTableDialog(deps: PivotTableDialogDeps): PivotTableDialogHandle {
  const { host, store } = deps;
  let wb = deps.wb;
  let strings = deps.strings ?? defaultStrings;
  let open = false;
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

  /** Table/Range mirrors the desktop dialog: sheet-qualified and absolute. */
  const sourceRangeLabel = (range: SheetRange): string =>
    formatSheetAbsoluteRange(wb.sheetName(range.sheet), range);

  const selectedRangeLabel = (): string => sourceRangeLabel(store.getState().selection.range);

  const activeCellLabel = (): string => {
    const active = store.getState().selection.active;
    return formatA1Cell(active.row, active.col);
  };

  const areas = createPivotFieldAreas({ rowSelect, colSelect, filterSelect, valueSelect });

  const updateFieldList = (fields: readonly PivotSourceField[]): void =>
    renderPivotFieldList(fieldList, {
      host,
      get strings() {
        return strings;
      },
      areas,
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
      fields,
      inferFilterItems: (fieldName) => {
        const range = rangeFromSourceInput();
        return range ? inferPivotFieldItems(wb, range, fieldName) : [];
      },
    });

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
    setOkDisabled(false, null);
    areas.resetSource(fields, t.none);
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

    sourceInput.value = sourceInput.value || sourceRangeLabel(range);
    nameInput.value = nameInput.value || `PivotTable${wb.getPivotTables().length + 1}`;
    const dest = formatA1Cell(range.r1 + 2, range.c0);
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
    const dest = useNewWorksheet ? { row: 0, col: 0 } : parseA1Atom(destInput.value);
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
    const aggregation = Number(aggSelect.value) as PivotAggregation;
    const { filterItems, pivotFilters, valueFieldSettings } = areas.toCreateSpec({
      aggregation,
      numberFormat: numberFormatInput.value,
    });
    const result = createPivotTableFromRange(wb, {
      source: range,
      destination: { sheet: destinationSheet, row: dest.row, col: dest.col },
      name: nameInput.value,
      rowField: rowSelect.value,
      columnField: colSelect.value || undefined,
      filterField: filterSelect.value || undefined,
      filterFields: areas.filterFields,
      filterItems,
      pivotFilters,
      valueField: valueSelect.value,
      valueFields: areas.valueFields,
      valueFieldSettings,
      aggregation,
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
    areas.selectSingleValue();
    configureForSource();
  };

  const onFilterSelectChange = (): void => {
    areas.selectSingleFilter();
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
      sourceInput.value = selectedRangeLabel();
      // New worksheet is the dialog's default placement; only an explicit
      // "existing sheet" entry point starts on the other radio.
      const placement = opts.placement ?? 'new';
      newWorksheetInput.checked = placement === 'new';
      existingWorksheetInput.checked = placement === 'existing';
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
