import { canExecuteBuiltIn } from '../commands/built-in-command-policy.js';
import {
  createDataTable,
  DATA_TABLE_COMMAND,
  type DataTableRequestError,
} from '../commands/data-table.js';
import type { InteractionOrigin } from '../commands/interaction-policy.js';
import { parseMacRange } from '../commands/mac-data-tools.js';
import type { Addr, Range } from '../engine/types.js';
import type { SpreadsheetInstance } from '../mount/types.js';
import {
  appendDialogActions,
  appendDialogFrame,
  createDialogShell,
  createDialogStatus,
  deferInitialFocus,
} from './dialog-shell.js';
import { isSubmitEnter } from './mac-dialog-keys.js';
import { createRangePickFormatter } from './range-pick-format.js';
import { attachRangePickerButton, updateRangePickerLabel } from './range-picker-control.js';

interface DataTableDialogHandle {
  open(origin: InteractionOrigin): void;
  close(): void;
  detach(): void;
}

const handles = new WeakMap<SpreadsheetInstance, DataTableDialogHandle>();

function attachDataTableDialog(instance: SpreadsheetInstance): DataTableDialogHandle {
  let t = instance.i18n.strings.dataTableDialog;
  const shell = createDialogShell({
    host: instance.host,
    className: 'fc-datatable',
    ariaLabel: t.title,
    onDismiss: () => close(),
  });
  shell.overlay.classList.add('fc-fmtdlg');
  const { header, body, footer } = appendDialogFrame(shell, {
    title: t.title,
    panelClasses: ['fc-fmtdlg__panel', 'fc-datatable__panel'],
    bodyClass: 'fc-fmtdlg__body fc-datatable__body',
    footerClass: 'fc-fmtdlg__footer',
  });
  const field = (name: 'row' | 'col'): { input: HTMLInputElement; text: HTMLSpanElement } => {
    const row = document.createElement('label');
    row.className = 'fc-fmtdlg__row';
    const text = document.createElement('span');
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'fc-fmtdlg__input';
    input.autocomplete = 'off';
    input.spellcheck = false;
    input.dataset.dataTableInput = name;
    row.append(text, input);
    body.appendChild(row);
    return { input, text };
  };
  const rowField = field('row');
  const colField = field('col');
  const pickFormat = createRangePickFormatter({
    store: instance.store,
    sheetName: (index) => instance.workbook.sheetName(index),
  });
  const pickers = [rowField, colField].map(({ input }) =>
    attachRangePickerButton(input, {
      label: instance.i18n.strings.pivotTableDialog.rangePickerSelect,
      getValue: () => pickFormat.cell(),
      subscribeToRangeChanges: pickFormat.subscribe,
      kind: `data-table-${input.dataset.dataTableInput}`,
    }),
  );
  // Picking an input cell moves the grid selection, so the table range is
  // the selection captured when the dialog opened.
  let tableRange: Range = instance.store.getState().selection.range;
  let origin: InteractionOrigin = 'instanceApi';
  const status = createDialogStatus(body, 'fc-datatable__status');
  const { cancelBtn, okBtn } = appendDialogActions(footer, {
    cancelLabel: t.cancel,
    okLabel: t.ok,
    buttonBaseClass: 'fc-fmtdlg__btn',
  });

  const applyLabels = (): void => {
    t = instance.i18n.strings.dataTableDialog;
    shell.setAriaLabel(t.title);
    header.textContent = t.title;
    rowField.text.textContent = t.rowInput;
    colField.text.textContent = t.colInput;
    okBtn.textContent = t.ok;
    cancelBtn.textContent = t.cancel;
    for (const picker of pickers) {
      updateRangePickerLabel(picker, instance.i18n.strings.pivotTableDialog.rangePickerSelect);
    }
  };
  applyLabels();
  const unsubscribeLocale = instance.i18n.on('change', applyLabels);

  const close = (): void => {
    shell.close();
  };

  /** A blank field is "no input"; anything else must name one cell. */
  const inputCell = (raw: string): Addr | null | 'invalid' => {
    if (raw.trim() === '') return null;
    const parsed = parseMacRange(instance.workbook, raw, tableRange.sheet);
    if (!parsed) return 'invalid';
    const { range } = parsed;
    if (range.r0 !== range.r1 || range.c0 !== range.c1) return 'invalid';
    return { sheet: range.sheet, row: range.r0, col: range.c0 };
  };

  const messages: Record<DataTableRequestError, () => string> = {
    'invalid-range': () => t.invalidRange,
    'invalid-input': () => t.invalidInput,
    rejected: () => t.rejected,
  };

  const showError = (message: string): void => {
    status.error(message);
  };

  const run = (): void => {
    const rowInput = inputCell(rowField.input.value);
    const colInput = inputCell(colField.input.value);
    if (rowInput === 'invalid' || colInput === 'invalid') {
      showError(t.invalidInput);
      return;
    }
    const error = createDataTable(
      { store: instance.store, workbook: instance.workbook, history: instance.history, origin },
      { selection: tableRange, rowInput, colInput },
    );
    if (error) {
      showError(messages[error]());
      return;
    }
    close();
  };

  shell.on(okBtn, 'click', run);
  shell.on(cancelBtn, 'click', close);
  shell.on(shell.overlay, 'keydown', (event) => {
    const e = event as KeyboardEvent;
    if (!isSubmitEnter(e)) return;
    e.preventDefault();
    run();
  });

  return {
    open(from: InteractionOrigin): void {
      applyLabels();
      origin = from;
      tableRange = instance.store.getState().selection.range;
      pickFormat.begin();
      rowField.input.value = '';
      colField.input.value = '';
      status.clear();
      shell.open();
      deferInitialFocus(shell.overlay, () => rowField.input.focus());
    },
    close,
    detach(): void {
      unsubscribeLocale();
      shell.dispose();
    },
  };
}

/** Open Data → What-If Analysis → Data Table over the current selection from `origin`; a refused
 *  command opens nothing. */
export function openDataTableDialogFrom(
  instance: SpreadsheetInstance,
  origin: InteractionOrigin,
): void {
  if (!canExecuteBuiltIn(instance.store, DATA_TABLE_COMMAND, origin).allowed) return;
  let handle = handles.get(instance);
  if (!handle) {
    handle = attachDataTableDialog(instance);
    handles.set(instance, handle);
  }
  handle.open(origin);
}

/** Open Data → What-If Analysis → Data Table over the current selection. */
export function openDataTableDialog(instance: SpreadsheetInstance): void {
  openDataTableDialogFrom(instance, 'instanceApi');
}

/** Dispose the Data Table overlay owned by an instance, if it has been opened. */
export function disposeDataTableDialog(instance: SpreadsheetInstance): void {
  const handle = handles.get(instance);
  if (!handle) return;
  handles.delete(instance);
  handle.detach();
}
