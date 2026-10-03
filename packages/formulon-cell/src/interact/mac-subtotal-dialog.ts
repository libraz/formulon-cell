import {
  commitMacSubtotal,
  describeMacDataError,
  formatMacRangeAddress,
  type MacDataFunction,
  parseMacRange,
  planMacSubtotal,
  type SubtotalRequest,
} from '../commands/mac-data-tools.js';
import { colFromLetters, colLetter } from '../engine/address.js';
import type { SpreadsheetInstance } from '../mount/types.js';
import { appendDialogSelectOptions } from '../toolbar/dialogs/form-controls.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import { syncCustomSelects } from './custom-select.js';
import { appendDialogActions, appendDialogFrame, createDialogShell } from './dialog-shell.js';
import { isSubmitEnter } from './mac-dialog-keys.js';

interface SubtotalDialogHandle {
  open(): void;
  close(): void;
  detach(): void;
}

const handles = new WeakMap<SpreadsheetInstance, SubtotalDialogHandle>();
const FUNCTIONS: readonly MacDataFunction[] = ['sum', 'average', 'count', 'min', 'max'];

interface Field {
  readonly input: HTMLInputElement;
  readonly text: HTMLSpanElement;
}

const appendField = (body: HTMLElement, id: string): Field => {
  const row = document.createElement('label');
  row.className = 'fc-fmtdlg__row fc-mac-subtotal__row';
  const text = document.createElement('span');
  const input = document.createElement('input');
  input.id = id;
  input.type = 'text';
  input.className = 'fc-fmtdlg__input';
  input.autocomplete = 'off';
  input.spellcheck = false;
  row.append(text, input);
  body.appendChild(row);
  return { input, text };
};

/** A 1-based column number or A1 column letters, as a 0-based index. */
const columnIndex = (token: string): number | null => {
  const trimmed = token.trim();
  if (!trimmed) return null;
  if (/^\d+$/.test(trimmed)) {
    const n = Number(trimmed);
    return Number.isInteger(n) && n > 0 ? n - 1 : null;
  }
  return /^[A-Za-z]+$/.test(trimmed) ? colFromLetters(trimmed) : null;
};

function attachSubtotalDialog(instance: SpreadsheetInstance): SubtotalDialogHandle {
  let strings = instance.i18n.strings.macData;
  let t = strings.subtotal;
  const shell = createDialogShell({
    host: instance.host,
    className: 'fc-mac-subtotal',
    ariaLabel: t.title,
    onDismiss: () => api.close(),
  });
  shell.overlay.classList.add('fc-fmtdlg');
  const { header, body, footer } = appendDialogFrame(shell, {
    title: t.title,
    panelClasses: ['fc-fmtdlg__panel', 'fc-mac-subtotal__panel'],
    bodyClass: 'fc-fmtdlg__body fc-mac-subtotal__body',
    footerClass: 'fc-fmtdlg__footer fc-mac-subtotal__footer',
  });
  const rangeField = appendField(body, 'fc-mac-subtotal-range');
  const groupByField = appendField(body, 'fc-mac-subtotal-group-by');
  const columnsField = appendField(body, 'fc-mac-subtotal-columns');
  const rangeInput = rangeField.input;
  const groupByInput = groupByField.input;
  const columnsInput = columnsField.input;

  const functionRow = document.createElement('label');
  functionRow.className = 'fc-fmtdlg__row fc-mac-subtotal__row';
  const functionLabel = document.createElement('span');
  const functionSelect = document.createElement('select');
  functionSelect.id = 'fc-mac-subtotal-function';
  functionSelect.className = 'fc-fmtdlg__select';
  appendDialogSelectOptions(
    functionSelect,
    FUNCTIONS.map((fn) => ({ value: fn, label: '' })),
  );
  const functionOptions = new Map<MacDataFunction, HTMLOptionElement>(
    FUNCTIONS.map((fn, index) => [fn, functionSelect.options[index] as HTMLOptionElement]),
  );
  functionRow.append(functionLabel, functionSelect);
  body.appendChild(functionRow);

  const replaceRow = document.createElement('label');
  replaceRow.className = 'fc-mac-subtotal__option';
  const replaceInput = document.createElement('input');
  replaceInput.type = 'checkbox';
  projectDisabledState(replaceInput, true, t.unsupportedOption, { datasetKey: 'disabledReason' });
  const replaceText = document.createElement('span');
  replaceRow.append(replaceInput, replaceText);
  body.appendChild(replaceRow);

  const summaryRow = document.createElement('label');
  summaryRow.className = 'fc-mac-subtotal__option';
  const summaryInput = document.createElement('input');
  summaryInput.type = 'checkbox';
  summaryInput.checked = true;
  projectDisabledState(summaryInput, true, t.unsupportedOption, { datasetKey: 'disabledReason' });
  const summaryText = document.createElement('span');
  summaryRow.append(summaryInput, summaryText);
  body.appendChild(summaryRow);

  const status = document.createElement('div');
  status.className = 'fc-mac-subtotal__status';
  status.setAttribute('role', 'status');
  status.setAttribute('aria-live', 'polite');
  body.appendChild(status);

  const { cancelBtn, okBtn } = appendDialogActions(footer, {
    cancelLabel: t.cancel,
    okLabel: t.run,
    buttonBaseClass: 'fc-fmtdlg__btn fc-mac-subtotal__btn',
  });
  okBtn.dataset.fcMacAction = 'subtotal-ok';
  cancelBtn.dataset.fcMacAction = 'subtotal-cancel';

  const applyLabels = (): void => {
    strings = instance.i18n.strings.macData;
    t = strings.subtotal;
    shell.setAriaLabel(t.title);
    header.textContent = t.title;
    rangeField.text.textContent = t.range;
    groupByField.text.textContent = t.groupBy;
    columnsField.text.textContent = t.columns;
    functionLabel.textContent = t.function;
    for (const [fn, option] of functionOptions) option.textContent = strings.functions[fn];
    replaceText.textContent = t.replace;
    summaryText.textContent = t.summaryBelow;
    for (const input of [replaceInput, summaryInput]) {
      projectDisabledState(input, true, t.unsupportedOption, { datasetKey: 'disabledReason' });
    }
    okBtn.textContent = t.run;
    cancelBtn.textContent = t.cancel;
    syncCustomSelects(shell.panel);
  };
  applyLabels();
  const unsubscribeLocale = instance.i18n.subscribe(applyLabels);

  const close = (): void => {
    shell.close();
    instance.host.focus();
  };

  const defaults = (): void => {
    const selection = instance.store.getState().selection.range;
    rangeInput.value = formatMacRangeAddress(instance.workbook, selection);
    groupByInput.value = colLetter(selection.c0);
    columnsInput.value = Array.from({ length: Math.max(0, selection.c1 - selection.c0) }, (_, i) =>
      colLetter(selection.c0 + i + 1),
    ).join(',');
    functionSelect.value = 'sum';
  };

  const run = (): void => {
    delete status.dataset.state;
    status.textContent = '';
    const parsed = parseMacRange(
      instance.workbook,
      rangeInput.value,
      instance.store.getState().data.sheetIndex,
    );
    if (!parsed) {
      status.dataset.state = 'error';
      status.textContent = strings.errors.invalidRange;
      return;
    }
    const groupBy = columnIndex(groupByInput.value);
    const columns = columnsInput.value
      .split(/[\s,;]+/)
      .map(columnIndex)
      .filter((value): value is number => value !== null);
    if (groupBy === null || columns.length === 0) {
      status.dataset.state = 'error';
      status.textContent = strings.errors.invalidColumn;
      return;
    }
    const request: SubtotalRequest = {
      range: rangeInput.value,
      groupByColumn: groupBy,
      subtotalColumns: columns,
      function: (FUNCTIONS.includes(functionSelect.value as MacDataFunction)
        ? functionSelect.value
        : 'sum') as MacDataFunction,
    };
    const plan = planMacSubtotal(instance, request);
    if (!plan.ok) {
      status.dataset.state = 'error';
      status.textContent = describeMacDataError(instance.i18n.strings, plan.error);
      return;
    }
    const committed = commitMacSubtotal(instance, request, plan.value);
    if (!committed.ok) {
      status.dataset.state = 'error';
      status.textContent = describeMacDataError(instance.i18n.strings, committed.error);
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

  const api: SubtotalDialogHandle = {
    open(): void {
      applyLabels();
      defaults();
      delete status.dataset.state;
      status.textContent = '';
      shell.open();
      queueMicrotask(() => rangeInput.focus());
    },
    close,
    detach(): void {
      unsubscribeLocale();
      shell.dispose();
    },
  };
  return api;
}

/** Open the Data → Outline → Subtotal dialog. */
export function openMacSubtotalDialog(instance: SpreadsheetInstance): void {
  let handle = handles.get(instance);
  if (!handle) {
    handle = attachSubtotalDialog(instance);
    handles.set(instance, handle);
  }
  handle.open();
}

/** Dispose the Subtotal overlay owned by an instance, if it has been opened. */
export function disposeMacSubtotalDialog(instance: SpreadsheetInstance): void {
  const handle = handles.get(instance);
  if (!handle) return;
  handles.delete(instance);
  handle.detach();
}
