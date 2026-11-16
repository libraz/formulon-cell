import { inferAutoFilterRange } from '../commands/filter.js';
import {
  commitMacSubtotal,
  describeMacDataError,
  formatMacRangeAddress,
  type MacDataFunction,
  planMacSubtotal,
  removeMacSubtotals,
  replaceMacSubtotals,
  type SubtotalRequest,
} from '../commands/mac-data-tools.js';
import { colLetter } from '../engine/address.js';
import type { Range } from '../engine/types.js';
import type { SpreadsheetInstance } from '../mount/types.js';
import { appendDialogSelectOptions } from '../toolbar/dialogs/form-controls.js';
import { syncCustomSelects } from './custom-select.js';
import {
  appendDialogActions,
  appendDialogButton,
  appendDialogFrame,
  createDialogShell,
} from './dialog-shell.js';
import { isSubmitEnter } from './mac-dialog-keys.js';

interface SubtotalDialogHandle {
  open(): void;
  close(): void;
  detach(): void;
}

const handles = new WeakMap<SpreadsheetInstance, SubtotalDialogHandle>();
const FUNCTIONS: readonly MacDataFunction[] = ['sum', 'average', 'count', 'min', 'max'];

const appendRow = (body: HTMLElement): { row: HTMLLabelElement; text: HTMLSpanElement } => {
  const row = document.createElement('label');
  row.className = 'fc-fmtdlg__row fc-mac-subtotal__row';
  const text = document.createElement('span');
  row.appendChild(text);
  body.appendChild(row);
  return { row, text };
};

const appendSelect = (row: HTMLElement, id: string): HTMLSelectElement => {
  const select = document.createElement('select');
  select.id = id;
  select.className = 'fc-fmtdlg__select';
  row.appendChild(select);
  return select;
};

const appendOption = (
  body: HTMLElement,
): { row: HTMLLabelElement; input: HTMLInputElement; text: HTMLSpanElement } => {
  const row = document.createElement('label');
  row.className = 'fc-mac-subtotal__option';
  const input = document.createElement('input');
  input.type = 'checkbox';
  const text = document.createElement('span');
  row.append(input, text);
  body.appendChild(row);
  return { row, input, text };
};

function attachSubtotalDialog(instance: SpreadsheetInstance): SubtotalDialogHandle {
  let strings = instance.i18n.strings.mac.data;
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

  // The list is the current region; its header row names the columns.
  const groupByRow = appendRow(body);
  const groupBySelect = appendSelect(groupByRow.row, 'fc-mac-subtotal-group-by');
  const functionRow = appendRow(body);
  const functionSelect = appendSelect(functionRow.row, 'fc-mac-subtotal-function');
  appendDialogSelectOptions(
    functionSelect,
    FUNCTIONS.map((fn) => ({ value: fn, label: '' })),
  );
  const functionOptions = new Map<MacDataFunction, HTMLOptionElement>(
    FUNCTIONS.map((fn, index) => [fn, functionSelect.options[index] as HTMLOptionElement]),
  );
  const fieldsLabel = document.createElement('div');
  fieldsLabel.className = 'fc-mac-subtotal__fields-label';
  const fields = document.createElement('div');
  fields.className = 'fc-mac-subtotal__fields';
  fields.setAttribute('role', 'group');
  body.append(fieldsLabel, fields);

  const replaceOption = appendOption(body);
  replaceOption.input.checked = true;
  const pageBreakOption = appendOption(body);
  const summaryOption = appendOption(body);
  summaryOption.input.checked = true;

  const status = document.createElement('div');
  status.className = 'fc-mac-subtotal__status';
  status.setAttribute('role', 'status');
  status.setAttribute('aria-live', 'polite');
  body.appendChild(status);

  const removeAllBtn = appendDialogButton(footer, {
    label: '',
    baseClass: 'fc-fmtdlg__btn fc-mac-subtotal__btn',
  });
  removeAllBtn.dataset.fcMacAction = 'subtotal-remove-all';
  // Anchored on the left, away from Cancel / OK.
  removeAllBtn.style.marginRight = 'auto';
  const { cancelBtn, okBtn } = appendDialogActions(footer, {
    cancelLabel: t.cancel,
    okLabel: t.run,
    buttonBaseClass: 'fc-fmtdlg__btn fc-mac-subtotal__btn',
  });
  okBtn.dataset.fcMacAction = 'subtotal-ok';
  cancelBtn.dataset.fcMacAction = 'subtotal-cancel';

  let region: Range = instance.store.getState().selection.range;
  const checks: HTMLInputElement[] = [];

  const applyLabels = (): void => {
    strings = instance.i18n.strings.mac.data;
    t = strings.subtotal;
    shell.setAriaLabel(t.title);
    header.textContent = t.title;
    groupByRow.text.textContent = t.groupBy;
    functionRow.text.textContent = t.function;
    fieldsLabel.textContent = t.columns;
    for (const [fn, option] of functionOptions) option.textContent = strings.functions[fn];
    replaceOption.text.textContent = t.replace;
    pageBreakOption.text.textContent = t.pageBreak;
    summaryOption.text.textContent = t.summaryBelow;
    removeAllBtn.textContent = t.removeAll;
    okBtn.textContent = t.run;
    cancelBtn.textContent = t.cancel;
    syncCustomSelects(shell.panel);
  };
  applyLabels();
  const unsubscribeLocale = instance.i18n.on('change', applyLabels);

  const close = (): void => {
    shell.close();
  };

  /** The header row's text names each column; a blank header falls back to its letter. */
  const headerName = (col: number): string => {
    const value = instance.workbook.getValue({ sheet: region.sheet, row: region.r0, col });
    const text =
      value.kind === 'text' ? value.value : value.kind === 'number' ? String(value.value) : '';
    return text.trim() || colLetter(col);
  };

  const defaults = (): void => {
    region = inferAutoFilterRange(instance.store.getState());
    const columns = Array.from({ length: region.c1 - region.c0 + 1 }, (_, i) => region.c0 + i);
    groupBySelect.replaceChildren();
    appendDialogSelectOptions(
      groupBySelect,
      columns.map((col) => ({ value: String(col), label: headerName(col) })),
    );
    groupBySelect.value = String(region.c0);
    fields.replaceChildren();
    checks.length = 0;
    for (const col of columns) {
      const label = document.createElement('label');
      label.className = 'fc-mac-subtotal__field';
      const input = document.createElement('input');
      input.type = 'checkbox';
      input.value = String(col);
      // The last column is totalled by default.
      input.checked = col === region.c1 && columns.length > 1;
      const text = document.createElement('span');
      text.textContent = headerName(col);
      label.append(input, text);
      fields.appendChild(label);
      checks.push(input);
    }
    functionSelect.value = 'sum';
    replaceOption.input.checked = true;
    pageBreakOption.input.checked = false;
    summaryOption.input.checked = true;
    syncCustomSelects(shell.panel);
  };

  const fail = (message: string): void => {
    status.dataset.state = 'error';
    status.textContent = message;
  };

  const removeAll = (): void => {
    delete status.dataset.state;
    status.textContent = '';
    const removed = removeMacSubtotals(instance, region);
    if (!removed.ok) {
      fail(describeMacDataError(instance.i18n.strings, removed.error));
      return;
    }
    close();
  };

  const run = (): void => {
    delete status.dataset.state;
    status.textContent = '';
    const columns = checks.filter((input) => input.checked).map((input) => Number(input.value));
    if (columns.length === 0) {
      fail(strings.errors.invalidColumn);
      return;
    }
    const makeRequest = (target: Range): SubtotalRequest => ({
      range: formatMacRangeAddress(instance.workbook, target),
      groupByColumn: Number(groupBySelect.value),
      subtotalColumns: columns,
      function: (FUNCTIONS.includes(functionSelect.value as MacDataFunction)
        ? functionSelect.value
        : 'sum') as MacDataFunction,
      pageBreaks: pageBreakOption.input.checked,
      summaryBelow: summaryOption.input.checked,
    });
    // Replacing the old subtotals and writing the new ones is a single undo step.
    if (replaceOption.input.checked) {
      const replaced = replaceMacSubtotals(instance, region, makeRequest);
      if (!replaced.ok) {
        fail(describeMacDataError(instance.i18n.strings, replaced.error));
        return;
      }
      close();
      return;
    }
    const request = makeRequest(region);
    const plan = planMacSubtotal(instance, request);
    if (!plan.ok) {
      fail(describeMacDataError(instance.i18n.strings, plan.error));
      return;
    }
    const committed = commitMacSubtotal(instance, request, plan.value);
    if (!committed.ok) {
      fail(describeMacDataError(instance.i18n.strings, committed.error));
      return;
    }
    close();
  };

  shell.on(removeAllBtn, 'click', removeAll);
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
      queueMicrotask(() => groupBySelect.focus());
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
