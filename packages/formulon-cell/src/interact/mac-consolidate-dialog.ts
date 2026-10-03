import {
  type ConsolidateRequest,
  commitMacConsolidate,
  describeMacDataError,
  formatMacRangeAddress,
  type MacDataFunction,
  planMacConsolidate,
} from '../commands/mac-data-tools.js';
import type { SpreadsheetInstance } from '../mount/types.js';
import { appendDialogSelectOptions } from '../toolbar/dialogs/form-controls.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import { syncCustomSelects } from './custom-select.js';
import { appendDialogActions, appendDialogFrame, createDialogShell } from './dialog-shell.js';
import { isSubmitEnter } from './mac-dialog-keys.js';

interface ConsolidateDialogHandle {
  open(): void;
  close(): void;
  detach(): void;
}

const handles = new WeakMap<SpreadsheetInstance, ConsolidateDialogHandle>();
const FUNCTIONS: readonly MacDataFunction[] = ['sum', 'average', 'count', 'min', 'max'];

interface Row {
  readonly row: HTMLDivElement;
  readonly text: HTMLSpanElement;
}

const appendRow = (body: HTMLElement): Row => {
  const row = document.createElement('div');
  row.className = 'fc-fmtdlg__row fc-mac-consolidate__row';
  const text = document.createElement('span');
  row.appendChild(text);
  body.appendChild(row);
  return { row, text };
};

interface DisabledOption {
  readonly input: HTMLInputElement;
  readonly text: HTMLSpanElement;
}

const appendDisabledOption = (body: HTMLElement): DisabledOption => {
  const row = document.createElement('label');
  row.className = 'fc-mac-consolidate__option';
  const input = document.createElement('input');
  input.type = 'checkbox';
  const text = document.createElement('span');
  row.append(input, text);
  body.appendChild(row);
  return { input, text };
};

function attachConsolidateDialog(instance: SpreadsheetInstance): ConsolidateDialogHandle {
  let strings = instance.i18n.strings.macData;
  let t = strings.consolidate;
  const shell = createDialogShell({
    host: instance.host,
    className: 'fc-mac-consolidate',
    ariaLabel: t.title,
    onDismiss: () => api.close(),
  });
  shell.overlay.classList.add('fc-fmtdlg');
  const { header, body, footer } = appendDialogFrame(shell, {
    title: t.title,
    panelClasses: ['fc-fmtdlg__panel', 'fc-mac-consolidate__panel'],
    bodyClass: 'fc-fmtdlg__body fc-mac-consolidate__body',
    footerClass: 'fc-fmtdlg__footer fc-mac-consolidate__footer',
  });

  const sourceRow = appendRow(body);
  const sourceInput = document.createElement('textarea');
  sourceInput.id = 'fc-mac-consolidate-sources';
  sourceInput.rows = 4;
  sourceInput.className = 'fc-fmtdlg__input fc-mac-consolidate__sources';
  sourceInput.spellcheck = false;
  sourceRow.row.appendChild(sourceInput);

  const destinationRow = appendRow(body);
  const destinationInput = document.createElement('input');
  destinationInput.id = 'fc-mac-consolidate-destination';
  destinationInput.type = 'text';
  destinationInput.className = 'fc-fmtdlg__input';
  destinationInput.autocomplete = 'off';
  destinationInput.spellcheck = false;
  destinationRow.row.appendChild(destinationInput);

  const functionRow = appendRow(body);
  const functionSelect = document.createElement('select');
  functionSelect.id = 'fc-mac-consolidate-function';
  functionSelect.className = 'fc-fmtdlg__select';
  appendDialogSelectOptions(
    functionSelect,
    FUNCTIONS.map((fn) => ({ value: fn, label: '' })),
  );
  const functionOptions = new Map<MacDataFunction, HTMLOptionElement>(
    FUNCTIONS.map((fn, index) => [fn, functionSelect.options[index] as HTMLOptionElement]),
  );
  functionRow.row.appendChild(functionSelect);

  const replaceOption = appendDisabledOption(body);
  const labelsOption = appendDisabledOption(body);
  const linksOption = appendDisabledOption(body);

  const status = document.createElement('div');
  status.className = 'fc-mac-consolidate__status';
  status.setAttribute('role', 'status');
  status.setAttribute('aria-live', 'polite');
  body.appendChild(status);

  const { cancelBtn, okBtn } = appendDialogActions(footer, {
    cancelLabel: t.cancel,
    okLabel: t.run,
    buttonBaseClass: 'fc-fmtdlg__btn fc-mac-consolidate__btn',
  });
  okBtn.dataset.fcMacAction = 'consolidate-ok';
  cancelBtn.dataset.fcMacAction = 'consolidate-cancel';

  const applyLabels = (): void => {
    strings = instance.i18n.strings.macData;
    t = strings.consolidate;
    shell.setAriaLabel(t.title);
    header.textContent = t.title;
    sourceRow.text.textContent = t.sources;
    sourceInput.placeholder = t.sourcesHint;
    destinationRow.text.textContent = t.destination;
    functionRow.text.textContent = t.function;
    for (const [fn, option] of functionOptions) option.textContent = strings.functions[fn];
    const options: readonly [DisabledOption, string][] = [
      [replaceOption, t.replace],
      [labelsOption, t.labels],
      [linksOption, t.links],
    ];
    for (const [option, label] of options) {
      option.text.textContent = label;
      projectDisabledState(option.input, true, t.unsupportedOption, {
        datasetKey: 'disabledReason',
      });
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
    sourceInput.value = formatMacRangeAddress(instance.workbook, selection);
    destinationInput.value = '';
    functionSelect.value = 'sum';
  };

  const run = (): void => {
    delete status.dataset.state;
    status.textContent = '';
    const sources = sourceInput.value
      .split(/[\n;,]+/)
      .map((value) => value.trim())
      .filter(Boolean);
    const request: ConsolidateRequest = {
      sources,
      destination: destinationInput.value,
      function: (FUNCTIONS.includes(functionSelect.value as MacDataFunction)
        ? functionSelect.value
        : 'sum') as MacDataFunction,
    };
    const plan = planMacConsolidate(instance, request);
    if (!plan.ok) {
      status.dataset.state = 'error';
      status.textContent = describeMacDataError(instance.i18n.strings, plan.error);
      return;
    }
    const committed = commitMacConsolidate(instance, request, plan.value);
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

  const api: ConsolidateDialogHandle = {
    open(): void {
      applyLabels();
      defaults();
      delete status.dataset.state;
      status.textContent = '';
      shell.open();
      queueMicrotask(() => sourceInput.focus());
    },
    close,
    detach(): void {
      unsubscribeLocale();
      shell.dispose();
    },
  };
  return api;
}

/** Open the Data → Consolidate dialog. */
export function openMacConsolidateDialog(instance: SpreadsheetInstance): void {
  let handle = handles.get(instance);
  if (!handle) {
    handle = attachConsolidateDialog(instance);
    handles.set(instance, handle);
  }
  handle.open();
}

/** Dispose the Consolidate overlay owned by an instance, if it has been opened. */
export function disposeMacConsolidateDialog(instance: SpreadsheetInstance): void {
  const handle = handles.get(instance);
  if (!handle) return;
  handles.delete(instance);
  handle.detach();
}
