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
import { createInteractionChipButton } from './chip-button.js';
import { syncCustomSelects } from './custom-select.js';
import {
  appendDialogActions,
  appendDialogFrame,
  createDialogButton,
  createDialogShell,
  createDialogStatus,
  deferInitialFocus,
} from './dialog-shell.js';
import { isSubmitEnter } from './mac-dialog-keys.js';
import { createRangePickFormatter } from './range-pick-format.js';
import { attachRangePickerButton } from './range-picker-control.js';

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
  let strings = instance.i18n.strings.mac.data;
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

  // Desktop order: function first, then the source reference.
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

  // The reference is typed into a single field and added to the list below it.
  const sourceRow = appendRow(body);
  const sourceInput = document.createElement('input');
  sourceInput.id = 'fc-mac-consolidate-sources';
  sourceInput.type = 'text';
  sourceInput.className = 'fc-fmtdlg__input fc-mac-consolidate__source';
  sourceInput.autocomplete = 'off';
  sourceInput.spellcheck = false;
  sourceRow.row.appendChild(sourceInput);
  const pickFormat = createRangePickFormatter({
    store: instance.store,
    sheetName: (index) => instance.workbook.sheetName(index),
  });
  attachRangePickerButton(sourceInput, {
    label: instance.i18n.strings.pivotTableDialog.rangePickerSelect,
    getValue: () => pickFormat.range(),
    subscribeToRangeChanges: pickFormat.subscribe,
    kind: 'mac-consolidate-source',
  });

  const listRow = appendRow(body);
  listRow.row.classList.add('fc-mac-consolidate__row--list');
  const list = document.createElement('div');
  list.className = 'fc-mac-consolidate__list';
  list.setAttribute('role', 'listbox');
  listRow.row.appendChild(list);
  const listActions = document.createElement('div');
  listActions.className = 'fc-mac-consolidate__list-actions';
  const selectBtn = createDialogButton({ label: '', baseClass: 'fc-fmtdlg__btn' });
  projectDisabledState(selectBtn, true, t.unsupportedOption, { datasetKey: 'disabledReason' });
  const addBtn = createDialogButton({
    label: '',
    baseClass: 'fc-fmtdlg__btn fc-mac-consolidate__small',
  });
  addBtn.dataset.fcMacAction = 'consolidate-add';
  addBtn.textContent = '+';
  const removeBtn = createDialogButton({
    label: '',
    baseClass: 'fc-fmtdlg__btn fc-mac-consolidate__small',
  });
  removeBtn.dataset.fcMacAction = 'consolidate-remove';
  removeBtn.textContent = '−';
  listActions.append(selectBtn, addBtn, removeBtn);
  body.appendChild(listActions);

  const references: string[] = [];
  let selectedReference = -1;
  const renderReferences = (): void => {
    list.replaceChildren(
      ...references.map((reference, index) => {
        const item = createInteractionChipButton({
          className: 'fc-mac-consolidate__reference',
          label: reference,
          role: 'option',
          selected: index === selectedReference,
        });
        item.addEventListener('click', () => {
          selectedReference = index;
          renderReferences();
        });
        return item;
      }),
    );
    projectDisabledState(removeBtn, selectedReference < 0, null);
  };
  const addReference = (): void => {
    const text = sourceInput.value.trim();
    if (!text) return;
    references.push(text);
    sourceInput.value = '';
    selectedReference = references.length - 1;
    renderReferences();
  };
  const removeReference = (): void => {
    if (selectedReference < 0) return;
    references.splice(selectedReference, 1);
    selectedReference = Math.min(selectedReference, references.length - 1);
    renderReferences();
  };

  const linksOption = appendDisabledOption(body);

  const criteria = document.createElement('div');
  criteria.className = 'fc-mac-consolidate__criteria';
  const topRowOption = appendDisabledOption(criteria);
  const leftColumnOption = appendDisabledOption(criteria);
  body.append(criteria);
  const criteriaHeading = document.createElement('div');
  criteriaHeading.className = 'fc-mac-consolidate__criteria-heading';
  criteria.prepend(criteriaHeading);

  const status = createDialogStatus(body, 'fc-mac-consolidate__status');

  const { cancelBtn, okBtn } = appendDialogActions(footer, {
    cancelLabel: t.cancel,
    okLabel: t.run,
    buttonBaseClass: 'fc-fmtdlg__btn fc-mac-consolidate__btn',
  });
  okBtn.dataset.fcMacAction = 'consolidate-ok';
  cancelBtn.dataset.fcMacAction = 'consolidate-cancel';

  const applyLabels = (): void => {
    strings = instance.i18n.strings.mac.data;
    t = strings.consolidate;
    shell.setAriaLabel(t.title);
    header.textContent = t.title;
    sourceRow.text.textContent = t.sources;
    listRow.text.textContent = t.referenceList;
    selectBtn.textContent = t.select;
    addBtn.setAttribute('aria-label', t.add);
    removeBtn.setAttribute('aria-label', t.remove);
    criteriaHeading.textContent = t.criteria;
    topRowOption.text.textContent = t.topRow;
    leftColumnOption.text.textContent = t.leftColumn;
    functionRow.text.textContent = t.function;
    for (const [fn, option] of functionOptions) option.textContent = strings.functions[fn];
    linksOption.text.textContent = t.links;
    projectDisabledState(linksOption.input, true, t.unsupportedOption, {
      datasetKey: 'disabledReason',
    });
    okBtn.textContent = t.run;
    cancelBtn.textContent = t.cancel;
    syncCustomSelects(shell.panel);
  };
  applyLabels();
  const unsubscribeLocale = instance.i18n.on('change', applyLabels);

  const close = (): void => {
    shell.close();
  };

  const defaults = (): void => {
    pickFormat.begin();
    const selection = instance.store.getState().selection.range;
    sourceInput.value = formatMacRangeAddress(instance.workbook, selection);
    references.length = 0;
    selectedReference = -1;
    renderReferences();
    topRowOption.input.checked = false;
    leftColumnOption.input.checked = false;
    functionSelect.value = 'sum';
  };

  const run = (): void => {
    status.clear();
    // A reference still in the field counts, so the list need not be filled first.
    const pending = sourceInput.value.trim();
    const sources = pending ? [...references, pending] : [...references];
    const active = instance.store.getState().selection.active;
    const request: ConsolidateRequest = {
      sources,
      // The result starts at the active cell, as on the desktop.
      destination: formatMacRangeAddress(instance.workbook, {
        sheet: active.sheet,
        r0: active.row,
        c0: active.col,
        r1: active.row,
        c1: active.col,
      }),
      topRow: topRowOption.input.checked,
      leftColumn: leftColumnOption.input.checked,
      function: (FUNCTIONS.includes(functionSelect.value as MacDataFunction)
        ? functionSelect.value
        : 'sum') as MacDataFunction,
    };
    const plan = planMacConsolidate(instance, request);
    if (!plan.ok) {
      status.error(describeMacDataError(instance.i18n.strings, plan.error));
      return;
    }
    const committed = commitMacConsolidate(instance, request, plan.value);
    if (!committed.ok) {
      status.error(describeMacDataError(instance.i18n.strings, committed.error));
      return;
    }
    close();
  };

  shell.on(addBtn, 'click', addReference);
  shell.on(removeBtn, 'click', removeReference);
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
      status.clear();
      shell.open();
      deferInitialFocus(shell.overlay, () => sourceInput.focus());
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
