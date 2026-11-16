import { parseLocalizedNumericInput } from '../commands/coerce-input.js';
import {
  cellSnapshot,
  commitMacGoalSeek,
  describeMacDataError,
  formatMacCellAddress,
  type GoalSeekRequest,
  type GoalSeekSolution,
  type MacDataError,
  parseMacRange,
  solveMacGoalSeek,
} from '../commands/mac-data-tools.js';
import { MAX_COL } from '../engine/address.js';
import type { Addr } from '../engine/types.js';
import { isMacHost } from '../extensions/ui-options.js';
import type { SpreadsheetInstance } from '../mount/types.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import {
  appendDialogActions,
  appendDialogFrame,
  createDialogShell,
  createDialogStatus,
} from './dialog-shell.js';
import { isSubmitEnter } from './mac-dialog-keys.js';
import { createRangePickFormatter } from './range-pick-format.js';
import { attachRangePickerButton, updateRangePickerLabel } from './range-picker-control.js';

interface GoalSeekDialogHandle {
  open(): void;
  close(): void;
  detach(): void;
}

const handles = new WeakMap<SpreadsheetInstance, GoalSeekDialogHandle>();

const singleCell = (instance: SpreadsheetInstance, input: string): Addr | null => {
  const parsed = parseMacRange(instance.workbook, input, instance.store.getState().data.sheetIndex);
  if (!parsed || parsed.range.r0 !== parsed.range.r1 || parsed.range.c0 !== parsed.range.c1)
    return null;
  return { sheet: parsed.range.sheet, row: parsed.range.r0, col: parsed.range.c0 };
};

interface Field {
  readonly row: HTMLLabelElement;
  readonly input: HTMLInputElement;
  readonly text: HTMLSpanElement;
}

const formatLocalizedNumericInitial = (value: number, locale: string | undefined): string => {
  try {
    return new Intl.NumberFormat(locale, {
      useGrouping: false,
      maximumSignificantDigits: 17,
    }).format(value);
  } catch {
    // A custom dictionary may use an invalid BCP 47 tag. Keep the dialog
    // usable while the input parser falls back to its default grammar.
    return String(value);
  }
};

const appendField = (
  body: HTMLElement,
  id: string,
  type: 'text' | 'number',
  inputMode?: HTMLInputElement['inputMode'],
): Field => {
  const row = document.createElement('label');
  row.className = 'fc-fmtdlg__row fc-mac-goalseek__row';
  const text = document.createElement('span');
  const input = document.createElement('input');
  input.id = id;
  input.type = type;
  if (inputMode) input.inputMode = inputMode;
  input.className = 'fc-fmtdlg__input';
  input.autocomplete = 'off';
  input.spellcheck = false;
  row.append(text, input);
  body.appendChild(row);
  return { row, input, text };
};

function attachGoalSeekDialog(instance: SpreadsheetInstance): GoalSeekDialogHandle {
  let t = instance.i18n.strings.mac.data.goalSeek;
  const shell = createDialogShell({
    host: instance.host,
    className: 'fc-mac-goalseek',
    ariaLabel: t.title,
    onDismiss: () => api.close(),
  });
  shell.overlay.classList.add('fc-fmtdlg');
  const { header, body, footer } = appendDialogFrame(shell, {
    title: t.title,
    panelClasses: ['fc-fmtdlg__panel', 'fc-mac-goalseek__panel'],
    bodyClass: 'fc-fmtdlg__body fc-mac-goalseek__body',
    footerClass: 'fc-fmtdlg__footer fc-mac-goalseek__footer',
  });
  const formulaField = appendField(body, 'fc-mac-goalseek-formula-cell', 'text');
  const targetField = appendField(body, 'fc-mac-goalseek-target-value', 'text', 'decimal');
  const changingField = appendField(body, 'fc-mac-goalseek-changing-cell', 'text');
  const iterationField = appendField(body, 'fc-mac-goalseek-iterations', 'number');
  const formulaInput = formulaField.input;
  const targetInput = targetField.input;
  const changingInput = changingField.input;
  const iterationInput = iterationField.input;
  iterationInput.min = '1';
  iterationInput.max = '100';
  iterationInput.step = '1';
  iterationInput.value = '100';
  const pickFormat = createRangePickFormatter({
    store: instance.store,
    sheetName: (index) => instance.workbook.sheetName(index),
  });
  const pickers = [formulaInput, changingInput].map((input) =>
    attachRangePickerButton(input, {
      label: instance.i18n.strings.pivotTableDialog.rangePickerSelect,
      getValue: () => pickFormat.cell(),
      subscribeToRangeChanges: pickFormat.subscribe,
      kind: input.id,
    }),
  );

  const status = createDialogStatus(body, 'fc-mac-goalseek__status');
  const { cancelBtn, okBtn } = appendDialogActions(footer, {
    cancelLabel: t.cancel,
    okLabel: t.run,
    buttonBaseClass: 'fc-fmtdlg__btn fc-mac-goalseek__btn',
  });
  okBtn.dataset.fcMacAction = 'goal-seek-ok';
  cancelBtn.dataset.fcMacAction = 'goal-seek-cancel';

  const applyLabels = (): void => {
    t = instance.i18n.strings.mac.data.goalSeek;
    shell.setAriaLabel(t.title);
    header.textContent = t.title;
    formulaField.text.textContent = t.formulaCell;
    targetField.text.textContent = t.targetValue;
    changingField.text.textContent = t.changingCell;
    iterationField.text.textContent = t.iterations;
    okBtn.textContent = t.run;
    cancelBtn.textContent = t.cancel;
    for (const picker of pickers) {
      updateRangePickerLabel(picker, instance.i18n.strings.pivotTableDialog.rangePickerSelect);
    }
  };
  applyLabels();
  const unsubscribeLocale = instance.i18n.on('change', applyLabels);

  let runToken = 0;
  let solving = false;
  let pending: {
    readonly request: GoalSeekRequest;
    readonly formulaSnapshot: ReturnType<typeof cellSnapshot>;
    readonly changingSnapshot: ReturnType<typeof cellSnapshot>;
    readonly solution: GoalSeekSolution;
  } | null = null;

  const setInputsDisabled = (disabled: boolean): void => {
    const reason = disabled ? t.calculating : null;
    projectDisabledState(formulaInput, disabled, reason, { datasetKey: 'disabledReason' });
    projectDisabledState(targetInput, disabled, reason, { datasetKey: 'disabledReason' });
    projectDisabledState(changingInput, disabled, reason, { datasetKey: 'disabledReason' });
    projectDisabledState(iterationInput, disabled, reason, { datasetKey: 'disabledReason' });
  };

  const setBusy = (busy: boolean): void => {
    solving = busy;
    setInputsDisabled(busy);
    projectDisabledState(okBtn, busy, busy ? t.calculating : null, {
      datasetKey: 'disabledReason',
    });
    projectDisabledState(cancelBtn, false, null, { datasetKey: 'disabledReason' });
  };

  const close = (): void => {
    runToken += 1;
    pending = null;
    setBusy(false);
    shell.close();
  };

  const currentDefaults = (): void => {
    pickFormat.begin();
    const active = instance.store.getState().selection.active;
    const next: Addr = { ...active, col: Math.min(MAX_COL, active.col + 1) };
    formulaInput.value = formatMacCellAddress(active);
    changingInput.value = formatMacCellAddress(next);
    iterationInput.value = '100';
    // The Mac desktop dialog starts with only the formula cell filled and has
    // no iteration field: its iteration cap is internal.
    const mac = isMacHost(instance.host);
    iterationField.row.hidden = mac;
    if (mac) {
      changingInput.value = '';
      targetInput.value = '';
      return;
    }
    const current = instance.workbook.getValue(active);
    targetInput.value =
      current.kind === 'number' && Number.isFinite(current.value)
        ? formatLocalizedNumericInitial(current.value, instance.locale || instance.i18n.locale)
        : '';
  };

  const showError = (error: MacDataError): void => {
    status.error(describeMacDataError(instance.i18n.strings, error));
  };

  const showMessage = (message: string): void => {
    status.error(message);
  };

  const clearStatus = (): void => {
    status.clear();
  };

  const run = async (): Promise<void> => {
    if (solving) return;
    if (pending) {
      const committed = commitMacGoalSeek(
        instance,
        pending.request,
        pending.solution,
        pending.formulaSnapshot,
        pending.changingSnapshot,
      );
      if (!committed.ok) {
        // The proposed solution is unusable; return to an editable state.
        pending = null;
        setBusy(false);
        showError(committed.error);
        return;
      }
      close();
      return;
    }
    clearStatus();
    const formulaCell = singleCell(instance, formulaInput.value);
    const changingCell = singleCell(instance, changingInput.value);
    if (!formulaCell || !changingCell) {
      showMessage(t.invalidCell);
      return;
    }
    const targetText = targetInput.value.trim();
    const targetValue = parseLocalizedNumericInput(
      targetText,
      instance.locale || instance.i18n.locale,
    );
    if (targetValue === null) {
      showMessage(t.invalidTarget);
      return;
    }
    const maxIterations = Number.parseInt(iterationInput.value, 10);
    const request: GoalSeekRequest = {
      formulaCell,
      changingCell,
      targetValue,
      maxIterations: Number.isFinite(maxIterations) ? maxIterations : 100,
    };
    const formulaSnapshot = cellSnapshot(instance.workbook, formulaCell);
    const changingSnapshot = cellSnapshot(instance.workbook, changingCell);
    const token = ++runToken;
    setBusy(true);
    status.el.textContent = t.calculating;
    const result = await solveMacGoalSeek(instance, request);
    if (token !== runToken) return;
    if (!result.ok) {
      setBusy(false);
      showError(result.error);
      return;
    }
    setBusy(false);
    pending = {
      request,
      formulaSnapshot,
      changingSnapshot,
      solution: result.value,
    };
    setInputsDisabled(true);
    // A non-converged run still offers its last iterate, as the desktop app keeps it.
    status.el.dataset.state = result.value.converged ? 'success' : 'error';
    const outcome = result.value.converged
      ? t.converged
      : instance.i18n.strings.mac.data.errors.notConverged;
    status.el.textContent = `${outcome} ${t.result}: ${result.value.changingValue}`;
  };

  shell.on(okBtn, 'click', () => {
    void run();
  });
  shell.on(cancelBtn, 'click', close);
  shell.on(shell.overlay, 'keydown', (event) => {
    const e = event as KeyboardEvent;
    if (!isSubmitEnter(e)) return;
    e.preventDefault();
    void run();
  });

  const api: GoalSeekDialogHandle = {
    open(): void {
      applyLabels();
      currentDefaults();
      clearStatus();
      shell.open();
      queueMicrotask(() => formulaInput.focus());
    },
    close,
    detach(): void {
      runToken += 1;
      pending = null;
      setBusy(false);
      unsubscribeLocale();
      shell.dispose();
    },
  };
  return api;
}

/** Open the Data → What-If Analysis → Goal Seek dialog. */
export function openMacGoalSeekDialog(instance: SpreadsheetInstance): void {
  let handle = handles.get(instance);
  if (!handle) {
    handle = attachGoalSeekDialog(instance);
    handles.set(instance, handle);
  }
  handle.open();
}

/** Dispose the Goal Seek overlay owned by an instance, if it has been opened. */
export function disposeMacGoalSeekDialog(instance: SpreadsheetInstance): void {
  const handle = handles.get(instance);
  if (!handle) return;
  handles.delete(instance);
  handle.detach();
}
