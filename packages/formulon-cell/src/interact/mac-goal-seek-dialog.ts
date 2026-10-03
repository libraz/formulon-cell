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
import type { Addr } from '../engine/types.js';
import type { SpreadsheetInstance } from '../mount/types.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import { appendDialogActions, appendDialogFrame, createDialogShell } from './dialog-shell.js';
import { isSubmitEnter } from './mac-dialog-keys.js';

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
  readonly input: HTMLInputElement;
  readonly text: HTMLSpanElement;
}

const appendField = (body: HTMLElement, id: string, type: 'text' | 'number'): Field => {
  const row = document.createElement('label');
  row.className = 'fc-fmtdlg__row fc-mac-goalseek__row';
  const text = document.createElement('span');
  const input = document.createElement('input');
  input.id = id;
  input.type = type;
  input.className = 'fc-fmtdlg__input';
  input.autocomplete = 'off';
  input.spellcheck = false;
  row.append(text, input);
  body.appendChild(row);
  return { input, text };
};

function attachGoalSeekDialog(instance: SpreadsheetInstance): GoalSeekDialogHandle {
  let t = instance.i18n.strings.macData.goalSeek;
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
  const targetField = appendField(body, 'fc-mac-goalseek-target-value', 'number');
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

  const status = document.createElement('div');
  status.className = 'fc-mac-goalseek__status';
  status.setAttribute('role', 'status');
  status.setAttribute('aria-live', 'polite');
  body.appendChild(status);
  const { cancelBtn, okBtn } = appendDialogActions(footer, {
    cancelLabel: t.cancel,
    okLabel: t.run,
    buttonBaseClass: 'fc-fmtdlg__btn fc-mac-goalseek__btn',
  });
  okBtn.dataset.fcMacAction = 'goal-seek-ok';
  cancelBtn.dataset.fcMacAction = 'goal-seek-cancel';

  const applyLabels = (): void => {
    t = instance.i18n.strings.macData.goalSeek;
    shell.setAriaLabel(t.title);
    header.textContent = t.title;
    formulaField.text.textContent = t.formulaCell;
    targetField.text.textContent = t.targetValue;
    changingField.text.textContent = t.changingCell;
    iterationField.text.textContent = t.iterations;
    okBtn.textContent = t.run;
    cancelBtn.textContent = t.cancel;
  };
  applyLabels();
  const unsubscribeLocale = instance.i18n.subscribe(applyLabels);

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
    instance.host.focus();
  };

  const currentDefaults = (): void => {
    const active = instance.store.getState().selection.active;
    const next: Addr = { ...active, col: Math.min(16_383, active.col + 1) };
    formulaInput.value = formatMacCellAddress(instance.workbook, active);
    changingInput.value = formatMacCellAddress(instance.workbook, next);
    const current = instance.workbook.getValue(active);
    targetInput.value =
      current.kind === 'number' && Number.isFinite(current.value) ? String(current.value) : '';
    iterationInput.value = '100';
  };

  const showError = (error: MacDataError): void => {
    status.dataset.state = 'error';
    status.textContent = describeMacDataError(instance.i18n.strings, error);
  };

  const showMessage = (message: string): void => {
    status.dataset.state = 'error';
    status.textContent = message;
  };

  const clearStatus = (): void => {
    delete status.dataset.state;
    status.textContent = '';
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
    const targetValue = Number(targetText);
    if (!targetText || !Number.isFinite(targetValue)) {
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
    status.textContent = t.calculating;
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
    status.dataset.state = 'success';
    status.textContent = `${t.converged} ${t.result}: ${result.value.changingValue}`;
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
