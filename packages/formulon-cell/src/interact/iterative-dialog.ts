import { inputLocaleForWorkbook, parseLocalizedNumericInput } from '../commands/coerce-input.js';
import type { History } from '../commands/history.js';
import { setPrecisionAsDisplayed } from '../commands/precision-as-displayed.js';
import { hydrateCalcSettingsFromEngine } from '../engine/calc-settings-sync.js';
import { formatEditableNumber } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { WorkbookCalcSettings } from '../engine/workbook-handle-formulas.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { showConfirm } from '../toolbar/dialogs/prompt.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import {
  appendDialogActions,
  appendDialogFrame,
  createDialogShell,
  deferInitialFocus,
} from './dialog-shell.js';

export interface IterativeDialogDeps {
  host: HTMLElement;
  /** Lazy workbook accessor — keeps the dialog in lockstep with `setWorkbook`
   *  swaps so every setting always targets the live engine. */
  getWb: () => WorkbookHandle | null;
  /** Live input locale. Defaults to the canonical English grammar. */
  getLocale?: () => string;
  strings?: Strings;
  /** Store mirroring the reference style and date system. Without it those
   *  settings go to the engine directly. */
  store?: SpreadsheetStore;
  /** Undo stack for the constants Set precision as displayed rounds. */
  history?: History | null;
}

export interface IterativeDialogHandle {
  open(): void;
  close(): void;
  detach(): void;
}

type CalcMode = 0 | 1 | 2;

interface CalcOptionsDraft extends WorkbookCalcSettings {
  calcMode: CalcMode | null;
  iterative: boolean;
  maxIterations: number;
  maxChange: number;
}

const DEFAULTS: CalcOptionsDraft = {
  calcMode: 0,
  calcOnSave: true,
  iterative: false,
  maxIterations: 100,
  maxChange: 0.001,
  r1c1: false,
  fullPrecision: true,
  date1904: false,
};

const CALC_MODES: readonly { mode: CalcMode; option: string }[] = [
  { mode: 0, option: 'auto' },
  { mode: 2, option: 'auto-no-table' },
  { mode: 1, option: 'manual' },
];

/**
 * Calculation options: workbook calculation mode, recalculate before saving,
 * iterative calculation, R1C1 reference style, Set precision as displayed and
 * the 1904 date system — the desktop Formulas / Calculation options pages.
 * Iteration wires a progress callback that surfaces residual / iteration
 * counts through the `fc:iterative-progress` custom event.
 */
export function attachIterativeDialog(deps: IterativeDialogDeps): IterativeDialogHandle {
  const { host, getWb } = deps;
  const getLocale = (): string => {
    const workbook = getWb();
    return (
      deps.getLocale?.() ?? (workbook ? inputLocaleForWorkbook(workbook) : undefined) ?? 'en-US'
    );
  };
  const strings = deps.strings ?? defaultStrings;
  const t = strings.iterativeDialog;
  const menuText = strings.ribbonMenu;

  const shell = createDialogShell({
    host,
    className: 'fc-iterdlg',
    ariaLabel: t.title,
    onDismiss: () => api.close(),
  });

  shell.overlay.classList.add('fc-fmtdlg');
  const { body, footer } = appendDialogFrame(shell, {
    title: t.title,
    panelClasses: ['fc-fmtdlg__panel'],
    headerClass: 'fc-fmtdlg__header fc-iterdlg__header',
    bodyClass: 'fc-fmtdlg__body fc-iterdlg__body',
    footerClass: 'fc-fmtdlg__footer fc-iterdlg__footer',
  });

  const section = (title: string): void => {
    const heading = document.createElement('div');
    heading.className = 'fc-iterdlg__section';
    heading.textContent = title;
    body.appendChild(heading);
  };
  const choice = (type: 'checkbox' | 'radio', option: string, label: string): HTMLInputElement => {
    const row = document.createElement('label');
    row.className = 'fc-iterdlg__row';
    const input = document.createElement('input');
    input.type = type;
    input.dataset.calcOption = option;
    if (type === 'radio') input.name = 'fc-iterdlg-calc-mode';
    const span = document.createElement('span');
    span.textContent = label;
    row.append(input, span);
    body.appendChild(row);
    return input;
  };

  section(t.calculationSection);
  const modeLabel = document.createElement('div');
  modeLabel.className = 'fc-iterdlg__label';
  modeLabel.textContent = t.workbookCalculation;
  body.appendChild(modeLabel);
  const modeLabels: Record<CalcMode, string> = {
    0: menuText.calcAutomatic,
    1: menuText.calcManual,
    2: menuText.calcAutoNoTable,
  };
  const modeInputs = CALC_MODES.map(({ mode, option }) => ({
    mode,
    input: choice('radio', option, modeLabels[mode]),
  }));
  const calcOnSaveInput = choice('checkbox', 'calcOnSave', t.recalcBeforeSave);

  const note = document.createElement('p');
  note.className = 'fc-iterdlg__note';
  note.textContent = t.note;
  body.appendChild(note);

  const enableInput = choice('checkbox', 'iterative', t.enable);

  const maxIterRow = document.createElement('label');
  maxIterRow.className = 'fc-iterdlg__row';
  const maxIterLabel = document.createElement('span');
  maxIterLabel.textContent = t.maxIterations;
  const maxIterInput = document.createElement('input');
  maxIterInput.type = 'number';
  maxIterInput.min = '1';
  maxIterInput.max = '32767';
  maxIterInput.step = '1';
  maxIterRow.append(maxIterLabel, maxIterInput);
  body.appendChild(maxIterRow);

  const maxChangeRow = document.createElement('label');
  maxChangeRow.className = 'fc-iterdlg__row';
  const maxChangeLabel = document.createElement('span');
  maxChangeLabel.textContent = t.maxChange;
  const maxChangeInput = document.createElement('input');
  maxChangeInput.type = 'text';
  maxChangeInput.inputMode = 'decimal';
  maxChangeInput.spellcheck = false;
  maxChangeInput.autocomplete = 'off';
  maxChangeRow.append(maxChangeLabel, maxChangeInput);
  body.appendChild(maxChangeRow);

  section(t.formulasSection);
  const r1c1Input = choice('checkbox', 'r1c1', t.r1c1);
  section(t.workbookSection);
  const precisionInput = choice('checkbox', 'precisionAsDisplayed', t.precisionAsDisplayed);
  const date1904Input = choice('checkbox', 'date1904', t.date1904);

  const status = document.createElement('div');
  status.className = 'fc-iterdlg__status';
  body.appendChild(status);

  const { cancelBtn, okBtn } = appendDialogActions(footer, {
    cancelLabel: t.cancel,
    okLabel: t.ok,
    buttonBaseClass: 'fc-fmtdlg__btn fc-iterdlg__btn',
    buttonPrimaryClass: 'fc-fmtdlg__btn--primary fc-iterdlg__btn--primary',
  });

  const draft: CalcOptionsDraft = { ...DEFAULTS };
  let opened: CalcOptionsDraft = { ...DEFAULTS };
  let invalidMaxIterations = false;
  let invalidMaxChange = false;
  let confirmingPrecision = false;

  const workbookSupports = (): { calcMode: boolean; settings: boolean } => {
    const caps = getWb()?.capabilities;
    return { calcMode: !!caps?.calcMode, settings: !!caps?.calcSettings };
  };

  const disable = (input: HTMLInputElement, reason: string | null, titlePrefix?: string): void => {
    projectDisabledState(input, reason !== null, reason, {
      datasetKey: 'disabledReason',
      ...(titlePrefix ? { titlePrefix } : {}),
    });
  };

  const syncControls = (): void => {
    const supports = workbookSupports();
    const unsaved = t.settingUnavailable;
    for (const { mode, input } of modeInputs) {
      input.checked = draft.calcMode === mode;
      disable(input, supports.calcMode ? null : unsaved);
    }
    calcOnSaveInput.checked = draft.calcOnSave;
    disable(
      calcOnSaveInput,
      !supports.settings ? unsaved : draft.calcMode !== 1 ? t.recalcBeforeSaveRequiresManual : null,
    );
    enableInput.checked = draft.iterative;
    if (!invalidMaxIterations) maxIterInput.value = String(draft.maxIterations);
    if (!invalidMaxChange)
      maxChangeInput.value = formatEditableNumber(draft.maxChange, getLocale());
    const iterationReason = draft.iterative ? null : t.inputsRequireEnabled;
    disable(maxIterInput, iterationReason, t.maxIterations);
    disable(maxChangeInput, iterationReason, t.maxChange);
    r1c1Input.checked = draft.r1c1;
    precisionInput.checked = !draft.fullPrecision;
    date1904Input.checked = draft.date1904;
    disable(date1904Input, supports.settings ? null : unsaved);
    // Rounding the constants needs the store's number formats.
    disable(precisionInput, supports.settings && deps.store ? null : unsaved);
    // The store alone can still flip the reference style for the session.
    disable(r1c1Input, supports.settings || deps.store ? null : unsaved);
  };

  const onIter = (): void => {
    const n = parseLocalizedNumericInput(maxIterInput.value, getLocale());
    invalidMaxIterations = n === null || !Number.isInteger(n);
    maxIterInput.setCustomValidity(invalidMaxIterations ? 'invalid' : '');
    if (n !== null && !invalidMaxIterations) {
      draft.maxIterations = Math.max(1, Math.min(32767, n));
    }
  };
  const onChange = (): void => {
    const n = parseLocalizedNumericInput(maxChangeInput.value, getLocale());
    invalidMaxChange = n === null || n <= 0;
    maxChangeInput.setCustomValidity(invalidMaxChange ? 'invalid' : '');
    if (n !== null && !invalidMaxChange) draft.maxChange = n;
  };

  /** Turning precision as displayed on loses data, so the desktop app asks
   *  first; declining leaves the box clear. */
  const onPrecision = async (): Promise<void> => {
    if (!precisionInput.checked) {
      draft.fullPrecision = true;
      return;
    }
    if (!opened.fullPrecision) {
      draft.fullPrecision = false;
      return;
    }
    confirmingPrecision = true;
    const accepted = await showConfirm({
      title: t.title,
      message: t.precisionWarning,
      okLabel: t.ok,
      cancelLabel: t.cancel,
    });
    confirmingPrecision = false;
    draft.fullPrecision = !accepted;
    syncControls();
  };

  const applyIterative = (wb: WorkbookHandle): boolean => {
    const changed =
      draft.iterative !== opened.iterative ||
      draft.maxIterations !== opened.maxIterations ||
      draft.maxChange !== opened.maxChange;
    if (!changed && !draft.iterative) return true;
    if (!wb.setIterative(draft.iterative, draft.maxIterations, draft.maxChange)) return false;
    if (draft.iterative) {
      // Wire the progress callback to a host event the chrome can render.
      wb.setIterativeProgress((iteration, maxResidual, maxIterations) => {
        host.dispatchEvent(
          new CustomEvent('fc:iterative-progress', {
            detail: { iteration, maxResidual, maxIterations },
          }),
        );
      });
    } else {
      wb.setIterativeProgress(null);
    }
    return true;
  };

  const applySettings = (wb: WorkbookHandle): void => {
    const { store } = deps;
    if (draft.calcMode !== null && draft.calcMode !== opened.calcMode) {
      wb.setCalcMode(draft.calcMode);
    }
    if (draft.calcOnSave !== opened.calcOnSave) wb.setCalcSetting('calcOnSave', draft.calcOnSave);
    if (draft.r1c1 !== opened.r1c1) {
      // The store binding writes refMode through.
      if (store) mutators.setR1C1(store, draft.r1c1);
      else wb.setCalcSetting('r1c1', draft.r1c1);
    }
    if (draft.date1904 !== opened.date1904 && wb.setCalcSetting('date1904', draft.date1904)) {
      // Stored serials stay put; the engine re-dirties the date formulas.
      if (store) hydrateCalcSettingsFromEngine(wb, store);
      wb.recalcAuto();
    }
    if (draft.fullPrecision !== opened.fullPrecision) {
      if (store) setPrecisionAsDisplayed(wb, store, deps.history ?? null, !draft.fullPrecision);
      else wb.setCalcSetting('fullPrecision', draft.fullPrecision);
    }
  };

  const onOk = (): void => {
    if (confirmingPrecision) return;
    if (invalidMaxIterations || invalidMaxChange) {
      const invalid = invalidMaxIterations ? maxIterInput : maxChangeInput;
      invalid.setCustomValidity('invalid');
      invalid.focus();
      return;
    }
    const wb = getWb();
    if (!wb) {
      api.close();
      return;
    }
    if (!applyIterative(wb)) {
      status.textContent = t.unsupported;
      return;
    }
    applySettings(wb);
    // Mode, date-system and precision changes can recalculate formulas.
    if (deps.store) {
      mutators.replaceCells(deps.store, wb.cells(deps.store.getState().data.sheetIndex));
    }
    api.close();
  };

  for (const { mode, input } of modeInputs) {
    shell.on(input, 'change', () => {
      if (!input.checked) return;
      draft.calcMode = mode;
      syncControls();
    });
  }
  shell.on(calcOnSaveInput, 'change', () => {
    draft.calcOnSave = calcOnSaveInput.checked;
  });
  shell.on(enableInput, 'change', () => {
    draft.iterative = enableInput.checked;
    syncControls();
  });
  shell.on(maxIterInput, 'input', onIter);
  shell.on(maxChangeInput, 'input', onChange);
  shell.on(r1c1Input, 'change', () => {
    draft.r1c1 = r1c1Input.checked;
  });
  shell.on(precisionInput, 'change', () => {
    void onPrecision();
  });
  shell.on(date1904Input, 'change', () => {
    draft.date1904 = date1904Input.checked;
  });
  shell.on(okBtn, 'click', onOk);
  shell.on(cancelBtn, 'click', () => api.close());
  shell.on(shell.overlay, 'keydown', (e) => {
    const event = e as KeyboardEvent;
    event.stopPropagation();
    if (event.key === 'Enter') {
      event.preventDefault();
      onOk();
    }
  });

  /** Open on what the workbook carries — a loaded file can hold settings this
   *  dialog never wrote, and the engine clamps the iteration cap on the way
   *  in. Engines without a readback leave that part of the draft alone. */
  const refreshFromWorkbook = (): void => {
    Object.assign(draft, DEFAULTS);
    const workbook = getWb();
    if (workbook) {
      if (typeof workbook.calcMode === 'function') draft.calcMode = workbook.calcMode();
      const settings = typeof workbook.calcSettings === 'function' ? workbook.calcSettings() : null;
      if (settings) Object.assign(draft, settings);
      const iterative =
        typeof workbook.getIterative === 'function' ? workbook.getIterative() : null;
      if (iterative) {
        draft.iterative = iterative.enabled;
        if (iterative.maxIterations > 0) draft.maxIterations = iterative.maxIterations;
        if (iterative.maxChange > 0) draft.maxChange = iterative.maxChange;
      }
    }
    if (deps.store) draft.r1c1 = deps.store.getState().ui.r1c1;
    opened = { ...draft };
  };

  const api: IterativeDialogHandle = {
    open(): void {
      status.textContent = '';
      invalidMaxIterations = false;
      invalidMaxChange = false;
      confirmingPrecision = false;
      maxIterInput.setCustomValidity('');
      maxChangeInput.setCustomValidity('');
      refreshFromWorkbook();
      syncControls();
      shell.open();
      deferInitialFocus(shell.overlay, () => enableInput.focus());
    },
    close(): void {
      shell.close();
      host.focus();
    },
    detach(): void {
      shell.dispose();
    },
  };
  return api;
}
