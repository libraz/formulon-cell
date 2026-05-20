import { extractRefs, type FormulaRef } from '../commands/refs.js';
import { formatCell } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import type { SpreadsheetStore } from '../store/store.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import { appendDialogButton, appendDialogFrame, createDialogShell } from './dialog-shell.js';

export interface EvaluateFormulaDialogDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  getWb: () => WorkbookHandle | null;
  strings?: Strings;
}

export interface EvaluateFormulaDialogHandle {
  open(): void;
  close(): void;
  detach(): void;
}

const cellRef = (row: number, col: number): string => {
  let n = col + 1;
  let letters = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return `${letters}${row + 1}`;
};

const formulaRefKey = (ref: FormulaRef): string => `${ref.r0}:${ref.c0}:${ref.r1}:${ref.c1}`;

const formulaRefValueText = (wb: WorkbookHandle | null, sheet: number, ref: FormulaRef): string => {
  if (!wb) return '';
  const value = wb.getValue({ sheet, row: ref.r0, col: ref.c0 });
  if (value.kind === 'text') return `"${value.value.replaceAll('"', '""')}"`;
  if (value.kind === 'blank') return '0';
  return formatCell(value);
};

const evaluatedFormulaText = (
  formula: string,
  refs: FormulaRef[],
  steps: number,
  wb: WorkbookHandle | null,
  sheet: number,
): string => {
  if (steps <= 0 || refs.length === 0) return formula;
  const replacedKeys = new Set<string>();
  const replacements: Array<{ start: number; end: number; text: string }> = [];
  for (const ref of refs) {
    if (replacements.length >= steps) break;
    const key = formulaRefKey(ref);
    if (replacedKeys.has(key)) continue;
    replacedKeys.add(key);
    replacements.push({
      start: ref.start,
      end: ref.end,
      text: formulaRefValueText(wb, sheet, ref),
    });
  }
  let out = formula;
  for (const replacement of replacements.sort((a, b) => b.start - a.start)) {
    out = `${out.slice(0, replacement.start)}${replacement.text}${out.slice(replacement.end)}`;
  }
  return out;
};

export function attachEvaluateFormulaDialog(
  deps: EvaluateFormulaDialogDeps,
): EvaluateFormulaDialogHandle {
  const { host, store, getWb } = deps;
  const strings = deps.strings ?? defaultStrings;
  const t = strings.evaluateFormulaDialog;

  const shell = createDialogShell({
    host,
    className: 'fc-evaldlg',
    ariaLabel: t.title,
    onDismiss: () => api.close(),
  });

  const { body, footer } = appendDialogFrame(shell, {
    title: t.title,
    headerClass: 'fc-evaldlg__header',
    bodyClass: 'fc-evaldlg__body',
    footerClass: 'fc-evaldlg__footer',
  });

  const target = document.createElement('div');
  target.className = 'fc-evaldlg__target';
  body.appendChild(target);

  const formulaLabel = document.createElement('div');
  formulaLabel.className = 'fc-evaldlg__label';
  formulaLabel.textContent = t.formula;
  body.appendChild(formulaLabel);

  const formulaBox = document.createElement('pre');
  formulaBox.className = 'fc-evaldlg__box';
  body.appendChild(formulaBox);

  const evaluationLabel = document.createElement('div');
  evaluationLabel.className = 'fc-evaldlg__label';
  evaluationLabel.textContent = t.evaluation;
  body.appendChild(evaluationLabel);

  const evaluationBox = document.createElement('pre');
  evaluationBox.className = 'fc-evaldlg__box fc-evaldlg__box--evaluation';
  body.appendChild(evaluationBox);

  const resultLabel = document.createElement('div');
  resultLabel.className = 'fc-evaldlg__label';
  resultLabel.textContent = t.result;
  body.appendChild(resultLabel);

  const resultBox = document.createElement('pre');
  resultBox.className = 'fc-evaldlg__box fc-evaldlg__box--result';
  body.appendChild(resultBox);

  const evalBtn = appendDialogButton(footer, {
    label: t.evaluate,
    baseClass: 'fc-evaldlg__btn',
  });
  const closeBtn = appendDialogButton(footer, {
    label: t.close,
    variant: 'primary',
    baseClass: 'fc-evaldlg__btn',
  });

  let currentFormula = '';
  let currentRefs: FormulaRef[] = [];
  let stepIndex = 0;

  const uniqueRefCount = (refs: FormulaRef[]): number => new Set(refs.map(formulaRefKey)).size;

  const setEvaluateDisabled = (disabled: boolean, reason: string | null): void => {
    projectDisabledState(evalBtn, disabled, reason, {
      datasetKey: 'disabledReason',
      titlePrefix: t.evaluate,
    });
  };

  const syncEvaluation = (): void => {
    const wb = getWb();
    const active = store.getState().selection.active;
    evaluationBox.textContent = evaluatedFormulaText(
      currentFormula,
      currentRefs,
      stepIndex,
      wb,
      active.sheet,
    );
    const uniqueRefs = uniqueRefCount(currentRefs);
    const reason =
      currentRefs.length === 0
        ? t.evaluateRequiresReference
        : stepIndex >= uniqueRefs
          ? t.evaluateComplete
          : null;
    setEvaluateDisabled(reason !== null, reason);
  };

  const refresh = (): void => {
    const wb = getWb();
    const active = store.getState().selection.active;
    target.textContent = cellRef(active.row, active.col);
    const formula =
      wb?.cellFormula(active) ??
      store.getState().data.cells.get(`${active.sheet}:${active.row}:${active.col}`)?.formula ??
      null;
    if (!formula) {
      formulaBox.textContent = t.noFormula;
      evaluationBox.textContent = '';
      resultBox.textContent = '';
      setEvaluateDisabled(true, t.evaluateRequiresFormula);
      return;
    }
    currentFormula = formula;
    currentRefs = extractRefs(formula);
    stepIndex = 0;
    formulaBox.textContent = formula;
    syncEvaluation();
    resultBox.textContent = formatCell(wb?.getValue(active) ?? { kind: 'blank' });
  };

  shell.on(evalBtn, 'click', () => {
    if (stepIndex < uniqueRefCount(currentRefs)) stepIndex += 1;
    syncEvaluation();
  });
  shell.on(closeBtn, 'click', () => api.close());

  const api: EvaluateFormulaDialogHandle = {
    open() {
      refresh();
      shell.open();
      closeBtn.focus({ preventScroll: true });
    },
    close() {
      shell.close();
    },
    detach() {
      shell.dispose();
    },
  };

  return api;
}
