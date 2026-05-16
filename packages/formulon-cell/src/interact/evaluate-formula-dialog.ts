import { extractRefs, type FormulaRef } from '../commands/refs.js';
import { formatCell } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import type { SpreadsheetStore } from '../store/store.js';
import { createDialogShell } from './dialog-shell.js';

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

  const header = document.createElement('div');
  header.className = 'fc-evaldlg__header';
  header.textContent = t.title;
  shell.panel.appendChild(header);

  const body = document.createElement('div');
  body.className = 'fc-evaldlg__body';
  shell.panel.appendChild(body);

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

  const footer = document.createElement('div');
  footer.className = 'fc-evaldlg__footer';
  shell.panel.appendChild(footer);

  const evalBtn = document.createElement('button');
  evalBtn.type = 'button';
  evalBtn.className = 'fc-evaldlg__btn';
  evalBtn.textContent = t.evaluate;

  const closeBtn = document.createElement('button');
  closeBtn.type = 'button';
  closeBtn.className = 'fc-evaldlg__btn fc-evaldlg__btn--primary';
  closeBtn.textContent = t.close;
  footer.append(evalBtn, closeBtn);

  let currentFormula = '';
  let currentRefs: FormulaRef[] = [];
  let stepIndex = 0;

  const uniqueRefCount = (refs: FormulaRef[]): number => new Set(refs.map(formulaRefKey)).size;

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
    evalBtn.disabled = currentRefs.length === 0 || stepIndex >= uniqueRefCount(currentRefs);
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
      evalBtn.disabled = true;
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
