import { inputLocaleForWorkbook } from '../commands/coerce-input.js';
import { extractRefs, type FormulaRef } from '../commands/refs.js';
import { formatA1Cell } from '../engine/address.js';
import type { CellValue } from '../engine/types.js';
import { formatCell } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { bundledLocaleForStrings } from '../i18n/locale-for-strings.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import type { SpreadsheetStore } from '../store/store.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import { appendDialogButton, appendDialogFrame, createDialogShell } from './dialog-shell.js';

export interface EvaluateFormulaDialogDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  getWb: () => WorkbookHandle | null;
  strings?: Strings;
  /** Current UI locale for display fallbacks when the workbook adapter does
   * not expose native display text. */
  getLocale?: () => string;
}

export interface EvaluateFormulaDialogHandle {
  open(): void;
  close(): void;
  detach(): void;
}

const formulaRefKey = (ref: FormulaRef, formula = ''): string =>
  `${ref.r0}:${ref.c0}:${ref.r1}:${ref.c1}:${formula.slice(ref.start, ref.end)}`;

const canonicalValueText = (value: CellValue): string => {
  switch (value.kind) {
    case 'blank':
      return '0';
    case 'number':
      return String(value.value);
    case 'bool':
      return value.value ? 'TRUE' : 'FALSE';
    case 'error':
      return value.text || '#ERROR!';
    case 'text':
      return `"${value.value.replaceAll('"', '""')}"`;
  }
};

const formulaRefSheet = (
  formula: string,
  ref: FormulaRef,
  wb: WorkbookHandle,
  activeSheet: number,
): number | null => {
  const raw = formula.slice(ref.start, ref.end);
  const prefix = /^(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_]*))!/.exec(raw);
  const before = formula.slice(0, ref.start);
  // The reference lexer can see the local-looking suffix of an external
  // reference. Leave those references untouched rather than reading a cell
  // from the active workbook by accident.
  if (before.endsWith(']') || raw.includes('[') || raw.includes(']')) return null;
  if (!prefix) return activeSheet;
  const requested = (prefix[1] ?? prefix[2] ?? '').replaceAll("''", "'");
  for (let sheet = 0; sheet < wb.sheetCount; sheet += 1) {
    if (wb.sheetName(sheet).toLocaleLowerCase() === requested.toLocaleLowerCase()) return sheet;
  }
  return null;
};

const formulaRefValueText = (
  wb: WorkbookHandle | null,
  formula: string,
  activeSheet: number,
  ref: FormulaRef,
): string | null => {
  if (!wb) return null;
  const sheet = formulaRefSheet(formula, ref, wb, activeSheet);
  if (sheet === null) return null;
  const value = wb.getValue({ sheet, row: ref.r0, col: ref.c0 });
  // Reference substitution is structural work on the canonical formula. A
  // localized decimal separator, boolean spelling, or function alias is
  // applied only after all offsets have been consumed for display.
  return canonicalValueText(value);
};

const evaluatedFormulaText = (
  formula: string,
  refs: FormulaRef[],
  steps: number,
  wb: WorkbookHandle | null,
  sheet: number,
): string => {
  if (steps <= 0 || refs.length === 0) return formula;
  const groups = new Map<string, { text: string; refs: FormulaRef[] }>();
  for (const ref of refs) {
    const key = formulaRefKey(ref, formula);
    let group = groups.get(key);
    if (!group) {
      if (groups.size >= steps) continue;
      group = {
        text: formulaRefValueText(wb, formula, sheet, ref) ?? formula.slice(ref.start, ref.end),
        refs: [],
      };
      groups.set(key, group);
    }
    group.refs.push(ref);
  }
  const replacements = [...groups.values()].flatMap((group) =>
    group.refs.map((ref) => ({ start: ref.start, end: ref.end, text: group.text })),
  );
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

  shell.overlay.classList.add('fc-fmtdlg');
  const { body, footer } = appendDialogFrame(shell, {
    title: t.title,
    panelClasses: ['fc-fmtdlg__panel'],
    headerClass: 'fc-fmtdlg__header fc-evaldlg__header',
    bodyClass: 'fc-fmtdlg__body fc-evaldlg__body',
    footerClass: 'fc-fmtdlg__footer fc-evaldlg__footer',
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
    baseClass: 'fc-fmtdlg__btn fc-evaldlg__btn',
  });
  const closeBtn = appendDialogButton(footer, {
    label: t.close,
    variant: 'primary',
    baseClass: 'fc-fmtdlg__btn fc-evaldlg__btn',
    primaryClass: 'fc-fmtdlg__btn--primary fc-evaldlg__btn--primary',
  });

  let currentFormula = '';
  let currentRefs: FormulaRef[] = [];
  let stepIndex = 0;

  const displayLocale = (): string => {
    const wb = getWb();
    return (
      deps.getLocale?.() ??
      (wb ? inputLocaleForWorkbook(wb) : undefined) ??
      bundledLocaleForStrings(strings) ??
      'en-US'
    );
  };

  const formulaForDisplay = (formula: string, wb: WorkbookHandle | null): string => {
    if (!wb || typeof wb.getFormulaForDisplay !== 'function') return formula;
    try {
      return wb.getFormulaForDisplay(formula);
    } catch {
      return formula;
    }
  };

  const valueForDisplay = (
    wb: WorkbookHandle | null,
    addr: { sheet: number; row: number; col: number },
  ): string => {
    if (!wb) return formatCell({ kind: 'blank' }, displayLocale());
    try {
      if (typeof wb.getDisplayText === 'function') {
        const native = wb.getDisplayText(addr.sheet, addr.row, addr.col);
        if (native !== null) return native;
      }
      return formatCell(wb.getValue(addr), displayLocale());
    } catch {
      return '';
    }
  };

  const uniqueRefCount = (formula: string, refs: FormulaRef[]): number =>
    new Set(refs.map((ref) => formulaRefKey(ref, formula))).size;

  const setEvaluateDisabled = (disabled: boolean, reason: string | null): void => {
    projectDisabledState(evalBtn, disabled, reason, {
      datasetKey: 'disabledReason',
      titlePrefix: t.evaluate,
    });
  };

  const syncEvaluation = (): void => {
    const wb = getWb();
    const active = store.getState().selection.active;
    const canonicalEvaluation = evaluatedFormulaText(
      currentFormula,
      currentRefs,
      stepIndex,
      wb,
      active.sheet,
    );
    evaluationBox.textContent = formulaForDisplay(canonicalEvaluation, wb);
    const uniqueRefs = uniqueRefCount(currentFormula, currentRefs);
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
    target.textContent = formatA1Cell(active.row, active.col);
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
    formulaBox.textContent = formulaForDisplay(formula, wb);
    syncEvaluation();
    resultBox.textContent = valueForDisplay(wb, active);
  };

  shell.on(evalBtn, 'click', () => {
    if (stepIndex < uniqueRefCount(currentFormula, currentRefs)) stepIndex += 1;
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
