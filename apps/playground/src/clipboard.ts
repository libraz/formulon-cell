import {
  applyPasteSpecial,
  type ClipboardSnapshot,
  captureSnapshot,
  copy,
  cut,
  mutators,
  type PasteOperation,
  type PasteSpecialOptions,
  type PasteWhat,
  pasteTSV,
  recordFormatChange,
  type SpreadsheetInstance,
} from '@libraz/formulon-cell';
import { showMessage } from './dialogs.js';

export interface ClipboardCtx {
  getInst: () => SpreadsheetInstance | null;
  refreshWorkbookCells: () => void;
  focusSheet: () => void;
  ribbonLang: 'ja' | 'en';
}

export interface ClipboardApi {
  copySelectionToClipboard: () => Promise<void>;
  cutSelectionToClipboard: () => Promise<void>;
  pasteClipboardIntoSelection: () => Promise<void>;
  applyRibbonPasteAction: (action: string) => Promise<void>;
}

export const createClipboard = (ctx: ClipboardCtx): ClipboardApi => {
  const { getInst, refreshWorkbookCells, focusSheet, ribbonLang } = ctx;

  // Module-local clipboard state — owned by this module so main.ts no longer
  // needs to track it.
  let ribbonClipboardSnapshot: ClipboardSnapshot | null = null;
  let ribbonClipboardText: string | null = null;

  const copySelectionToClipboard = async (): Promise<void> => {
    const inst = getInst();
    if (!inst) return;
    const state = inst.store.getState();
    const result = copy(state);
    if (!result) return;
    ribbonClipboardSnapshot = captureSnapshot(state, result.range);
    ribbonClipboardText = result.tsv;
    await navigator.clipboard?.writeText(result.tsv);
    focusSheet();
  };

  const cutSelectionToClipboard = async (): Promise<void> => {
    const inst = getInst();
    if (!inst) return;
    const state = inst.store.getState();
    ribbonClipboardSnapshot = captureSnapshot(state, state.selection.range);
    inst.history.begin();
    let result: ReturnType<typeof cut> = null;
    try {
      result = cut(state, inst.workbook);
      if (result) {
        const ranges = result.payloadRanges ?? result.ranges ?? [result.range];
        recordFormatChange(inst.history, inst.store, () => {
          inst?.store.setState((s) => {
            const formats = new Map(s.format.formats);
            for (const range of ranges) {
              for (let row = range.r0; row <= range.r1; row += 1) {
                for (let col = range.c0; col <= range.c1; col += 1) {
                  formats.delete(`${range.sheet}:${row}:${col}`);
                }
              }
            }
            return { ...s, format: { formats } };
          });
        });
      }
    } finally {
      inst.history.end();
    }
    if (!result) {
      ribbonClipboardSnapshot = null;
      ribbonClipboardText = null;
      return;
    }
    ribbonClipboardText = result.tsv;
    await navigator.clipboard?.writeText(result.tsv);
    refreshWorkbookCells();
    focusSheet();
  };

  const pasteClipboardIntoSelection = async (): Promise<void> => {
    const i = getInst();
    if (!i) return;
    let text = '';
    try {
      text = (await navigator.clipboard?.readText()) ?? '';
    } catch {
      text = '';
    }
    if (!text && ribbonClipboardText) text = ribbonClipboardText;
    if (!text) return;
    if (ribbonClipboardSnapshot && text === ribbonClipboardText) {
      const source = ribbonClipboardSnapshot;
      i.history.begin();
      let result: ReturnType<typeof applyPasteSpecial> = null;
      try {
        recordFormatChange(i.history, i.store, () => {
          result = applyPasteSpecial(i.store.getState(), i.store, i.workbook, source, {
            what: 'all',
            operation: 'none',
            skipBlanks: false,
            transpose: false,
          });
        });
      } finally {
        i.history.end();
      }
      const applied = result as ReturnType<typeof applyPasteSpecial>;
      if (applied) mutators.setRange(i.store, applied.writtenRange);
    } else {
      i.history.begin();
      let result: ReturnType<typeof pasteTSV> = null;
      try {
        result = pasteTSV(i.store.getState(), i.workbook, text);
      } finally {
        i.history.end();
      }
      if (result) mutators.setRange(i.store, result.writtenRange);
    }
    refreshWorkbookCells();
    focusSheet();
  };

  const pasteOptionsForAction = (action: string): PasteSpecialOptions | null => {
    const base = {
      operation: 'none' as PasteOperation,
      skipBlanks: false,
      transpose: false,
    };
    const whatByAction: Record<string, PasteWhat> = {
      all: 'all',
      formulas: 'formulas',
      'formulas-and-numfmt': 'formulas-and-numfmt',
      values: 'values',
      'values-and-numfmt': 'values-and-numfmt',
      formats: 'formats',
    };
    if (action === 'transpose') return { ...base, what: 'all', transpose: true };
    const what = whatByAction[action];
    return what ? { ...base, what } : null;
  };

  const ribbonPasteWhatOptions = (): Array<{ value: PasteWhat; label: string }> =>
    ribbonLang === 'ja'
      ? [
          { value: 'all', label: 'すべて' },
          { value: 'formulas', label: '数式' },
          { value: 'values', label: '値' },
          { value: 'formats', label: '書式' },
          { value: 'formulas-and-numfmt', label: '数式と数値の書式' },
          { value: 'values-and-numfmt', label: '値と数値の書式' },
        ]
      : [
          { value: 'all', label: 'All' },
          { value: 'formulas', label: 'Formulas' },
          { value: 'values', label: 'Values' },
          { value: 'formats', label: 'Formats' },
          { value: 'formulas-and-numfmt', label: 'Formulas and number formats' },
          { value: 'values-and-numfmt', label: 'Values and number formats' },
        ];

  const ribbonPasteOperationOptions = (): Array<{ value: PasteOperation; label: string }> =>
    ribbonLang === 'ja'
      ? [
          { value: 'none', label: 'しない' },
          { value: 'add', label: '加算' },
          { value: 'subtract', label: '減算' },
          { value: 'multiply', label: '乗算' },
          { value: 'divide', label: '除算' },
        ]
      : [
          { value: 'none', label: 'None' },
          { value: 'add', label: 'Add' },
          { value: 'subtract', label: 'Subtract' },
          { value: 'multiply', label: 'Multiply' },
          { value: 'divide', label: 'Divide' },
        ];

  const makeRibbonPasteRadio = <T extends string>(
    name: string,
    value: T,
    label: string,
    checked: boolean,
  ): HTMLLabelElement => {
    const wrap = document.createElement('label');
    wrap.className = 'fc-fmtdlg__radio';
    const input = document.createElement('input');
    input.type = 'radio';
    input.name = name;
    input.value = value;
    input.checked = checked;
    const span = document.createElement('span');
    span.textContent = label;
    wrap.append(input, span);
    return wrap;
  };

  const makeRibbonPasteCheck = (
    label: string,
  ): { input: HTMLInputElement; label: HTMLLabelElement } => {
    const wrap = document.createElement('label');
    wrap.className = 'fc-fmtdlg__check';
    const input = document.createElement('input');
    input.type = 'checkbox';
    const span = document.createElement('span');
    span.textContent = label;
    wrap.append(input, span);
    return { input, label: wrap };
  };

  const selectedRibbonPasteRadio = <T extends string>(
    root: HTMLElement,
    name: string,
    fallback: T,
  ): T =>
    (root.querySelector<HTMLInputElement>(`input[name="${name}"]:checked`)?.value as
      | T
      | undefined) ?? fallback;

  const applyRibbonPasteSpecialSnapshot = (
    source: ClipboardSnapshot,
    opts: PasteSpecialOptions,
  ): boolean => {
    const i = getInst();
    if (!i) return false;
    i.history.begin();
    let result: ReturnType<typeof applyPasteSpecial> = null;
    try {
      recordFormatChange(i.history, i.store, () => {
        result = applyPasteSpecial(i.store.getState(), i.store, i.workbook, source, opts);
      });
    } finally {
      i.history.end();
    }
    const applied = result as ReturnType<typeof applyPasteSpecial>;
    if (!applied) return false;
    mutators.setRange(i.store, applied.writtenRange);
    refreshWorkbookCells();
    focusSheet();
    return true;
  };

  const openRibbonPasteSpecialDialog = (source: ClipboardSnapshot): void => {
    const ja = ribbonLang === 'ja';
    const title = ja ? '形式を選択して貼り付け' : 'Paste Special';
    const opener = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    const overlay = document.createElement('div');
    overlay.className = 'fc-fmtdlg app__dlg fc-pastesp';
    overlay.setAttribute('role', 'dialog');
    overlay.setAttribute('aria-modal', 'true');
    overlay.setAttribute('aria-label', title);

    const panel = document.createElement('div');
    panel.className = 'fc-fmtdlg__panel app__dlg__panel fc-pastesp__panel';
    overlay.appendChild(panel);

    const header = document.createElement('div');
    header.className = 'fc-fmtdlg__header';
    header.textContent = title;
    panel.appendChild(header);

    const body = document.createElement('div');
    body.className = 'fc-fmtdlg__body fc-pastesp__body';
    panel.appendChild(body);

    const cols = document.createElement('div');
    cols.className = 'fc-pastesp__cols';
    body.appendChild(cols);

    const whatName = `app-ribbon-paste-what-${Math.random().toString(36).slice(2)}`;
    const whatGroup = document.createElement('div');
    whatGroup.className = 'fc-pastesp__group';
    const whatLegend = document.createElement('div');
    whatLegend.className = 'fc-pastesp__legend';
    whatLegend.textContent = ja ? '貼り付け' : 'Paste';
    const whatList = document.createElement('div');
    whatList.className = 'fc-pastesp__list';
    whatList.setAttribute('role', 'radiogroup');
    whatList.setAttribute('aria-label', whatLegend.textContent);
    for (const option of ribbonPasteWhatOptions()) {
      whatList.appendChild(
        makeRibbonPasteRadio(whatName, option.value, option.label, option.value === 'all'),
      );
    }
    whatGroup.append(whatLegend, whatList);
    cols.appendChild(whatGroup);

    const opName = `app-ribbon-paste-op-${Math.random().toString(36).slice(2)}`;
    const opGroup = document.createElement('div');
    opGroup.className = 'fc-pastesp__group';
    const opLegend = document.createElement('div');
    opLegend.className = 'fc-pastesp__legend';
    opLegend.textContent = ja ? '演算' : 'Operation';
    const opList = document.createElement('div');
    opList.className = 'fc-pastesp__list';
    opList.setAttribute('role', 'radiogroup');
    opList.setAttribute('aria-label', opLegend.textContent);
    for (const option of ribbonPasteOperationOptions()) {
      opList.appendChild(
        makeRibbonPasteRadio(opName, option.value, option.label, option.value === 'none'),
      );
    }
    opGroup.append(opLegend, opList);
    cols.appendChild(opGroup);

    const bottomRow = document.createElement('div');
    bottomRow.className = 'fc-pastesp__bottomrow';
    const skipBlanks = makeRibbonPasteCheck(ja ? '空白セルを無視する' : 'Skip blanks');
    const transpose = makeRibbonPasteCheck(ja ? '行/列の入れ替え' : 'Transpose');
    bottomRow.append(skipBlanks.label, transpose.label);
    body.appendChild(bottomRow);

    const footer = document.createElement('div');
    footer.className = 'fc-fmtdlg__footer';
    panel.appendChild(footer);
    const cancelBtn = document.createElement('button');
    cancelBtn.type = 'button';
    cancelBtn.className = 'fc-fmtdlg__btn';
    cancelBtn.textContent = ja ? 'キャンセル' : 'Cancel';
    const okBtn = document.createElement('button');
    okBtn.type = 'button';
    okBtn.className = 'fc-fmtdlg__btn fc-fmtdlg__btn--primary';
    okBtn.textContent = 'OK';
    footer.append(cancelBtn, okBtn);

    const close = (): void => {
      overlay.removeEventListener('keydown', onKey);
      overlay.remove();
      opener?.focus({ preventScroll: true });
    };
    const apply = (): void => {
      const what = selectedRibbonPasteRadio<PasteWhat>(overlay, whatName, 'all');
      const operation = selectedRibbonPasteRadio<PasteOperation>(overlay, opName, 'none');
      applyRibbonPasteSpecialSnapshot(source, {
        what,
        operation,
        skipBlanks: skipBlanks.input.checked,
        transpose: transpose.input.checked,
      });
      close();
    };
    const onKey = (event: KeyboardEvent): void => {
      event.stopPropagation();
      if (event.key === 'Escape') {
        event.preventDefault();
        close();
      } else if (event.key === 'Enter') {
        event.preventDefault();
        apply();
      }
    };
    cancelBtn.addEventListener('click', close);
    okBtn.addEventListener('click', apply);
    overlay.addEventListener('keydown', onKey);
    overlay.addEventListener('click', (event) => {
      if (event.target === overlay) close();
    });
    document.body.appendChild(overlay);
    requestAnimationFrame(() => {
      whatList.querySelector<HTMLInputElement>('input[type="radio"]')?.focus();
    });
  };

  const applyRibbonPasteAction = async (action: string): Promise<void> => {
    const i = getInst();
    if (!i) return;
    if (action === 'dialog') {
      if (ribbonClipboardSnapshot) {
        openRibbonPasteSpecialDialog(ribbonClipboardSnapshot);
        return;
      }
      i.openPasteSpecial();
      return;
    }
    const opts = pasteOptionsForAction(action);
    if (!opts) return;
    let text = '';
    try {
      text = (await navigator.clipboard?.readText()) ?? '';
    } catch {
      text = '';
    }
    if (!text && ribbonClipboardText) text = ribbonClipboardText;
    if (ribbonClipboardSnapshot && text === ribbonClipboardText) {
      const source = ribbonClipboardSnapshot;
      applyRibbonPasteSpecialSnapshot(source, opts);
      return;
    }
    if (action === 'all' || action === 'values') {
      if (!text) return;
      let result: ReturnType<typeof pasteTSV> = null;
      i.history.begin();
      try {
        result = pasteTSV(i.store.getState(), i.workbook, text);
      } finally {
        i.history.end();
      }
      if (result) mutators.setRange(i.store, result.writtenRange);
      refreshWorkbookCells();
      focusSheet();
      return;
    }
    void showMessage({
      title: ribbonLang === 'ja' ? '貼り付け' : 'Paste',
      message:
        ribbonLang === 'ja'
          ? 'この貼り付け形式には、このブック内でコピーしたセルが必要です。'
          : 'This paste option requires cells copied inside this workbook.',
    });
  };

  return {
    copySelectionToClipboard,
    cutSelectionToClipboard,
    pasteClipboardIntoSelection,
    applyRibbonPasteAction,
  };
};
