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

// The core paste-special dialog (instance.openPasteSpecial / instance.pasteSpecial)
// historically read its snapshot from the in-grid clipboard handle. The ribbon
// flow below keeps its own snapshot (populated by the ribbon Copy/Cut buttons),
// so we feed it back to core via the optional `{ snapshot }` override rather
// than duplicating the entire dialog DOM inside playground.

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

  const applyRibbonPasteAction = async (action: string): Promise<void> => {
    const i = getInst();
    if (!i) return;
    if (action === 'dialog') {
      i.openPasteSpecial(
        ribbonClipboardSnapshot ? { snapshot: ribbonClipboardSnapshot } : undefined,
      );
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
      const applied = i.pasteSpecial(opts, { snapshot: ribbonClipboardSnapshot });
      if (applied) focusSheet();
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
