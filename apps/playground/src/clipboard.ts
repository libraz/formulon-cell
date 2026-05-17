import {
  dispatchHostClipboard,
  type PasteOperation,
  type PasteSpecialOptions,
  type PasteWhat,
  type SpreadsheetInstance,
} from '@libraz/formulon-cell';
import { showMessage } from './dialogs.js';

// Phase 1.5 collapsed the previous ribbon-side snapshot tracking into core's
// `instance.clipboard` handle. This module is now a thin glue that routes
// ribbon Copy/Cut/Paste clicks through `dispatchHostClipboard` (which prefers
// `runShortcut`) and translates the Paste sub-menu actions into either
// `instance.pasteSpecial(opts)` or the same host shortcut for plain text
// pastes from outside the workbook.

export interface ClipboardCtx {
  getInst: () => SpreadsheetInstance | null;
  focusSheet: () => void;
  ribbonLang: 'ja' | 'en';
}

export interface ClipboardApi {
  copySelectionToClipboard: () => Promise<void>;
  cutSelectionToClipboard: () => Promise<void>;
  pasteClipboardIntoSelection: () => Promise<void>;
  applyRibbonPasteAction: (action: string) => Promise<void>;
}

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

export const createClipboard = (ctx: ClipboardCtx): ClipboardApi => {
  const { getInst, focusSheet, ribbonLang } = ctx;

  const copySelectionToClipboard = async (): Promise<void> => {
    dispatchHostClipboard(getInst(), 'copy');
    focusSheet();
  };

  const cutSelectionToClipboard = async (): Promise<void> => {
    dispatchHostClipboard(getInst(), 'cut');
    focusSheet();
  };

  const pasteClipboardIntoSelection = async (): Promise<void> => {
    dispatchHostClipboard(getInst(), 'paste');
    focusSheet();
  };

  const applyRibbonPasteAction = async (action: string): Promise<void> => {
    const i = getInst();
    if (!i) return;
    if (action === 'dialog') {
      i.openPasteSpecial();
      return;
    }
    const opts = pasteOptionsForAction(action);
    if (!opts) return;
    if (i.pasteSpecial(opts)) {
      focusSheet();
      return;
    }
    // No internal snapshot — fall back to the host shortcut for the two
    // actions that have a sensible plain-text equivalent.
    if (action === 'all' || action === 'values') {
      dispatchHostClipboard(i, 'paste');
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
