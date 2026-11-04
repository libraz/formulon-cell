import type { FormulaDialogsStrings } from '../formula-dialogs/types.js';
import type { WorkbookChromeStrings } from '../workbook-chrome/types.js';

/** Accessible names the Mac formula bar and status bar use. */
export interface MacChromeStrings {
  cancelFormulaEdit: string;
  enterFormula: string;
  /** The fx button, which opens the Formula Builder on the Mac host. */
  fxButton: string;
  expandFormulaBar: string;
  pageBreakPreview: string;
}

/** A locale whose Mac wording is not measured repeats its own shared text. */
export const macChromeFrom = (
  chrome: WorkbookChromeStrings,
  formula: FormulaDialogsStrings,
): MacChromeStrings => ({
  cancelFormulaEdit: chrome.a11y.cancelFormulaEdit,
  enterFormula: chrome.a11y.enterFormula,
  fxButton: formula.fxDialog.fxButtonLabel,
  expandFormulaBar: chrome.a11y.expandFormulaBar,
  pageBreakPreview: chrome.statusBar.pageBreakPreview,
});

/** Japanese Mac app wording, measured from Excel for Mac. */
export const macChromeJa: MacChromeStrings = {
  cancelFormulaEdit: '数式をキャンセルします',
  enterFormula: '数式を入力します',
  fxButton: '数式パレットを開きます',
  expandFormulaBar: '数式バーの展開',
  pageBreakPreview: '改ページのプレビュー',
};
