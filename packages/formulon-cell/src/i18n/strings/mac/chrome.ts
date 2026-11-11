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
  /** Title row of the status bar's right-click menu. */
  statusBarCustomize: string;
  capsLock: string;
}

/** Each locale starts from its own shared text and overrides the fields the Excel bundle answers. */
export const macChromeFrom = (
  chrome: WorkbookChromeStrings,
  formula: FormulaDialogsStrings,
): MacChromeStrings => ({
  cancelFormulaEdit: chrome.a11y.cancelFormulaEdit,
  enterFormula: chrome.a11y.enterFormula,
  fxButton: formula.fxDialog.fxButtonLabel,
  expandFormulaBar: chrome.a11y.expandFormulaBar,
  pageBreakPreview: chrome.statusBar.pageBreakPreview,
  statusBarCustomize: chrome.statusBar.aggregatesHeading,
  capsLock: chrome.statusBar.capsLock,
});

/** Japanese Mac app wording, measured from Excel for Mac. */
export const macChromeJa: MacChromeStrings = {
  cancelFormulaEdit: '数式をキャンセルします',
  enterFormula: '数式を入力します',
  fxButton: '数式パレットを開きます',
  expandFormulaBar: '数式バーの展開',
  pageBreakPreview: '改ページのプレビュー',
  statusBarCustomize: 'ステータス バーのユーザー設定',
  capsLock: 'CapsLock',
};

/** Mac app wording from the Excel bundle; not answered by it: cancelFormulaEdit, fxButton, capsLock. */
export const macChromeDe: Partial<MacChromeStrings> = {
  enterFormula: 'Formel eingeben',
  expandFormulaBar: 'Bearbeitungsleiste erweitern',
  pageBreakPreview: 'Seitenumbruchvorschau',
  statusBarCustomize: 'Statusleiste anpassen',
};

/** Mac app wording from the Excel bundle; not answered by it: cancelFormulaEdit, fxButton, expandFormulaBar, capsLock. */
export const macChromeFr: Partial<MacChromeStrings> = {
  enterFormula: 'entrer une formule',
  pageBreakPreview: 'Aperçu des sauts de page',
  statusBarCustomize: 'Personnaliser la barre d’état',
};

/** Mac app wording from the Excel bundle; not answered by it: cancelFormulaEdit, fxButton, capsLock. */
export const macChromeKo: Partial<MacChromeStrings> = {
  enterFormula: '수식 입력',
  expandFormulaBar: '수식 입력줄 확장',
  pageBreakPreview: '페이지 나누기 미리 보기',
  statusBarCustomize: '상태 표시줄 사용자 지정',
};

/** Mac app wording from the Excel bundle; not answered by it: cancelFormulaEdit, fxButton, capsLock. */
export const macChromeTh: Partial<MacChromeStrings> = {
  enterFormula: 'ใส่สูตร',
  expandFormulaBar: 'ขยายแถบสูตร',
  pageBreakPreview: 'แสดงตัวอย่างตัวแบ่งหน้า',
  statusBarCustomize: 'กำหนดแถบสถานะเอง',
};

/** Mac app wording from the Excel bundle; not answered by it: cancelFormulaEdit, fxButton, capsLock. */
export const macChromeZh: Partial<MacChromeStrings> = {
  enterFormula: '输入公式',
  expandFormulaBar: '展开编辑栏',
  pageBreakPreview: '分页预览',
  statusBarCustomize: '自定义状态栏',
};
