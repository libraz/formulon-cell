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
  /** Status-bar button that recalculates a manual-mode workbook holding stale values. */
  calculate: string;
  /** Help text of that button. */
  calculateHint: string;
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
  calculate: chrome.statusBar.calcLabel,
  calculateHint: chrome.statusBar.calcRecalcHint,
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
  calculate: '再計算',
  calculateHint:
    '自動計算を行うように設定されていないため、数式の計算結果が最新の状態でない可能性があります。手動で再計算するには、ここをクリックするか、F9 キーを押してください。',
};

/** Mac app wording from the Excel bundle; not answered by it: capsLock. */
export const macChromeDe: Partial<MacChromeStrings> = {
  cancelFormulaEdit: 'Formel abbrechen',
  enterFormula: 'Formel eingeben',
  fxButton: 'Formel-Generator öffnen',
  expandFormulaBar: 'Bearbeitungsleiste erweitern',
  pageBreakPreview: 'Seitenumbruchvorschau',
  statusBarCustomize: 'Statusleiste anpassen',
  calculate: 'Berechnen',
  calculateHint:
    'Die Ergebnisse von Formeln sind möglicherweise veraltet, weil nicht die automatische Berechnung festgelegt ist. Klicken Sie hier, oder drücken Sie F9, um sie manuell neu zu berechnen.',
};

/** Mac app wording from the Excel bundle; not answered by it: expandFormulaBar, capsLock. */
export const macChromeFr: Partial<MacChromeStrings> = {
  cancelFormulaEdit: 'annuler une formule',
  enterFormula: 'entrer une formule',
  fxButton: 'ouvrir le concepteur de formule',
  pageBreakPreview: 'Aperçu des sauts de page',
  statusBarCustomize: 'Personnaliser la barre d’état',
  calculate: 'Calculer',
  calculateHint:
    "Les résultats des formules sont peut-être obsolètes, car le calcul automatique n'est pas activé. Cliquez ici ou appuyez sur F9 pour recalculer manuellement.",
};

/** Mac app wording from the Excel bundle; not answered by it: capsLock. */
export const macChromeKo: Partial<MacChromeStrings> = {
  cancelFormulaEdit: '수식 취소',
  enterFormula: '수식 입력',
  fxButton: '수식 작성기 열기',
  expandFormulaBar: '수식 입력줄 확장',
  pageBreakPreview: '페이지 나누기 미리 보기',
  statusBarCustomize: '상태 표시줄 사용자 지정',
  calculate: '계산',
  calculateHint:
    '자동으로 계산하도록 설정되어 있지 않으므로 수식의 결과가 오래되었을 수 있습니다. 여기를 클릭하거나 <F9> 키를 눌러 수동으로 다시 계산하세요.',
};

/** Mac app wording from the Excel bundle; not answered by it: capsLock. */
export const macChromeTh: Partial<MacChromeStrings> = {
  cancelFormulaEdit: 'ยกเลิกสูตร',
  enterFormula: 'ใส่สูตร',
  fxButton: 'เปิดตัวสร้างสูตร',
  expandFormulaBar: 'ขยายแถบสูตร',
  pageBreakPreview: 'แสดงตัวอย่างตัวแบ่งหน้า',
  statusBarCustomize: 'กำหนดแถบสถานะเอง',
  calculate: 'คำนวณ',
  calculateHint:
    'ผลลัพธ์ของสูตรอาจเป็นข้อมูลเก่า เนื่องจากไม่ได้ถูกตั้งให้คำนวณอัตโนมัติ คลิกที่นี่หรือกด F9 เพื่อคำนวณใหม่ด้วยตนเอง',
};

/** Mac app wording from the Excel bundle; not answered by it: capsLock. */
export const macChromeZh: Partial<MacChromeStrings> = {
  cancelFormulaEdit: '取消公式',
  enterFormula: '输入公式',
  fxButton: '打开公式生成器',
  expandFormulaBar: '展开编辑栏',
  pageBreakPreview: '分页预览',
  statusBarCustomize: '自定义状态栏',
  calculate: '计算',
  calculateHint: '公式结果可能已过期，因为未被设置为自动计算。请单击此处或按 F9 手动重新计算。',
};

/** Mac app wording from the Excel bundle. */
export const macChromeEn: Partial<MacChromeStrings> = {
  enterFormula: 'enter formula',
  expandFormulaBar: 'Expand Formula Bar',
  cancelFormulaEdit: 'cancel formula',
  fxButton: 'open the formula builder',
  calculate: 'Calculate',
  calculateHint:
    'Results of formulas may be out of date because the workbook is not set to calculate automatically. Click here or press F9 to recalculate manually.',
};
