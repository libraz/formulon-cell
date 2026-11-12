/** The Mac sheet-tab context menu: the entries whose wording differs from the shared tab menu. */
export interface MacSheetTabMenuStrings {
  insertSheet: string;
  tabColor: string;
  hideSheet: string;
  unhideSheet: string;
  /** Title of the dialog that lists the hidden sheets. */
  unhideTitle: string;
  /** Caption above that dialog's list. */
  unhideList: string;
  viewCode: string;
  selectAllSheets: string;
  /** Why the entries without a host feature cannot be used. */
  unavailable: string;
}

/** Mac wording for this locale; not measured. */
export const macSheetTabMenuEn: MacSheetTabMenuStrings = {
  insertSheet: 'Insert Sheet',
  tabColor: 'Tab Color',
  hideSheet: 'Hide',
  unhideSheet: 'Unhide...',
  unhideTitle: 'Unhide',
  unhideList: 'Unhide sheet:',
  viewCode: 'View Code',
  selectAllSheets: 'Select All Sheets',
  unavailable: 'This feature is not implemented yet.',
};

/** Japanese Mac app wording, measured from Excel for Mac. */
export const macSheetTabMenuJa: MacSheetTabMenuStrings = {
  insertSheet: 'シートの挿入 (⇧F11)',
  tabColor: 'シート見出しの色',
  hideSheet: '表示しない',
  unhideSheet: '再表示...',
  unhideTitle: '再表示',
  unhideList: '再表示するシート:',
  viewCode: 'コードの表示',
  selectAllSheets: 'すべてのシートを選択',
  unavailable: 'この機能は未実装です。',
};

/** Mac app wording from the Excel bundle; not answered by it: insertSheet, unhideList, unavailable. */
export const macSheetTabMenuDe: MacSheetTabMenuStrings = {
  insertSheet: 'Blatt einfügen',
  tabColor: 'Registerfarbe',
  hideSheet: 'Ausblenden',
  unhideSheet: 'Einblenden...',
  unhideTitle: 'Einblenden',
  unhideList: 'Blatt einblenden:',
  viewCode: 'Code anzeigen',
  selectAllSheets: 'Alle Blätter markieren',
  unavailable: 'Diese Funktion ist noch nicht implementiert.',
};

/** Mac app wording from the Excel bundle; not answered by it: insertSheet, unhideList, unavailable. */
export const macSheetTabMenuFr: MacSheetTabMenuStrings = {
  insertSheet: 'Insérer une feuille',
  tabColor: 'Couleur d’onglet',
  hideSheet: 'Masquer',
  unhideSheet: 'Afficher...',
  unhideTitle: 'Afficher',
  unhideList: 'Afficher la feuille :',
  viewCode: 'Afficher le code',
  selectAllSheets: 'Sélectionner toutes les feuilles',
  unavailable: 'Cette fonctionnalité n’est pas encore implémentée.',
};

/** Mac app wording from the Excel bundle; not answered by it: insertSheet, unhideList, unavailable. */
export const macSheetTabMenuZh: MacSheetTabMenuStrings = {
  insertSheet: '插入工作表',
  tabColor: '工作表标签颜色',
  hideSheet: '隐藏',
  unhideSheet: '取消隐藏...',
  unhideTitle: '取消隐藏',
  unhideList: '取消隐藏工作表:',
  viewCode: '查看代码',
  selectAllSheets: '选择所有工作表',
  unavailable: '此功能尚未实现。',
};

/** Mac app wording from the Excel bundle; not answered by it: insertSheet, unhideList, unavailable. */
export const macSheetTabMenuKo: MacSheetTabMenuStrings = {
  insertSheet: '시트 삽입',
  tabColor: '시트 탭 색',
  hideSheet: '숨기기',
  unhideSheet: '숨기기 취소...',
  unhideTitle: '숨기기 취소',
  unhideList: '숨기기를 취소할 시트:',
  viewCode: '코드 보기',
  selectAllSheets: '모든 시트 선택',
  unavailable: '이 기능은 아직 구현되지 않았습니다.',
};

/** Mac app wording from the Excel bundle; not answered by it: insertSheet, unhideTitle, unhideList, viewCode, unavailable. */
export const macSheetTabMenuTh: MacSheetTabMenuStrings = {
  insertSheet: 'แทรกแผ่นงาน',
  tabColor: 'สีแท็บ',
  hideSheet: 'ซ่อน',
  unhideSheet: 'ยกเลิกการซ่อน...',
  unhideTitle: 'ยกเลิกการซ่อน',
  unhideList: 'ยกเลิกการซ่อนแผ่นงาน:',
  viewCode: 'ดูโค้ด',
  selectAllSheets: 'เลือกแผ่นงานทั้งหมด',
  unavailable: 'ฟีเจอร์นี้ยังไม่ได้ใช้งาน',
};
