/** The Mac sheet-tab context menu: the entries whose wording differs from the shared tab menu. */
export interface MacSheetTabMenuStrings {
  insertSheet: string;
  tabColor: string;
  hideSheet: string;
  unhideSheet: string;
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
  viewCode: 'コードの表示',
  selectAllSheets: 'すべてのシートを選択',
  unavailable: 'この機能は未実装です。',
};

/** Mac wording for this locale; not measured. */
export const macSheetTabMenuDe: MacSheetTabMenuStrings = {
  insertSheet: 'Blatt einfügen',
  tabColor: 'Registerkartenfarbe',
  hideSheet: 'Ausblenden',
  unhideSheet: 'Einblenden...',
  viewCode: 'Code anzeigen',
  selectAllSheets: 'Alle Blätter auswählen',
  unavailable: 'Diese Funktion ist noch nicht implementiert.',
};

/** Mac wording for this locale; not measured. */
export const macSheetTabMenuFr: MacSheetTabMenuStrings = {
  insertSheet: 'Insérer une feuille',
  tabColor: 'Couleur de l’onglet',
  hideSheet: 'Masquer',
  unhideSheet: 'Afficher...',
  viewCode: 'Afficher le code',
  selectAllSheets: 'Sélectionner toutes les feuilles',
  unavailable: 'Cette fonctionnalité n’est pas encore implémentée.',
};

/** Mac wording for this locale; not measured. */
export const macSheetTabMenuZh: MacSheetTabMenuStrings = {
  insertSheet: '插入工作表',
  tabColor: '选项卡颜色',
  hideSheet: '隐藏',
  unhideSheet: '取消隐藏...',
  viewCode: '查看代码',
  selectAllSheets: '选定全部工作表',
  unavailable: '此功能尚未实现。',
};

/** Mac wording for this locale; not measured. */
export const macSheetTabMenuKo: MacSheetTabMenuStrings = {
  insertSheet: '시트 삽입',
  tabColor: '탭 색',
  hideSheet: '숨기기',
  unhideSheet: '숨기기 취소...',
  viewCode: '코드 보기',
  selectAllSheets: '모든 시트 선택',
  unavailable: '이 기능은 아직 구현되지 않았습니다.',
};

/** Mac wording for this locale; not measured. */
export const macSheetTabMenuTh: MacSheetTabMenuStrings = {
  insertSheet: 'แทรกแผ่นงาน',
  tabColor: 'สีแท็บ',
  hideSheet: 'ซ่อน',
  unhideSheet: 'ยกเลิกการซ่อน...',
  viewCode: 'ดูโค้ด',
  selectAllSheets: 'เลือกแผ่นงานทั้งหมด',
  unavailable: 'ฟีเจอร์นี้ยังไม่ได้ใช้งาน',
};
