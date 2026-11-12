/** Wording of the PivotTable Field dialog (value and row/column fields). Function
 *  and calculation names come from `pivotTableDialog`. */
export interface MacPivotFieldSettingsStrings {
  title: string;
  /** Label of the read-only source field name row. */
  sourceName: string;
  /** Label of the editable name field. */
  customName: string;
  summarizeTab: string;
  showAsTab: string;
  baseField: string;
  baseItem: string;
  numberFormat: string;
  subtotals: string;
  subtotalAutomatic: string;
  subtotalCustom: string;
  subtotalNone: string;
  showItemsWithNoData: string;
  hide: string;
  ok: string;
  cancel: string;
  /** Reason shown on a control this build cannot apply. */
  unavailable: string;
}

/** Mac wording for this locale; not measured. */
export const macPivotFieldSettingsEn: MacPivotFieldSettingsStrings = {
  title: 'PivotTable Field',
  sourceName: 'Source name:',
  customName: 'Field name:',
  summarizeTab: 'Summarize by',
  showAsTab: 'Show data as',
  baseField: 'Base field:',
  baseItem: 'Base item:',
  numberFormat: 'Number Format…',
  subtotals: 'Subtotals',
  subtotalAutomatic: 'Automatic',
  subtotalCustom: 'Custom',
  subtotalNone: 'None',
  showItemsWithNoData: 'Show items with no data',
  hide: 'Hide',
  ok: 'OK',
  cancel: 'Cancel',
  unavailable: 'Not available in this version.',
};

/** Japanese Mac app wording, measured from Excel for Mac. */
export const macPivotFieldSettingsJa: MacPivotFieldSettingsStrings = {
  title: 'ピボットテーブル フィールド',
  sourceName: 'フィールド名:',
  customName: 'フィールド名:',
  summarizeTab: '集計の方法',
  showAsTab: '計算の種類',
  baseField: '基準フィールド:',
  baseItem: '基準アイテム:',
  numberFormat: '表示形式...',
  subtotals: '小計',
  subtotalAutomatic: '自動',
  subtotalCustom: '指定',
  subtotalNone: 'なし',
  showItemsWithNoData: 'データのないアイテムを表示する',
  hide: '非表示',
  ok: 'OK',
  cancel: 'キャンセル',
  unavailable: 'このバージョンでは使用できません。',
};

/** Mac app wording from the Excel bundle; not answered by it: sourceName, numberFormat, unavailable. */
export const macPivotFieldSettingsDe: MacPivotFieldSettingsStrings = {
  title: 'PivotTable-Feld',
  sourceName: 'Quellname:',
  customName: 'Feldname:',
  summarizeTab: 'Zusammenfassen mit',
  showAsTab: 'Daten zeigen als',
  baseField: 'Basisfeld:',
  baseItem: 'Basiselement:',
  numberFormat: 'Zahlenformat…',
  subtotals: 'Teilergebnisse',
  subtotalAutomatic: 'Automatisch',
  subtotalCustom: 'Benutzerdefiniert',
  subtotalNone: 'Keine',
  showItemsWithNoData: 'Elemente ohne Daten anzeigen',
  hide: 'Ausblenden',
  ok: 'OK',
  cancel: 'Abbrechen',
  unavailable: 'In dieser Version nicht verfügbar.',
};

/** Mac app wording from the Excel bundle; not answered by it: sourceName, numberFormat, subtotalAutomatic, unavailable. */
export const macPivotFieldSettingsFr: MacPivotFieldSettingsStrings = {
  title: 'Champ dynamique',
  sourceName: 'Nom de la source :',
  customName: 'Nom de champ :',
  summarizeTab: 'Synthèse par',
  showAsTab: 'Afficher les données',
  baseField: 'Champ de base :',
  baseItem: 'Élément de base :',
  numberFormat: 'Format de nombre…',
  subtotals: 'Sous-totaux',
  subtotalAutomatic: 'Automatique',
  subtotalCustom: 'Personnalisés',
  subtotalNone: 'Aucun',
  showItemsWithNoData: 'Afficher les éléments sans données',
  hide: 'Masquer',
  ok: 'OK',
  cancel: 'Annuler',
  unavailable: 'Non disponible dans cette version.',
};

/** Mac app wording from the Excel bundle; not answered by it: sourceName, numberFormat, unavailable. */
export const macPivotFieldSettingsZh: MacPivotFieldSettingsStrings = {
  title: '数据透视表字段',
  sourceName: '源名称:',
  customName: '字段名称:',
  summarizeTab: '汇总方式',
  showAsTab: '数据显示方式',
  baseField: '基本字段:',
  baseItem: '基本项:',
  numberFormat: '数字格式…',
  subtotals: '小计',
  subtotalAutomatic: '自动',
  subtotalCustom: '自定义',
  subtotalNone: '无',
  showItemsWithNoData: '显示不含数据的项目',
  hide: '隐藏',
  ok: '确定',
  cancel: '取消',
  unavailable: '此版本中不可用。',
};

/** Mac app wording from the Excel bundle; not answered by it: sourceName, numberFormat, unavailable. */
export const macPivotFieldSettingsKo: MacPivotFieldSettingsStrings = {
  title: '피벗 테이블 필드',
  sourceName: '원본 이름:',
  customName: '필드 이름:',
  summarizeTab: '사용할 함수:',
  showAsTab: '데이터 표시 형식:',
  baseField: '기준 필드:',
  baseItem: '기준 항목:',
  numberFormat: '표시 형식…',
  subtotals: '소계',
  subtotalAutomatic: '자동',
  subtotalCustom: '사용자 지정',
  subtotalNone: '없음',
  showItemsWithNoData: '데이터가 없는 항목 표시',
  hide: '숨기기',
  ok: '확인',
  cancel: '취소',
  unavailable: '이 버전에서는 사용할 수 없습니다.',
};

/** Mac app wording from the Excel bundle; not answered by it: sourceName, numberFormat, unavailable. */
export const macPivotFieldSettingsTh: MacPivotFieldSettingsStrings = {
  title: 'เขตข้อมูล PivotTable',
  sourceName: 'ชื่อแหล่งข้อมูล:',
  customName: 'ชื่อเขตข้อมูล:',
  summarizeTab: 'สรุปโดย',
  showAsTab: 'แสดงข้อมูลเป็น',
  baseField: 'เขตข้อมูลพื้นฐาน:',
  baseItem: 'รายการพื้นฐาน:',
  numberFormat: 'รูปแบบตัวเลข…',
  subtotals: 'ผลรวมย่อย',
  subtotalAutomatic: 'อัตโนมัติ',
  subtotalCustom: 'กำหนดเอง',
  subtotalNone: 'ไม่มี',
  showItemsWithNoData: 'แสดงรายการที่ไม่มีข้อมูล',
  hide: 'ซ่อน',
  ok: 'ตกลง',
  cancel: 'ยกเลิก',
  unavailable: 'ไม่พร้อมใช้งานในเวอร์ชันนี้',
};
