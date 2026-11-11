import type {
  CellStyleDialogMacStrings,
  CellStyleSummaryStrings,
} from '../../../toolbar/dialogs/cell-style.js';

/** The Mac Style dialog's labels plus the wording of its format summaries. */
export type MacStyleDialogStrings = CellStyleDialogMacStrings & {
  summary: CellStyleSummaryStrings;
};

/** Mac wording for this locale; not measured. */
export const macStyleDialogEn: MacStyleDialogStrings = {
  title: 'Modify Cell Style',
  nameLabel: 'Style name:',
  previewLabel: 'Preview:',
  caption: 'Style includes (by example):',
  propertyColumn: 'Style property',
  formatColumn: 'Format',
  properties: {
    number: 'Number',
    alignment: 'Alignment',
    font: 'Font',
    border: 'Border',
    fill: 'Fill',
    protection: 'Protection',
  },
  formatButton: 'Format...',
  summary: {
    general: 'General',
    alignment: 'Alignment: {vertical}, {horizontal}',
    vertical: {
      top: 'Top',
      middle: 'Center',
      bottom: 'Bottom',
    },
    horizontal: {
      general: 'General',
      left: 'Left',
      center: 'Center',
      right: 'Right',
    },
    noBorder: 'No borders',
    border: 'Borders',
    noFill: 'No shading',
    locked: 'Locked',
    unlocked: 'Unlocked',
  },
};

/** Japanese Mac app wording, measured from Excel for Mac. */
export const macStyleDialogJa: MacStyleDialogStrings = {
  title: 'セル スタイルの変更',
  nameLabel: 'スタイル名:',
  previewLabel: 'プレビュー:',
  caption: '次のプロパティに対して、選択されたスタイルが書式設定されています:',
  propertyColumn: 'スタイル プロパティ',
  formatColumn: '書式',
  properties: {
    number: '数値',
    alignment: '配置',
    font: 'フォント',
    border: '枠線',
    fill: 'フィル',
    protection: '保護',
  },
  formatButton: '書式...',
  summary: {
    general: 'G/標準',
    alignment: '配置: {vertical}, {horizontal}',
    vertical: {
      top: '上詰め',
      middle: '中央',
      bottom: '下詰め',
    },
    horizontal: {
      general: '標準',
      left: '左詰め',
      center: '中央揃え',
      right: '右詰め',
    },
    noBorder: '罫線なし',
    border: '罫線あり',
    noFill: '網かけなし',
    locked: 'ロック',
    unlocked: 'ロックなし',
  },
};

/** Mac app wording from the Excel bundle; not answered by it: previewLabel, caption, formatColumn, properties.number, properties.alignment, properties.font, and 13 more. */
export const macStyleDialogDe: MacStyleDialogStrings = {
  title: 'Zellenformatvorlage ändern',
  nameLabel: 'Name der Formatvorlage:',
  previewLabel: 'Vorschau:',
  caption: 'Die Formatvorlage enthält folgende Eigenschaften:',
  propertyColumn: 'Formatvorlageneigenschaft',
  formatColumn: 'Format',
  properties: {
    number: 'Zahl',
    alignment: 'Ausrichtung',
    font: 'Schriftart',
    border: 'Rahmen',
    fill: 'Ausfüllen',
    protection: 'Schutz',
  },
  formatButton: 'Formatieren...',
  summary: {
    general: 'Standard',
    alignment: 'Ausrichtung: {vertical}, {horizontal}',
    vertical: {
      top: 'Oben',
      middle: 'Mitte',
      bottom: 'Unten',
    },
    horizontal: {
      general: 'Standard',
      left: 'Links',
      center: 'Zentriert',
      right: 'Rechts',
    },
    noBorder: 'Keine Rahmen',
    border: 'Rahmen',
    noFill: 'Nicht schattiert',
    locked: 'Gesperrt',
    unlocked: 'Nicht gesperrt',
  },
};

/** Mac app wording from the Excel bundle; not answered by it: previewLabel, caption, formatColumn, properties.number, properties.alignment, properties.font, and 13 more. */
export const macStyleDialogFr: MacStyleDialogStrings = {
  title: 'Modifier le style de cellule',
  nameLabel: 'Nom du style :',
  previewLabel: 'Aperçu :',
  caption: 'Le style inclut les propriétés suivantes :',
  propertyColumn: 'Propriété du style',
  formatColumn: 'Format',
  properties: {
    number: 'Nombre',
    alignment: 'Alignement',
    font: 'Police',
    border: 'Bordure',
    fill: 'Remplissage',
    protection: 'Protection',
  },
  formatButton: 'Format...',
  summary: {
    general: 'Standard',
    alignment: 'Alignement : {vertical}, {horizontal}',
    vertical: {
      top: 'Haut',
      middle: 'Centre',
      bottom: 'Bas',
    },
    horizontal: {
      general: 'Standard',
      left: 'Gauche',
      center: 'Centré',
      right: 'Droite',
    },
    noBorder: 'Aucune bordure',
    border: 'Bordures',
    noFill: "Pas d'ombrage",
    locked: 'Verrouillé',
    unlocked: 'Déverrouillée',
  },
};

/** Mac app wording from the Excel bundle; not answered by it: previewLabel, caption, formatColumn, properties.number, properties.alignment, properties.font, and 10 more. */
export const macStyleDialogZh: MacStyleDialogStrings = {
  title: '修改单元格样式',
  nameLabel: '样式名称:',
  previewLabel: '预览:',
  caption: '样式包括(例):',
  propertyColumn: '样式属性',
  formatColumn: '格式',
  properties: {
    number: '数字',
    alignment: '对齐',
    font: '字体',
    border: '边框',
    fill: '填充',
    protection: '保护',
  },
  formatButton: '格式...',
  summary: {
    general: '常规',
    alignment: '对齐: {vertical}, {horizontal}',
    vertical: {
      top: '靠上',
      middle: '居中',
      bottom: '靠下',
    },
    horizontal: {
      general: '常规',
      left: '靠左',
      center: '居中',
      right: '靠右',
    },
    noBorder: '无边框',
    border: '有边框',
    noFill: '无底纹',
    locked: '锁定',
    unlocked: '未锁定',
  },
};

/** Mac app wording from the Excel bundle; not answered by it: previewLabel, caption, formatColumn, properties.alignment, properties.font, properties.protection, and 6 more. */
export const macStyleDialogKo: MacStyleDialogStrings = {
  title: '셀 스타일 수정',
  nameLabel: '스타일 이름:',
  previewLabel: '미리 보기:',
  caption: '스타일에 포함할 항목(예):',
  propertyColumn: '스타일 속성',
  formatColumn: '서식',
  properties: {
    number: '숫자',
    alignment: '맞춤',
    font: '글꼴',
    border: '테두리',
    fill: '채우기',
    protection: '보호',
  },
  formatButton: '서식...',
  summary: {
    general: '일반',
    alignment: '맞춤: {vertical}, {horizontal}',
    vertical: {
      top: '위쪽',
      middle: '가운데',
      bottom: '아래쪽',
    },
    horizontal: {
      general: '일반',
      left: '왼쪽',
      center: '가운데',
      right: '오른쪽',
    },
    noBorder: '테두리 없음',
    border: '테두리 있음',
    noFill: '음영 없음',
    locked: '잠금',
    unlocked: '잠금 해제',
  },
};

/** Mac app wording from the Excel bundle; not answered by it: previewLabel, caption, formatColumn, properties.alignment, properties.font, properties.protection, and 7 more. */
export const macStyleDialogTh: MacStyleDialogStrings = {
  title: 'ปรับเปลี่ยนสไตล์เซลล์',
  nameLabel: 'ชื่อสไตล์:',
  previewLabel: 'แสดงตัวอย่าง:',
  caption: 'สไตล์ประกอบด้วย (ตามตัวอย่าง):',
  propertyColumn: 'คุณสมบัติของสไตล์',
  formatColumn: 'รูปแบบ',
  properties: {
    number: 'ตัวเลข',
    alignment: 'การจัดแนว',
    font: 'ฟอนต์',
    border: 'เส้นขอบ',
    fill: 'เติม',
    protection: 'การป้องกัน',
  },
  formatButton: 'รูปแบบ...',
  summary: {
    general: 'ทั่วไป',
    alignment: 'การจัดแนว: {vertical}, {horizontal}',
    vertical: {
      top: 'บน',
      middle: 'กึ่งกลาง',
      bottom: 'ล่าง',
    },
    horizontal: {
      general: 'ทั่วไป',
      left: 'ซ้าย',
      center: 'กึ่งกลาง',
      right: 'ขวา',
    },
    noBorder: 'ไม่มีเส้นขอบ',
    border: 'มีเส้นขอบ',
    noFill: 'ไม่แรเงา',
    locked: 'ล็อก',
    unlocked: 'ไม่ล็อก',
  },
};
