import type { CellFormatStrings } from '../cell-format/types.js';

type FormatDialogStrings = CellFormatStrings['formatDialog'];

/** The `formatDialog` keys the Mac host can override. Each locale starts from its own shared text and overrides the fields the Excel bundle answers. */
const MAC_FORMAT_DIALOG_KEYS = [
  'categoryLabel',
  'fontFamily',
  'fontStyle',
  'fontSize',
  'fontUnderline',
  'color',
  'fontEffects',
  'descGeneral',
  'descFixed',
  'descCurrency',
  'descAccounting',
  'descPercent',
  'descDate',
  'descTime',
  'descText',
  'descOther',
  'decimals',
  'thousandsSeparator',
  'negativeNumbers',
  'symbol',
  'patternType',
  'languageLocation',
  'calendarType',
  'specialFormatLabels',
  'horizontalAlign',
  'verticalAlign',
  'indent',
  'textControl',
  'direction',
  'rotation',
  'verticalText',
  'borderStyle',
  'borderColor',
  'borderStyleHair',
  'borderStyleDotted',
  'borderStyleDashDotDot',
  'borderStyleDashDot',
  'borderStyleDashed',
  'borderStyleThin',
  'borderStyleMediumDashDotDot',
  'borderStyleSlantDashDot',
  'borderStyleMediumDashDot',
  'borderStyleMediumDashed',
  'borderStyleMedium',
  'borderStyleThick',
  'borderStyleDouble',
  'fill',
  'fillPatternColor',
  'fillPatternStyle',
] as const satisfies readonly (keyof FormatDialogStrings)[];

export const macFormatDialogFrom = (dialog: FormatDialogStrings): Partial<FormatDialogStrings> =>
  Object.fromEntries(MAC_FORMAT_DIALOG_KEYS.map((key) => [key, dialog[key]]));

/** Mac app wording from the Excel bundle; not answered by it: categoryLabel, fontFamily, fontStyle, fontSize, fontUnderline, color, and 31 more. */
export const macFormatDialogDe: Partial<FormatDialogStrings> = {
  decimals: 'Dezimalstellen:',
  symbol: 'Symbol:',
  languageLocation: 'Sprache (Ort):',
  calendarType: 'Kalendertyp:',
  textControl: 'Textsteuerung:',
  borderStyleDashed: 'Gestrichelt',
  borderStyleThin: 'Dünn',
  borderStyleMediumDashed: 'Gestrichelt Mittel',
  borderStyleMedium: 'Mittel',
  borderStyleDouble: 'Doppelt',
  fillPatternColor: 'Musterfarbe:',
  fillPatternStyle: 'Musterformat:',
};

/** Mac app wording from the Excel bundle; not answered by it: categoryLabel, fontFamily, fontStyle, fontSize, fontUnderline, color, and 34 more. */
export const macFormatDialogFr: Partial<FormatDialogStrings> = {
  decimals: 'Nombre de décimales :',
  symbol: 'Symbole :',
  languageLocation: 'Langue (emplacement) :',
  calendarType: 'Type de calendrier :',
  textControl: 'Contrôle du texte :',
  borderStyleMediumDashed: 'Tiret moyen',
  borderStyleDouble: 'Double',
  fillPatternColor: 'Couleur de motif :',
  fillPatternStyle: 'Style de motif :',
};

/** Mac app wording from the Excel bundle; not answered by it: categoryLabel, fontFamily, fontStyle, fontSize, fontUnderline, color, and 34 more. */
export const macFormatDialogKo: Partial<FormatDialogStrings> = {
  decimals: '소수 자릿수:',
  thousandsSeparator: '1000 단위 구분 기호(,) 사용',
  symbol: '기호:',
  languageLocation: '언어(위치):',
  textControl: '텍스트 조정:',
  borderStyleDashed: '파선',
  borderStyleMediumDashed: '중간 굵기 파선',
  fillPatternColor: '패턴 색:',
  fillPatternStyle: '무늬 스타일:',
};

/** Mac app wording from the Excel bundle; not answered by it: categoryLabel, fontFamily, fontStyle, fontSize, fontUnderline, color, and 30 more. */
export const macFormatDialogTh: Partial<FormatDialogStrings> = {
  thousandsSeparator: 'ใช้ตัวคั่นหลักพัน (,)',
  symbol: 'สัญลักษณ์:',
  languageLocation: 'ภาษา (ตำแหน่งที่ตั้ง):',
  calendarType: 'ชนิดปฏิทิน:',
  textControl: 'การควบคุมข้อความ:',
  borderStyleDotted: 'จุด',
  borderStyleDashed: 'เส้นประ',
  borderStyleThin: 'บาง',
  borderStyleMediumDashed: 'เส้นประยาวปานกลาง',
  borderStyleMedium: 'ปานกลาง',
  borderStyleThick: 'หนา',
  fillPatternColor: 'สีลวดลาย:',
  fillPatternStyle: 'สไตล์ลวดลาย:',
};

/** Mac app wording from the Excel bundle; not answered by it: categoryLabel, fontFamily, fontStyle, fontSize, fontUnderline, color, and 35 more. */
export const macFormatDialogZh: Partial<FormatDialogStrings> = {
  decimals: '小数位数:',
  thousandsSeparator: '使用千位分隔符(,)',
  languageLocation: '语言(位置):',
  calendarType: '日历类型:',
  textControl: '文本控制:',
  borderStyleMediumDashed: '中划线',
  fillPatternColor: '图案颜色:',
  fillPatternStyle: '图案样式:',
};
