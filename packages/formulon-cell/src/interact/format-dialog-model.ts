import type {
  CellAlign,
  CellBorders,
  CellVAlign,
  CellValidation,
  NumFmt,
  ValidationErrorStyle,
  ValidationOp,
} from '../store/store.js';

/** Discriminator for the dialog's "kind" dropdown. `none` means clear the
 *  validation; the rest mirror `CellValidation['kind']`. */
export type ValidationKind = 'none' | CellValidation['kind'];

export type TabId = 'number' | 'align' | 'font' | 'border' | 'fill' | 'protection' | 'more';
export type NumberCategory =
  | 'general'
  | 'fixed'
  | 'currency'
  | 'percent'
  | 'scientific'
  | 'accounting'
  | 'date'
  | 'time'
  | 'datetime'
  | 'text'
  | 'custom';
export type BorderStyleKey = 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double';
export type SideKey = 'top' | 'right' | 'bottom' | 'left' | 'diagonalDown' | 'diagonalUp';

export interface DraftState {
  numFmt: NumFmt | undefined;
  numberCategory: NumberCategory;
  decimals: number;
  currencySymbol: string;
  /** Pattern for date/time/datetime/custom categories. */
  pattern: string;
  align: CellAlign | undefined;
  vAlign: CellVAlign | undefined;
  wrap: boolean;
  indent: number;
  rotation: number;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  fontFamily: string;
  fontSize: number | undefined;
  color: string | undefined;
  fill: string | undefined;
  borders: CellBorders;
  /** "Active" line style — applied when a side checkbox is turned on. */
  borderStyle: BorderStyleKey;
  /** "Active" line color in #rrggbb form, or undefined for theme default. */
  borderColor: string | undefined;
  hyperlink: string;
  comment: string;
  validationList: string;
  /** When kind === 'list', selects between inline string array and a range
   *  reference (spreadsheet-style `Sheet1!$A$1:$A$10`). */
  validationListSourceKind: 'literal' | 'range';
  validationListRange: string;
  validationKind: ValidationKind;
  validationOp: ValidationOp;
  validationA: number;
  validationB: number;
  validationFormula: string;
  validationAllowBlank: boolean;
  validationErrorStyle: ValidationErrorStyle;
  /** Sheet-protection lock flag. desktop default is `true` (locked); the
   *  Protection tab exposes a single checkbox. */
  locked: boolean;
}

export const COMMON_FONTS = [
  'system-ui',
  'Helvetica',
  'Arial',
  'Georgia',
  'Times New Roman',
  'Courier New',
  'monospace',
];
export const CURRENCY_SYMBOLS = ['$', '¥', '€', '£'];
export const THEME_SWATCHES = [
  '#000000',
  '#ffffff',
  '#c00000',
  '#ff0000',
  '#ffc000',
  '#ffff00',
  '#92d050',
  '#00b050',
  '#00b0f0',
  '#0070c0',
  '#002060',
  '#7030a0',
] as const;

export const normalizeFormatLocale = (locale: string): string => {
  if (locale === 'ja') return 'ja-JP';
  if (locale === 'en') return 'en-US';
  return locale || 'en-US';
};

export const defaultCurrencySymbolFor = (locale: string): string =>
  normalizeFormatLocale(locale).startsWith('ja') ? '¥' : '$';

export const patternPresetsFor = (
  locale: string,
): Record<'date' | 'time' | 'datetime' | 'custom', string[]> => {
  if (normalizeFormatLocale(locale).startsWith('ja')) {
    return {
      date: [
        'yyyy"年"m"月"d"日"',
        'yyyy/m/d',
        'yyyy-mm-dd',
        'm"月"d"日"',
        'yyyy"年"m"月"d"日" ddd',
      ],
      time: ['HH:MM', 'HH:MM:SS', 'h:MM AM/PM', 'h:MM:SS AM/PM'],
      datetime: ['yyyy"年"m"月"d"日" HH:MM', 'yyyy/m/d HH:MM', 'yyyy-mm-dd HH:MM'],
      custom: ['#,##0', '#,##0.00', '0%', '0.00%', '¥#,##0;[Red]-¥#,##0', '0.00E+00'],
    };
  }
  return {
    date: ['m/d/yyyy', 'mmmm d, yyyy', 'd-mmm-yy', 'yyyy-mm-dd', 'dddd, mmmm d, yyyy'],
    time: ['h:MM AM/PM', 'h:MM:SS AM/PM', 'HH:MM', 'HH:MM:SS'],
    datetime: ['m/d/yyyy h:MM AM/PM', 'mmmm d, yyyy h:MM AM/PM', 'yyyy-mm-dd HH:MM'],
    custom: ['0.00', '#,##0', '#,##0.00', '0%', '0.00%', '$#,##0;[Red]-$#,##0', '0.00E+00'],
  };
};

export function isHexColor(s: string): boolean {
  return /^#[0-9a-fA-F]{6}$/.test(s);
}
