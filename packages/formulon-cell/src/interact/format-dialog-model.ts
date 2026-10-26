import {
  currencySymbolForLocale,
  datePatternsForLocale,
  normalizeFormatLocale,
} from '../format/locale.js';
import type {
  CellAlign,
  CellBorderStyle,
  CellBorders,
  CellFormat,
  CellGradientFill,
  CellVAlign,
  CellValidation,
  FillPattern,
  NegativeStyle,
  NumFmt,
  TextDirection,
  ValidationErrorStyle,
  ValidationOp,
} from '../store/store.js';

export { normalizeFormatLocale };

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
  | 'text'
  | 'fraction'
  | 'special'
  | 'custom';
export type BorderStyleKey = CellBorderStyle;
export type SideKey = 'top' | 'right' | 'bottom' | 'left' | 'diagonalDown' | 'diagonalUp';

export interface DraftState {
  numFmt: NumFmt | undefined;
  numberCategory: NumberCategory;
  decimals: number;
  thousands: boolean;
  negativeStyle: NegativeStyle;
  currencySymbol: string;
  /** Pattern for date/time/fraction/custom categories. */
  pattern: string;
  align: CellAlign | undefined;
  vAlign: CellVAlign | undefined;
  wrap: boolean;
  justifyLastLine: boolean;
  shrinkToFit: boolean;
  indent: number;
  rotation: number;
  textDirection: TextDirection;
  bold: boolean;
  italic: boolean;
  underline: CellFormat['underline'];
  strike: boolean;
  fontVertAlign: CellFormat['fontVertAlign'];
  fontOutline: boolean;
  fontShadow: boolean;
  fontFamily: string;
  fontSize: number | undefined;
  color: string | undefined;
  fill: string | undefined;
  fillPattern: FillPattern | undefined;
  fillPatternColor: string | undefined;
  fillGradient: CellGradientFill | undefined;
  borders: CellBorders;
  /** "Active" line style — applied when a side checkbox is turned on. */
  borderStyle: BorderStyleKey;
  /** "Active" line color in #rrggbb form, or undefined for theme default. */
  borderColor: string | undefined;
  hyperlink: string;
  originalHyperlink: string;
  hyperlinkDisplay: string | undefined;
  hyperlinkTooltip: string | undefined;
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
  /** Formula limits of a loaded rule; kept until the numeric field is edited. */
  validationAFormula?: string;
  validationBFormula?: string;
  validationFormula: string;
  validationAllowBlank: boolean;
  validationErrorStyle: ValidationErrorStyle;
  validationShowInputMessage: boolean;
  validationPromptTitle: string;
  validationPromptMessage: string;
  validationShowErrorMessage: boolean;
  validationErrorTitle: string;
  validationErrorMessage: string;
  /** In-cell list dropdown arrow visibility. Excel/OOXML stores the hidden
   *  state inverted as `showDropDown=true`; the draft keeps the UI meaning. */
  validationShowDropdown: boolean;
  /** Sheet-protection lock flag. desktop default is `true` (locked); the
   *  Protection tab exposes a single checkbox. */
  locked: boolean;
  /** Formula-hidden flag exposed by the Protection tab. */
  formulaHidden: boolean;
}

export const CURRENCY_SYMBOLS = ['$', '¥', '€', '£', '₩', '฿'];
export const FRACTION_PATTERNS = [
  '# ?/?',
  '# ??/??',
  '# ???/???',
  '# ?/2',
  '# ?/4',
  '# ?/8',
  '# ?/16',
  '# ?/10',
  '# ?/100',
] as const;

export const defaultCurrencySymbolFor = currencySymbolForLocale;

export const patternPresetsFor = (
  locale: string,
): Record<'date' | 'time' | 'fraction' | 'special' | 'custom', string[]> => {
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
      fraction: [...FRACTION_PATTERNS],
      special: [
        '000',
        '[<=9999999]000-0000;000-0000',
        '+0;△0',
        '[DBNum1]0',
        '[DBNum2]0',
        '[DBNum3]0',
        '[DBNum3]#,##0',
      ],
      custom: ['#,##0', '#,##0.00', '0%', '0.00%', '¥#,##0;[Red]-¥#,##0', '0.00E+00'],
    };
  }
  const language = normalizeFormatLocale(locale).split('-')[0];
  if (language !== 'en') {
    const patterns = datePatternsForLocale(locale);
    const symbol = currencySymbolForLocale(locale);
    return {
      date: [patterns.short, patterns.long, 'yyyy-mm-dd'],
      time: [patterns.time, 'HH:MM:SS', 'h:MM AM/PM'],
      fraction: [...FRACTION_PATTERNS],
      special: ['00000'],
      custom: [
        '#,##0',
        '#,##0.00',
        '0%',
        '0.00%',
        `${symbol}#,##0;[Red]-${symbol}#,##0`,
        '0.00E+00',
      ],
    };
  }
  return {
    date: ['m/d/yyyy', 'mmmm d, yyyy', 'd-mmm-yy', 'yyyy-mm-dd', 'dddd, mmmm d, yyyy'],
    time: ['h:MM AM/PM', 'h:MM:SS AM/PM', 'HH:MM', 'HH:MM:SS'],
    fraction: [...FRACTION_PATTERNS],
    special: ['00000', '00000-0000', '[<=9999999]000-0000;(000) 000-0000', '000-00-0000'],
    custom: ['0.00', '#,##0', '#,##0.00', '0%', '0.00%', '$#,##0;[Red]-$#,##0', '0.00E+00'],
  };
};

/** First locale preset for a pattern-driven category, or a fixed fallback;
 *  '' for categories whose number format carries no pattern. */
export const defaultPatternForLocale = (category: NumberCategory, locale: string): string => {
  const presets =
    category === 'date' ||
    category === 'time' ||
    category === 'fraction' ||
    category === 'special' ||
    category === 'custom'
      ? patternPresetsFor(locale)[category]
      : [];
  if (presets[0]) return presets[0];
  switch (category) {
    case 'date':
      return 'yyyy-mm-dd';
    case 'time':
      return 'HH:MM:SS';
    case 'fraction':
      return '# ?/?';
    case 'special':
      return '000';
    case 'custom':
      return '0.00';
    default:
      return '';
  }
};

export function isHexColor(s: string): boolean {
  return /^#[0-9a-fA-F]{6}$/.test(s);
}

/** Fill Effects shading styles, in the desktop dialog's order. */
export const GRADIENT_SHADING_STYLES = [
  'horizontal',
  'vertical',
  'diagonalUp',
  'diagonalDown',
  'fromCorner',
  'fromCenter',
] as const;
export type GradientShadingStyle = (typeof GRADIENT_SHADING_STYLES)[number];

const LINEAR_DEGREES: Record<Exclude<GradientShadingStyle, 'fromCorner' | 'fromCenter'>, number> = {
  horizontal: 90,
  vertical: 0,
  diagonalUp: 45,
  diagonalDown: 135,
};

/** Path-gradient focus per From Corner variant: top-left, top-right,
 *  bottom-left, bottom-right. */
const CORNERS: readonly (readonly [number, number])[] = [
  [0, 0],
  [1, 0],
  [0, 1],
  [1, 1],
];

export const gradientVariantCount = (style: GradientShadingStyle): number =>
  style === 'fromCenter' ? 2 : 4;

/** The two-colour gradient a shading style and variant produce. Linear
 *  variants run colour 1 to 2, 2 to 1, then the two mirrored three-stop forms. */
export function gradientForShading(
  style: GradientShadingStyle,
  variant: number,
  color1: string,
  color2: string,
): CellGradientFill {
  const ramp = (a: string, b: string): CellGradientFill['stops'] => [
    { position: 0, color: a },
    { position: 1, color: b },
  ];
  if (style === 'fromCenter') {
    return {
      type: 'path',
      degree: 0,
      left: 0.5,
      right: 0.5,
      top: 0.5,
      bottom: 0.5,
      stops: variant === 1 ? ramp(color2, color1) : ramp(color1, color2),
    };
  }
  if (style === 'fromCorner') {
    const [x, y] = CORNERS[variant] ?? [0, 0];
    return {
      type: 'path',
      degree: 0,
      left: x,
      right: x,
      top: y,
      bottom: y,
      stops: ramp(color1, color2),
    };
  }
  const mirrored = (a: string, b: string): CellGradientFill['stops'] => [
    { position: 0, color: a },
    { position: 0.5, color: b },
    { position: 1, color: a },
  ];
  const stops =
    variant === 1
      ? ramp(color2, color1)
      : variant === 2
        ? mirrored(color1, color2)
        : variant === 3
          ? mirrored(color2, color1)
          : ramp(color1, color2);
  return {
    type: 'linear',
    degree: LINEAR_DEGREES[style],
    left: 0,
    right: 0,
    top: 0,
    bottom: 0,
    stops,
  };
}

/** The shading style, variant and colours that reproduce `gradient`, or null
 *  for a loaded gradient the two-colour dialog cannot express. */
export function shadingForGradient(gradient: CellGradientFill): {
  style: GradientShadingStyle;
  variant: number;
  color1: string;
  color2: string;
} | null {
  const first = gradient.stops[0]?.color;
  const second = gradient.stops[1]?.color;
  if (!first || !second) return null;
  for (const style of GRADIENT_SHADING_STYLES) {
    for (let variant = 0; variant < gradientVariantCount(style); variant += 1) {
      for (const [color1, color2] of [
        [first, second],
        [second, first],
      ] as const) {
        const candidate = gradientForShading(style, variant, color1, color2);
        if (JSON.stringify(candidate) === JSON.stringify(gradient)) {
          return { style, variant, color1, color2 };
        }
      }
    }
  }
  return null;
}
