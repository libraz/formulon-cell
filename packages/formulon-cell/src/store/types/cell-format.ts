import type { PhoneticRun } from '../../engine/types.js';

export type NumFmt =
  | { kind: 'general' }
  | { kind: 'fixed'; decimals: number; thousands?: boolean; negativeStyle?: NegativeStyle }
  | {
      kind: 'currency';
      decimals: number;
      symbol?: string;
      negativeStyle?: NegativeStyle;
    }
  | { kind: 'percent'; decimals: number }
  | { kind: 'scientific'; decimals: number }
  | { kind: 'accounting'; decimals: number; symbol?: string }
  | { kind: 'date'; pattern: string }
  | { kind: 'time'; pattern: string }
  | { kind: 'datetime'; pattern: string }
  | { kind: 'special'; pattern: string }
  | { kind: 'text' }
  | { kind: 'custom'; pattern: string };

/** How negative numbers display. */
export type NegativeStyle = 'minus' | 'parens' | 'red' | 'red-parens';

export type CellAlign =
  | 'left'
  | 'center'
  | 'right'
  | 'fill'
  | 'justify'
  | 'centerContinuous'
  | 'distributed';
export type CellVAlign = 'top' | 'middle' | 'bottom' | 'justify' | 'distributed';
export type TextDirection = 'context' | 'ltr' | 'rtl';
/** OOXML font underline variants. `true` remains a legacy alias for single. */
export type UnderlineStyle = 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';
export type FillPattern =
  | 'gray125'
  | 'gray0625'
  | 'gray25'
  | 'gray50'
  | 'gray75'
  | 'darkHorizontal'
  | 'darkVertical'
  | 'darkDown'
  | 'darkUp'
  | 'darkGrid'
  | 'darkTrellis'
  | 'lightHorizontal'
  | 'lightVertical'
  | 'lightDown'
  | 'lightUp'
  | 'lightGrid'
  | 'lightTrellis'
  /** Legacy aliases accepted on input; the UI writes the OOXML names above. */
  | 'horizontal'
  | 'vertical'
  | 'diagonalDown'
  | 'diagonalUp';

/** Per-side border style. The renderer treats `false`/missing as "no border"
 *  and `true` as the legacy single-line border (back-compat). Object form
 *  carries a spreadsheet-style style + optional color. The full OOXML
 *  repertoire is supported; common spreadsheet border ordinals map to these
 *  names verbatim. */
export type CellBorderStyle =
  | 'thin'
  | 'medium'
  | 'thick'
  | 'dashed'
  | 'dotted'
  | 'double'
  | 'hair'
  | 'mediumDashed'
  | 'dashDot'
  | 'mediumDashDot'
  | 'dashDotDot'
  | 'mediumDashDotDot'
  | 'slantDashDot';

export type CellBorderSide =
  | boolean
  | {
      style: CellBorderStyle;
      color?: string;
    };

export interface CellBorders {
  top?: CellBorderSide;
  right?: CellBorderSide;
  bottom?: CellBorderSide;
  left?: CellBorderSide;
  /** Diagonal border directions: `\` runs top-left → bottom-right and `/`
   *  runs bottom-left → top-right. */
  diagonalDown?: CellBorderSide;
  diagonalUp?: CellBorderSide;
}

export interface CellFormat {
  /** Named cell style id last applied through the Cell Styles gallery. Direct
   *  formatting may still override individual fields, mirroring spreadsheets'
   *  style + local-format layering. */
  cellStyle?: string;
  numFmt?: NumFmt;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean | UnderlineStyle;
  strike?: boolean;
  /** Baseline shift for super/subscript text. */
  fontVertAlign?: 'superscript' | 'subscript';
  /** OOXML phonetic guide (ruby/furigana) over the cell text, one entry per
   *  `<rPh>` block. A guide that reads the whole cell is a single run; a
   *  partially annotated cell keeps one run per annotated span. */
  phonetic?: readonly PhoneticRun[];
  align?: CellAlign;
  /** Vertical alignment. Default is 'bottom'. */
  vAlign?: CellVAlign;
  /** Wrap text within the cell — paint multi-line with hard wrapping. */
  wrap?: boolean;
  /** With distributed horizontal alignment, justify the final text line too. */
  justifyLastLine?: boolean;
  /** Shrink single-line text to fit the available cell width. */
  shrinkToFit?: boolean;
  /** Indent level (left-align padding in increments of ~8px). 0..15. */
  indent?: number;
  /** Text rotation in degrees, -90..90. 0 = horizontal. */
  rotation?: number;
  /** Text reading direction. Context maps to Excel's default/context mode. */
  textDirection?: TextDirection;
  borders?: CellBorders;
  /** Foreground (font) color as a CSS color string. */
  color?: string;
  /** Background fill color as a CSS color string. */
  fill?: string;
  /** Spreadsheet fill pattern drawn over `fill` using `fillPatternColor`. */
  fillPattern?: FillPattern;
  /** Foreground color for `fillPattern`. */
  fillPatternColor?: string;
  /** Override theme font family for this cell. */
  fontFamily?: string;
  /** Font size in CSS pixels. */
  fontSize?: number;
  /** Hyperlink URL. When set, the cell paints text underlined+blue and
   *  Ctrl/Cmd+click opens the link. */
  hyperlink?: string;
  /** Optional hyperlink display text from the engine. Preserved for xlsx round-trip. */
  hyperlinkDisplay?: string;
  /** Optional hyperlink tooltip from the engine. Preserved for xlsx round-trip. */
  hyperlinkTooltip?: string;
  /** Free-form note attached to the cell. Surfaced as a small triangle
   *  marker + hover tooltip; not exported to .xlsx for now. */
  comment?: string;
  /** Optional comment author from the engine. Preserved for xlsx round-trip. */
  commentAuthor?: string;
  /** Data validation. When kind === 'list', the cell paints a small ▼ on its
   *  right edge; clicking it opens a dropdown of `source` values. Other
   *  kinds (whole/decimal/date/time/textLength/custom) constrain typed input
   *  through `validateAgainst()` — the chevron is list-only. */
  validation?: CellValidation;
  /** Sheet-protection lock flag. Default is `true` (locked) — `undefined`
   *  is treated as locked. Set to `false` to opt the cell out of the
   *  per-sheet protection gate via `setCellLocked(range, false)`. The flag
   *  only takes effect when the containing sheet is also marked protected
   *  via `setSheetProtected`. */
  locked?: boolean;
  /** Formula-hidden flag. When true and the containing sheet is protected,
   *  formula text is suppressed from the formula bar, matching the desktop
   *  Format Cells > Protection > Hidden behavior. */
  formulaHidden?: boolean;
}

export interface CustomCellStyle {
  id: string;
  label: string;
  format: Partial<CellFormat>;
}

/** Comparison ordinals match OOXML data-validation `op`:
 *  0 between, 1 notBetween, 2 equal, 3 notEqual,
 *  4 lessThan, 5 lessThanOrEqual, 6 greaterThan, 7 greaterThanOrEqual. */
export type ValidationOp = 'between' | 'notBetween' | '=' | '<>' | '<' | '<=' | '>' | '>=';

/** OOXML errorStyle: 0 stop, 1 warning, 2 information. `stop` rejects the
 *  input outright; `warning` and `information` let the user keep the entry. */
export type ValidationErrorStyle = 'stop' | 'warning' | 'information';

/** Optional metadata that every validation kind carries. Mirrors the
 *  upstream `DataValidationEntry` shape minus `type` / `op` / formulas which
 *  the discriminated cases own. */
export interface ValidationMeta {
  /** Allow empty input regardless of constraint. Default true. */
  allowBlank?: boolean;
  errorStyle?: ValidationErrorStyle;
  errorTitle?: string;
  errorMessage?: string;
  promptTitle?: string;
  promptMessage?: string;
  /** Suppress the prompt tooltip even when the metadata is set. Default true. */
  showInputMessage?: boolean;
  /** Suppress the error dialog on invalid entry. Default true. */
  showErrorMessage?: boolean;
  /** Show the in-cell dropdown affordance for list validations. Default true. */
  showDropdown?: boolean;
}

/** A list-source can be either an inline literal array of strings or a range
 *  reference (`Sheet1!$A$1:$A$10` or `$A$1:$A$10`). Range refs are resolved
 *  lazily by the dropdown / validator so the source-of-truth stays a single
 *  string in the OOXML formula1 slot. */
export type ValidationListSource = string[] | { ref: string };

/** Discriminated union — `kind` mirrors the OOXML `type` ordinal:
 *  list (3), whole (1), decimal (2), date (4), time (5), textLength (6),
 *  custom (7). */
export type CellValidation =
  | ({ kind: 'list'; source: ValidationListSource } & ValidationMeta)
  | ({ kind: 'whole'; op: ValidationOp; a: number; b?: number } & ValidationMeta)
  | ({ kind: 'decimal'; op: ValidationOp; a: number; b?: number } & ValidationMeta)
  | ({ kind: 'date'; op: ValidationOp; a: number; b?: number } & ValidationMeta)
  | ({ kind: 'time'; op: ValidationOp; a: number; b?: number } & ValidationMeta)
  | ({ kind: 'textLength'; op: ValidationOp; a: number; b?: number } & ValidationMeta)
  | ({ kind: 'custom'; formula: string } & ValidationMeta);
