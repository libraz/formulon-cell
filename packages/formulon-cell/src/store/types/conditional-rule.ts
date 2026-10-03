import type { Range } from '../../engine/types.js';
import type { CellFormat } from './cell-format.js';

/** Icon-set artwork name. 3-slot families classify by [0.33, 0.67];
 *  5-slot families classify by [0.20, 0.40, 0.60, 0.80]. */
export type ConditionalIconSet =
  | 'arrows3'
  | 'arrows5'
  | 'triangles3'
  | 'traffic3'
  | 'trafficRim3'
  | 'symbols3'
  | 'flags3'
  | 'stars3'
  | 'quarters5'
  | 'ratings5'
  | 'bars5'
  | 'boxes5';

export type ConditionalScalePoint =
  | { kind: 'min' | 'max'; gte?: boolean }
  | { kind: 'number' | 'percent' | 'percentile'; value: number; gte?: boolean };

/** Conditional formatting rule. Evaluated by the renderer against cell
 *  values; the predicate kinds (cell-value, top-bottom, formula, blanks,
 *  duplicates, etc.) skip cells that don't satisfy their type-specific
 *  filter. */
type ConditionalRulePriorityOptions = {
  /** When true, lower-priority conditional rules are skipped for cells matched
   *  by this rule. Mirrors Excel's "Stop If True" flag. */
  stopIfTrue?: boolean;
  /** Engine rule id when this rule was hydrated from an imported workbook.
   *  Session writeback skips these rules so imported CF is not duplicated. */
  engineId?: string;
};

export type ConditionalRule = ConditionalRulePriorityOptions &
  (
    | {
        kind: 'cell-value';
        range: Range;
        op: '>' | '<' | '>=' | '<=' | '=' | '<>' | 'between' | 'not-between';
        a: number | string;
        b?: number | string;
        /** Format applied when the predicate matches. Same shape as CellFormat
         *  but only `fill`, `color`, `bold`, `italic`, `underline`, `strike`
         *  are honored by the renderer. */
        apply: Partial<CellFormat>;
      }
    | {
        kind: 'color-scale';
        range: Range;
        /** Two- or three-stop gradient. Stops are CSS color strings; the values
         *  are interpolated linearly across [min, max] (or [min, mid, max]) of
         *  the range. */
        stops: [string, string] | [string, string, string];
        /** Excel-style threshold metadata for each color stop. Omitted rules use
         *  min/max for two-color scales and min/50th percentile/max for
         *  three-color scales. */
        thresholds?:
          | [ConditionalScalePoint, ConditionalScalePoint]
          | [ConditionalScalePoint, ConditionalScalePoint, ConditionalScalePoint];
      }
    | {
        kind: 'data-bar';
        range: Range;
        color: string;
        /** Position of the data-bar axis. Omitted rules use Excel automatic
         *  signed-axis behaviour. */
        axisPosition?: 'automatic' | 'middle' | 'none';
        /** Fill used for negative values. Omitted rules inherit `color`. */
        negativeColor?: string;
        /** Border used for positive values. Omitted rules have no border. */
        borderColor?: string;
        /** Border used for negative values. Omitted rules have no negative border. */
        negativeBorderColor?: string;
        /** Axis line colour. Omitted rules use black. */
        axisColor?: string;
        /** Lower endpoint used to scale the bar. Omitted rules use the
         *  existing automatic minimum derived from the range. */
        min?: ConditionalScalePoint;
        /** Upper endpoint used to scale the bar. Omitted rules use the
         *  existing automatic maximum derived from the range. */
        max?: ConditionalScalePoint;
        /** Edge from which the bar grows. Context follows the sheet's
         *  reading direction. Omitted rules use the engine default. */
        direction?: 'context' | 'left-to-right' | 'right-to-left';
        /** True for Excel's gradient-fill data bars, false for solid-fill bars.
         *  Omitted legacy rules render as solid bars. */
        gradient?: boolean;
        /** When true, paint the bar across the whole cell with the text on top
         *  (like the spreadsheet's "Show Bar Only" being false). */
        showValue?: boolean;
      }
    | {
        kind: 'icon-set';
        range: Range;
        /** Icon family. 3-slot or 5-slot determined by the suffix. */
        icons: ConditionalIconSet;
        /** When false, render only the icon and suppress the cell value text. */
        showValue?: boolean;
        /** Lower bound of the lowest icon bucket. Numeric values below this
         *  threshold render without an icon. */
        floor?: ConditionalScalePoint;
        /** Boundaries between icon slots, ordered from low to high. Omitted
         *  rules use Excel's default percent thresholds. */
        thresholds?: ConditionalScalePoint[];
        /** Invert slot index so the highest values get the "low" icon. */
        reverseOrder?: boolean;
      }
    | {
        kind: 'top-bottom';
        range: Range;
        /** `top` selects the N largest values, `bottom` the N smallest. */
        mode: 'top' | 'bottom';
        n: number;
        /** When true, `n` is interpreted as a percentage (0..100) of the
         *  range's numeric-cell count. */
        percent?: boolean;
        apply: Partial<CellFormat>;
      }
    | {
        kind: 'average';
        range: Range;
        /** Spreadsheet average rules over numeric cells in the range. */
        mode:
          | 'above'
          | 'below'
          | 'equal-or-above'
          | 'equal-or-below'
          | 'above-std-dev'
          | 'below-std-dev';
        /** Standard-deviation tier for std-dev average rules. Defaults to 1. */
        stdDev?: 1 | 2 | 3;
        apply: Partial<CellFormat>;
      }
    | {
        kind: 'text-contains';
        range: Range;
        text: string;
        mode?: 'contains' | 'not-contains' | 'begins-with' | 'ends-with';
        caseSensitive?: boolean;
        apply: Partial<CellFormat>;
      }
    | {
        kind: 'date-occurring';
        range: Range;
        period:
          | 'yesterday'
          | 'today'
          | 'tomorrow'
          | 'last7'
          | 'last-week'
          | 'this-week'
          | 'next-week'
          | 'last-month'
          | 'this-month'
          | 'next-month';
        apply: Partial<CellFormat>;
      }
    | {
        kind: 'formula';
        range: Range;
        /** Lightweight predicate. Supports comparator-prefix forms
         *  (`>10`, `<>"foo"`, `<= 0`, `=42`) and an `=`-prefixed cell formula
         *  evaluated through `wb.evaluateText` when the engine exposes one;
         *  otherwise the rule is a no-op. */
        formula: string;
        apply: Partial<CellFormat>;
      }
    | {
        kind: 'duplicates' | 'unique';
        range: Range;
        apply: Partial<CellFormat>;
      }
    | {
        kind: 'blanks' | 'non-blanks' | 'errors' | 'no-errors';
        range: Range;
        apply: Partial<CellFormat>;
      }
  );
