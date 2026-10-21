import { PivotAggregation } from './types.js';

/** `PivotCell.kind` ordinals reported by `pivotLayout()`. */
export const PIVOT_CELL_KIND = {
  Header: 0,
  RowLabel: 1,
  ColLabel: 2,
  Data: 3,
  RowSubtotal: 4,
  ColSubtotal: 5,
  GrandTotal: 6,
  Blank: 7,
} as const;

/** UI-language wording for the captions the engine emits in English, plus the
 *  alert shown when an edit would touch a PivotTable cell. */
export interface PivotCaptions {
  rowLabels: string;
  columnLabels: string;
  grandTotal: string;
  blank: string;
  /** Subtotal caption; `{item}` stands for the item name. */
  subtotal: string;
  cannotChange: string;
}

/** Engine wording; also the fallback for a handle that was never given captions. */
export const ENGLISH_PIVOT_CAPTIONS: PivotCaptions = {
  rowLabels: 'Row Labels',
  columnLabels: 'Column Labels',
  grandTotal: 'Grand Total',
  blank: '(blank)',
  subtotal: '{item} Total',
  cannotChange: 'Cannot change this part of a PivotTable report.',
};

const SUBTOTAL_SUFFIX = ' Total';

const localizeItem = (item: string, captions: PivotCaptions): string =>
  item === ENGLISH_PIVOT_CAPTIONS.blank ? captions.blank : item;

/** Translate an engine caption by cell kind and exact English text. Item names
 *  and data are returned untouched. */
export function localizePivotCaption(kind: number, text: string, captions: PivotCaptions): string {
  switch (kind) {
    case PIVOT_CELL_KIND.Header:
      if (text === ENGLISH_PIVOT_CAPTIONS.rowLabels) return captions.rowLabels;
      if (text === ENGLISH_PIVOT_CAPTIONS.columnLabels) return captions.columnLabels;
      if (text === ENGLISH_PIVOT_CAPTIONS.grandTotal) return captions.grandTotal;
      return text;
    case PIVOT_CELL_KIND.GrandTotal:
      return text === ENGLISH_PIVOT_CAPTIONS.grandTotal ? captions.grandTotal : text;
    case PIVOT_CELL_KIND.RowLabel:
    case PIVOT_CELL_KIND.ColLabel:
      return localizeItem(text, captions);
    case PIVOT_CELL_KIND.RowSubtotal:
    case PIVOT_CELL_KIND.ColSubtotal:
      if (text.length > SUBTOTAL_SUFFIX.length && text.endsWith(SUBTOTAL_SUFFIX)) {
        const item = text.slice(0, -SUBTOTAL_SUFFIX.length);
        return captions.subtotal.replace('{item}', () => localizeItem(item, captions));
      }
      return localizeItem(text, captions);
    default:
      return text;
  }
}

/** Per-aggregation prefix of a value field's default name ("Sum of"). */
export interface PivotDataFieldPrefixes {
  sum: string;
  count: string;
  average: string;
  max: string;
  min: string;
  product: string;
  countNumbers: string;
  stdDev: string;
  stdDevP: string;
  variance: string;
  varianceP: string;
}

/** How a new value field is named in the UI language. */
export interface PivotDataFieldNaming extends PivotDataFieldPrefixes {
  /** Joins prefix and field name; `{prefix}` and `{field}` are substituted. */
  template: string;
}

export const ENGLISH_PIVOT_DATA_FIELD_NAMING: PivotDataFieldNaming = {
  template: '{prefix} {field}',
  sum: 'Sum of',
  count: 'Count of',
  average: 'Average of',
  max: 'Max of',
  min: 'Min of',
  product: 'Product of',
  countNumbers: 'Count of',
  stdDev: 'StdDev of',
  stdDevP: 'StdDevp of',
  variance: 'Var of',
  varianceP: 'Varp of',
};

const PREFIX_KEY: Record<PivotAggregation, keyof PivotDataFieldPrefixes> = {
  [PivotAggregation.Sum]: 'sum',
  [PivotAggregation.Count]: 'count',
  [PivotAggregation.Average]: 'average',
  [PivotAggregation.Max]: 'max',
  [PivotAggregation.Min]: 'min',
  [PivotAggregation.Product]: 'product',
  [PivotAggregation.CountNumbers]: 'countNumbers',
  [PivotAggregation.StdDev]: 'stdDev',
  [PivotAggregation.StdDevP]: 'stdDevP',
  [PivotAggregation.Var]: 'variance',
  [PivotAggregation.VarP]: 'varianceP',
};

/** Default value-field name, e.g. "Sum of Sales" or "合計 / Sales". */
export const pivotDataFieldName = (
  aggregation: PivotAggregation,
  field: string,
  naming: PivotDataFieldNaming = ENGLISH_PIVOT_DATA_FIELD_NAMING,
): string =>
  naming.template
    .replace('{prefix}', () => naming[PREFIX_KEY[aggregation]])
    .replace('{field}', () => field);
