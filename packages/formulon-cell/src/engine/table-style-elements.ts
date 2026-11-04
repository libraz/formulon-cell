import type { TableStyleElementType } from '@libraz/formulon';
import type { CellBorders, CellFormat } from '../store/types.js';

/**
 * The style-element model shared by table and PivotTable styles: a style is a
 * set of elements (`<tableStyleElement>` types), each painting a format over
 * the cells of the regions it covers. A cell takes every element covering it,
 * composed in the spreadsheet's precedence order; an element's outer border
 * sides land only on the edges of the region it covers.
 */

/** What one style element paints over the cells it covers. */
export interface StyleElementFormat {
  fill?: string;
  color?: string;
  bold?: boolean;
  borders?: Pick<CellBorders, 'top' | 'bottom' | 'left' | 'right'>;
}

export type StyleElementFormats = Partial<Record<TableStyleElementType, StyleElementFormat>>;

/** One rectangle an element covers, in absolute sheet rows / columns. */
export interface StyleElementRegion {
  type: TableStyleElementType;
  r0: number;
  r1: number;
  c0: number;
  c1: number;
}

/** Table elements, lowest precedence first. */
export const TABLE_ELEMENT_ORDER: readonly TableStyleElementType[] = [
  'wholeTable',
  'firstColumnStripe',
  'secondColumnStripe',
  'firstRowStripe',
  'secondRowStripe',
  'lastColumn',
  'firstColumn',
  'headerRow',
  'totalRow',
  'firstHeaderCell',
  'lastHeaderCell',
  'firstTotalCell',
  'lastTotalCell',
];

/** PivotTable elements, lowest precedence first. */
export const PIVOT_ELEMENT_ORDER: readonly TableStyleElementType[] = [
  'wholeTable',
  'pageFieldLabels',
  'pageFieldValues',
  'firstColumnStripe',
  'secondColumnStripe',
  'firstRowStripe',
  'secondRowStripe',
  'firstColumn',
  'headerRow',
  'firstHeaderCell',
  'firstSubtotalColumn',
  'secondSubtotalColumn',
  'thirdSubtotalColumn',
  'firstSubtotalRow',
  'secondSubtotalRow',
  'thirdSubtotalRow',
  'firstColumnSubheading',
  'secondColumnSubheading',
  'thirdColumnSubheading',
  'firstRowSubheading',
  'secondRowSubheading',
  'thirdRowSubheading',
  'blankRow',
  'lastColumn',
  'totalRow',
];

/** The format the elements of `regions` that cover (row, col) compose to;
 *  undefined when none paints anything. */
export function composeStyleElements(
  formats: StyleElementFormats,
  order: readonly TableStyleElementType[],
  regions: readonly StyleElementRegion[],
  row: number,
  col: number,
): CellFormat | undefined {
  let out: CellFormat | undefined;
  for (const type of order) {
    const element = formats[type];
    if (!element) continue;
    for (const region of regions) {
      if (region.type !== type) continue;
      if (row < region.r0 || row > region.r1 || col < region.c0 || col > region.c1) continue;
      out ??= {};
      if (element.fill !== undefined) out.fill = element.fill;
      if (element.color !== undefined) out.color = element.color;
      if (element.bold !== undefined) out.bold = element.bold;
      const b = element.borders;
      if (!b) continue;
      const sides: CellBorders = {};
      if (b.top && row === region.r0) sides.top = b.top;
      if (b.bottom && row === region.r1) sides.bottom = b.bottom;
      if (b.left && col === region.c0) sides.left = b.left;
      if (b.right && col === region.c1) sides.right = b.right;
      if (Object.keys(sides).length > 0) out.borders = { ...out.borders, ...sides };
    }
  }
  return out;
}
