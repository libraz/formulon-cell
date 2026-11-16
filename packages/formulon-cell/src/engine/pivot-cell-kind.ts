import type { CellValue } from './types.js';

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

/** `PivotCell.expansion` ordinals: the state of an item label's expand/collapse button.
 *  `null` marks a cell without one (innermost label, subtotal, total, header, data). */
export const PIVOT_CELL_EXPANSION = {
  Expanded: 0,
  Collapsed: 1,
} as const;

export type PivotCellExpansion = (typeof PIVOT_CELL_EXPANSION)[keyof typeof PIVOT_CELL_EXPANSION];

/** A report whose layout holds nothing but the caption the engine projects for
 *  a report with no fields: it is drawn as a placeholder and shows no cells. */
export function isEmptyPivotLayout(cells: readonly { kind: number }[]): boolean {
  return cells.every((c) => c.kind === PIVOT_CELL_KIND.Header);
}

/** An empty report reserves this many columns and rows from its anchor. */
export const PIVOT_PLACEHOLDER_COLS = 3;
export const PIVOT_PLACEHOLDER_ROWS = 18;

/** The text of a report item label as the report shows it; empty for anything but text or a number. */
export const pivotItemText = (value: CellValue): string =>
  value.kind === 'text' || value.kind === 'number' ? String(value.value) : '';
