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

/** A report whose layout holds nothing but the caption the engine projects for
 *  a report with no fields: it is drawn as a placeholder and shows no cells. */
export function isEmptyPivotLayout(cells: readonly { kind: number }[]): boolean {
  return cells.every((c) => c.kind === PIVOT_CELL_KIND.Header);
}

/** An empty report reserves this many columns and rows from its anchor. */
export const PIVOT_PLACEHOLDER_COLS = 3;
export const PIVOT_PLACEHOLDER_ROWS = 18;
