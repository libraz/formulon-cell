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
