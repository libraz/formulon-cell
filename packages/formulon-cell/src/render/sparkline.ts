import { resolveNumericRange, resolveNumericRangeFromCells } from '../engine/range-resolver.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Sparkline, State } from '../store/store.js';
import { paintSparkline } from './painters.js';

/** Cell-paint helper — resolves the source series and draws the sparkline.
 *  Same-sheet refs read from `state.data.cells` so the painter still works
 *  in tests with no engine. Cross-sheet refs need a workbook handle. */
export function paintCellSparkline(
  ctx: CanvasRenderingContext2D,
  bounds: { x: number; y: number; w: number; h: number },
  spec: Sparkline,
  state: State,
  wb: WorkbookHandle | null,
): void {
  const sheet = state.data.sheetIndex;
  let values: number[];
  if (wb) {
    values = resolveNumericRange(wb, spec.source, sheet);
  } else {
    values = resolveNumericRangeFromCells(state.data.cells, spec.source, sheet);
  }
  paintSparkline(ctx, bounds, spec, values);
}
