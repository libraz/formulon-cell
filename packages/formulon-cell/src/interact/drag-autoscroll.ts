/** Tick period of the autoscroll timer. */
export const AUTOSCROLL_TICK_MS = 50;
/** Autoscroll starts this far inside the body edge, so the last visible row or column already scrolls. */
export const AUTOSCROLL_EDGE_INSET_PX = 4;
/** Rows or columns per second once the pointer sits on the edge. */
export const AUTOSCROLL_BASE_RATE = 12;
/** Extra rows or columns per second for each pixel past the inset edge. */
export const AUTOSCROLL_RATE_PER_PX = 0.5;
/** Ceiling of the rate, in rows or columns per second. */
export const AUTOSCROLL_MAX_RATE = 60;

export interface AutoscrollBounds {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

export interface AutoscrollPoint {
  x: number;
  y: number;
}

export interface DragAutoscrollOptions {
  scrollBy: (dRow: number, dCol: number) => void;
  /** Called after each scroll step with the last pointer position. */
  onScrolled: (point: AutoscrollPoint) => void;
}

export interface DragAutoscroll {
  /** Feed the pointer position; starts, retunes or stops the timer. `axes` limits which directions scroll. */
  update(
    point: AutoscrollPoint,
    bounds: AutoscrollBounds,
    axes: { row: boolean; col: boolean },
  ): void;
  stop(): void;
}

/** Signed rows/s for a coordinate against the [lo, hi] span; zero while inside it. */
function axisRate(v: number, lo: number, hi: number): number {
  const past =
    v < lo + AUTOSCROLL_EDGE_INSET_PX
      ? lo + AUTOSCROLL_EDGE_INSET_PX - v
      : v - (hi - AUTOSCROLL_EDGE_INSET_PX);
  if (past < 0) return 0;
  const rate = Math.min(AUTOSCROLL_MAX_RATE, AUTOSCROLL_BASE_RATE + past * AUTOSCROLL_RATE_PER_PX);
  return v < lo + AUTOSCROLL_EDGE_INSET_PX ? -rate : rate;
}

export function createDragAutoscroll(opts: DragAutoscrollOptions): DragAutoscroll {
  let timer: ReturnType<typeof setInterval> | null = null;
  let rowRate = 0;
  let colRate = 0;
  let rowResidual = 0;
  let colResidual = 0;
  let last: AutoscrollPoint = { x: 0, y: 0 };

  const stop = (): void => {
    if (timer !== null) clearInterval(timer);
    timer = null;
    rowRate = colRate = rowResidual = colResidual = 0;
  };

  const tick = (): void => {
    rowResidual += (rowRate * AUTOSCROLL_TICK_MS) / 1000;
    colResidual += (colRate * AUTOSCROLL_TICK_MS) / 1000;
    const dRow = Math.trunc(rowResidual);
    const dCol = Math.trunc(colResidual);
    rowResidual -= dRow;
    colResidual -= dCol;
    if (dRow === 0 && dCol === 0) return;
    opts.scrollBy(dRow, dCol);
    opts.onScrolled(last);
  };

  return {
    update(point, bounds, axes) {
      last = point;
      const nextRow = axes.row ? axisRate(point.y, bounds.top, bounds.bottom) : 0;
      const nextCol = axes.col ? axisRate(point.x, bounds.left, bounds.right) : 0;
      if (nextRow === 0 && nextCol === 0) {
        stop();
        return;
      }
      rowRate = nextRow;
      colRate = nextCol;
      if (timer === null) timer = setInterval(tick, AUTOSCROLL_TICK_MS);
    },
    stop,
  };
}
