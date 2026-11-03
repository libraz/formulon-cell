import type { ResolvedTheme } from '../../theme/resolve.js';
import { mirrorInRect, type Rect } from '../geometry.js';

/** Side of the in-cell AutoFilter button, CSS px. */
export const FILTER_BUTTON_SIZE = 17;

/** The button's rectangle in a header cell of a filter range: right-aligned,
 *  one CSS px inside the cell's trailing edge and one below its top (leading
 *  edge when `rtl`). Also the click target. */
export function filterButtonRect(cell: Rect, rtl = false): Rect {
  const x = mirrorInRect(cell, cell.x + cell.w - 1 - FILTER_BUTTON_SIZE, FILTER_BUTTON_SIZE, rtl);
  return { x, y: cell.y + 1, w: FILTER_BUTTON_SIZE, h: FILTER_BUTTON_SIZE };
}

const FRAME = '#999999';
const ARROW = '#838383';

export interface FilterButtonState {
  /** Rows are hidden by a filter on this sheet. */
  active?: boolean;
  /** Sort direction of the column, when sorted. */
  sorted?: 'asc' | 'desc';
}

/** Paint the AutoFilter button into `rect`: a framed gradient face with a
 *  drop-down arrow, or a stemmed arrow for a sorted column. */
export function paintFilterButton(
  ctx: CanvasRenderingContext2D,
  rect: Rect,
  theme: ResolvedTheme,
  state: FilterButtonState = {},
): void {
  const { x, y, w, h } = rect;
  ctx.save();
  const face = ctx.createLinearGradient(0, y, 0, y + h);
  face.addColorStop(0, '#ffffff');
  face.addColorStop(1, '#f7f7f7');
  ctx.fillStyle = state.active ? theme.accent : face;
  ctx.fillRect(x, y, w, h);
  if (!state.active) {
    ctx.fillStyle = '#f9f9f9';
    ctx.fillRect(x + 1, y + 1, w - 2, 0.5);
  }
  ctx.strokeStyle = state.active ? theme.accent : FRAME;
  ctx.lineWidth = 1;
  ctx.strokeRect(x + 0.5, y + 0.5, w - 1, h - 1);
  const ink = state.active ? theme.bg : ARROW;
  const cx = x + w / 2;
  ctx.fillStyle = ink;
  ctx.beginPath();
  if (state.sorted) {
    const up = state.sorted === 'asc';
    const tip = y + (up ? 4 : h - 4);
    const tail = y + (up ? h - 4 : 4);
    const wing = up ? tip + 3.5 : tip - 3.5;
    ctx.strokeStyle = ink;
    ctx.lineWidth = 1.25;
    ctx.moveTo(cx, tail);
    ctx.lineTo(cx, tip);
    ctx.moveTo(cx - 3.5, wing);
    ctx.lineTo(cx, tip);
    ctx.lineTo(cx + 3.5, wing);
    ctx.stroke();
  } else {
    ctx.moveTo(cx - 4.5, y + 6.25);
    ctx.lineTo(cx + 4.5, y + 6.25);
    ctx.lineTo(cx, y + 11.9);
    ctx.closePath();
    ctx.fill();
  }
  ctx.restore();
}
