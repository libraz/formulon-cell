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

/** Side of a PivotTable field-caption button on the Mac profile, CSS px (14 pt). */
export const MAC_PIVOT_BUTTON_SIZE = 14;

/** A PivotTable field-caption button: `size` square, its trailing edge on the
 *  cell's trailing gridline and its top over the gridline above the cell, so
 *  it covers the cell's whole fill. */
export function pivotFieldButtonRect(cell: Rect, size: number, rule: number, rtl = false): Rect {
  const x = mirrorInRect(cell, cell.x + cell.w - rule - size, size, rtl);
  return { x, y: cell.y - rule, w: size, h: size };
}

/** Frame, face and arrow of a drop-down button. */
export interface FilterButtonLook {
  frame: string;
  /** Second frame ring one device pixel inside the first, on 2x displays. */
  innerFrame?: string;
  faceBottom: string;
  /** Highlight line under the top frame. */
  highlight?: string;
  arrow: string;
  arrowWidth: number;
  arrowHeight: number;
  /** Offset of the arrow's flat top from the button top. */
  arrowTop: number;
}

/** The AutoFilter button. */
export const FILTER_BUTTON_LOOK: FilterButtonLook = {
  frame: '#999999',
  faceBottom: '#f7f7f7',
  highlight: '#f9f9f9',
  arrow: '#838383',
  arrowWidth: 9,
  arrowHeight: 5.65,
  arrowTop: 6.25,
};

/** The Mac PivotTable field-caption button (measured at 2x). */
export const MAC_PIVOT_BUTTON_LOOK: FilterButtonLook = {
  frame: '#848a90',
  innerFrame: '#93989e',
  faceBottom: '#f5f5f5',
  arrow: '#7d7d7d',
  arrowWidth: 7,
  arrowHeight: 4.5,
  arrowTop: 4.75,
};

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
  look: FilterButtonLook = FILTER_BUTTON_LOOK,
  dpr = 1,
): void {
  const { x, y, w, h } = rect;
  ctx.save();
  const face = ctx.createLinearGradient(0, y, 0, y + h);
  face.addColorStop(0, '#ffffff');
  face.addColorStop(1, look.faceBottom);
  ctx.fillStyle = state.active ? theme.accent : face;
  ctx.fillRect(x, y, w, h);
  if (!state.active && look.highlight) {
    ctx.fillStyle = look.highlight;
    ctx.fillRect(x + 1, y + 1, w - 2, 0.5);
  }
  ctx.strokeStyle = state.active ? theme.accent : look.frame;
  const ring = look.innerFrame && dpr >= 2 ? 1 / Math.round(dpr) : 1;
  ctx.lineWidth = ring;
  ctx.strokeRect(x + ring / 2, y + ring / 2, w - ring, h - ring);
  if (look.innerFrame && ring < 1 && !state.active) {
    ctx.strokeStyle = look.innerFrame;
    ctx.strokeRect(x + ring * 1.5, y + ring * 1.5, w - ring * 3, h - ring * 3);
  }
  const ink = state.active ? theme.bg : look.arrow;
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
    const half = look.arrowWidth / 2;
    const top = y + look.arrowTop;
    ctx.moveTo(cx - half, top);
    ctx.lineTo(cx + half, top);
    ctx.lineTo(cx, top + look.arrowHeight);
    ctx.closePath();
    ctx.fill();
  }
  ctx.restore();
}
