import { mirrorInRect, type Rect } from '../geometry.js';

/** Horizontal step of one PivotTable outline level, in points. A compact row
 *  label's text starts this many points further in per step. */
export const PIVOT_OUTLINE_STEP_PT = 11.5;

/** Side of the collapse (-) button, in points. */
export const PIVOT_COLLAPSE_BUTTON_PT = 7;

/** The collapse button's offset from the leading edge of its outline step, in points. */
const COLLAPSE_INSET_PT = 4;

/** Leading inset of a compact row label's text for `steps` outline steps, in screen px. */
export const pivotOutlineIndent = (steps: number, ptPx: number): number =>
  steps * PIVOT_OUTLINE_STEP_PT * ptPx;

/** The collapse button of a label `steps` deep (its own step is the last one),
 *  centred in the row the way the spreadsheet places it. `ptPx` is screen px
 *  per point. */
export function pivotCollapseButtonRect(
  cell: Rect,
  steps: number,
  ptPx: number,
  rtl = false,
): Rect {
  const size = PIVOT_COLLAPSE_BUTTON_PT * ptPx;
  const left = cell.x + (COLLAPSE_INSET_PT + Math.max(0, steps - 1) * PIVOT_OUTLINE_STEP_PT) * ptPx;
  return {
    x: mirrorInRect(cell, left, size, rtl),
    y: cell.y + (cell.h - size + ptPx) / 2,
    w: size,
    h: size,
  };
}

/** Colours of the collapse button (measured on the Mac desktop at 2x). */
export const PIVOT_COLLAPSE_LOOK = {
  frameLead: '#aab5c9',
  frameTrail: '#8b9ebf',
  faceTop: '#e7e8e8',
  faceBottom: '#dedfdf',
  bar: '#3d4e7c',
} as const;

/** Paint the collapse (-) button into `rect`: a gradient face with a two-tone
 *  frame whose corner pixels stay open, and a horizontal minus bar. */
export function paintPivotCollapseButton(
  ctx: CanvasRenderingContext2D,
  rect: Rect,
  ptPx: number,
  dpr = 1,
): void {
  const { x, y, w, h } = rect;
  const px = 1 / Math.max(1, Math.round(dpr));
  const look = PIVOT_COLLAPSE_LOOK;
  ctx.save();
  const face = ctx.createLinearGradient(0, y, 0, y + h);
  face.addColorStop(0, look.faceTop);
  face.addColorStop(1, look.faceBottom);
  ctx.fillStyle = face;
  ctx.fillRect(x + px, y + px, w - 2 * px, h - 2 * px);
  ctx.fillStyle = look.frameLead;
  ctx.fillRect(x + px, y, w - 2 * px, px);
  ctx.fillRect(x, y + px, px, h - 2 * px);
  ctx.fillStyle = look.frameTrail;
  ctx.fillRect(x + px, y + h - px, w - 2 * px, px);
  ctx.fillRect(x + w - px, y + px, px, h - 2 * px);
  const barW = 3 * ptPx;
  const barH = ptPx;
  ctx.fillStyle = look.bar;
  ctx.fillRect(x + (w - barW) / 2, y + (h - barH) / 2, barW, barH);
  ctx.restore();
}
