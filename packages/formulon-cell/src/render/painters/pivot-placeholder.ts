import type { Rect } from '../geometry.js';

/** Measured on the Mac desktop (points, sRGB). */
const FRAME = '#999999';
const TAG_INK = '#969696';
const TAG = { left: 18, top: 20, w: 120, h: 17, frame: 2 };
const PROMPT = { firstCentre: 53, pitch: 14, inset: 4 };

/** Break `text` into lines no wider than `width`, at any character. */
export function wrapPlaceholderText(
  text: string,
  width: number,
  measure: (s: string) => number,
): string[] {
  const lines: string[] = [];
  let line = '';
  for (const ch of text) {
    if (line && measure(line + ch) > width) {
      lines.push(line);
      line = ch === ' ' ? '' : ch;
    } else {
      line += ch;
    }
  }
  if (line) lines.push(line);
  return lines;
}

export interface PivotPlaceholderPaint {
  /** The reserved area, cell boundaries in screen px. */
  box: Rect;
  /** Report name shown in the tag. */
  name: string;
  /** Prompt shown under the tag. */
  prompt: string;
  /** Screen px per point. */
  ptPx: number;
  /** CSS font family and size (points) of the sheet's Normal font. */
  fontFamily: string;
  fontSizePt: number;
  /** Sheet background and default ink. */
  bg: string;
  ink: string;
}

/** Paint the empty-report placeholder: the reserved area cleared to the sheet
 *  background inside a 1 pt grey frame on its gridlines, the report-name tag,
 *  and the prompt centred and wrapped below it. */
export function paintPivotPlaceholder(
  ctx: CanvasRenderingContext2D,
  p: PivotPlaceholderPaint,
): void {
  const { box, ptPx } = p;
  const rule = ptPx;
  ctx.save();
  ctx.fillStyle = p.bg;
  ctx.fillRect(box.x, box.y, box.w - rule, box.h - rule);
  ctx.strokeStyle = FRAME;
  ctx.lineWidth = rule;
  ctx.strokeRect(box.x - rule / 2, box.y - rule / 2, box.w, box.h);

  const tag = {
    x: box.x + TAG.left * ptPx,
    y: box.y + TAG.top * ptPx,
    w: TAG.w * ptPx,
    h: TAG.h * ptPx,
  };
  const frame = TAG.frame * ptPx;
  ctx.fillStyle = p.bg;
  ctx.fillRect(tag.x, tag.y, tag.w, tag.h);
  ctx.strokeStyle = TAG_INK;
  ctx.lineWidth = frame;
  ctx.strokeRect(tag.x + frame / 2, tag.y + frame / 2, tag.w - frame, tag.h - frame);

  ctx.font = `${p.fontSizePt * ptPx}px ${p.fontFamily}`;
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillStyle = TAG_INK;
  ctx.fillText(p.name, tag.x + tag.w / 2, tag.y + tag.h / 2, tag.w - 2 * frame);

  const width = box.w - 2 * PROMPT.inset * ptPx;
  const lines = wrapPlaceholderText(p.prompt, width, (s) => ctx.measureText(s).width);
  ctx.fillStyle = p.ink;
  const cx = box.x + box.w / 2;
  for (const [i, line] of lines.entries()) {
    ctx.fillText(line, cx, box.y + (PROMPT.firstCentre + i * PROMPT.pitch) * ptPx);
  }
  ctx.restore();
}
