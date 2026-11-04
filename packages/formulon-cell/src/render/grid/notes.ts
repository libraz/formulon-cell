// Shown notes: a yellow box beside the anchor cell with a leader line back to
// it, clipped to the cell's freeze / split pane. The engine exposes no note
// geometry, so the box takes the desktop default for a new note (15 px right
// of the cell, 144 x 79 px at 100%), or the Mac card's measured points; a
// right-to-left sheet mirrors it.

import { parseAddrKey } from '../../engine/address.js';
import { fontPointsToCssPx } from '../../engine/font-units.js';
import type { Addr } from '../../engine/types.js';
import { renderScale } from '../../store/render-scale.js';
import type { CellFormat } from '../../store/store.js';
import {
  type AxisLayout,
  cellRectIn,
  gridOriginX,
  gridOriginY,
  isColVisible,
  isRowVisible,
  mirrorX,
  type Rect,
  type ViewLayout,
  type ViewState,
} from '../geometry.js';
import { wholePointAdvances } from '../text-metrics.js';
import { wrapText } from '../text-wrap.js';

export const NOTE_BOX_FILL = '#ffffe1';
const NOTE_BOX_STROKE = '#1f1f1f';
const NOTE_OFFSET_X = 15;
const NOTE_OFFSET_Y = 7;
const NOTE_WIDTH = 144;
const NOTE_HEIGHT = 79;
const NOTE_PAD = 4;
const NOTE_FONT_PX = 12;

/** The Mac desktop's card, in points (Excel Mac at 2x: 105 x 77 incl. a 1 pt
 *  black frame, 13.5 right of and 11.5 above the cell corner, bold 9 pt text
 *  2.5 in with its baseline 15.5 down). */
const MAC_NOTE = {
  offsetX: 13.5,
  offsetY: 11.5,
  width: 105,
  height: 77,
  frame: '#000000',
  textInset: 2.5,
  baseline: 15.5,
  fontPt: 9,
  lineHeight: 11,
} as const;

export interface VisibleNote {
  addr: Addr;
  text: string;
  author?: string;
}

// The store replaces the formats map on every change, so its identity keys
// the shown-note index and an unchanged sheet is never rescanned.
const noteIndex = new WeakMap<ReadonlyMap<string, CellFormat>, readonly VisibleNote[]>();

/** Every shown note in `formats`, cached per map. */
export function visibleNotes(formats: ReadonlyMap<string, CellFormat>): readonly VisibleNote[] {
  const cached = noteIndex.get(formats);
  if (cached) return cached;
  const notes: VisibleNote[] = [];
  for (const [key, format] of formats) {
    if (format.commentVisible !== true || !format.comment) continue;
    const addr = parseAddrKey(key);
    if (!addr) continue;
    notes.push({
      addr,
      text: format.comment,
      ...(format.commentAuthor ? { author: format.commentAuthor } : {}),
    });
  }
  noteIndex.set(formats, notes);
  return notes;
}

/** Screen rect of the freeze / split pane holding the cell, headers excluded. */
function paneRect(
  layout: ViewLayout,
  cols: AxisLayout,
  rows: AxisLayout,
  addr: Addr,
  bodyEnd: number,
): Rect {
  const originX = gridOriginX(layout);
  const originY = gridOriginY(layout);
  const inFrozenCols = (cols.positionAt.get(addr.col) ?? 0) < cols.frozenTotal;
  const inFrozenRows = (rows.positionAt.get(addr.row) ?? 0) < rows.frozenTotal;
  const x0 = originX + (inFrozenCols ? 0 : cols.frozenTotal);
  const x1 = inFrozenCols ? originX + cols.frozenTotal : Math.max(layout.viewWidth, x0);
  const y0 = originY + (inFrozenRows ? 0 : rows.frozenTotal);
  const y1 = originY + (inFrozenRows ? rows.frozenTotal : bodyEnd);
  return { x: mirrorX(layout, x0, x1 - x0), y: y0, w: x1 - x0, h: y1 - y0 };
}

/** Bottom edge of the scrolling body: the last visible row's end. */
function bodyEndOf(rows: AxisLayout): number {
  let bodyEnd = rows.frozenTotal;
  for (const r of rows.visible) {
    bodyEnd = Math.max(bodyEnd, (rows.positionAt.get(r) ?? 0) + (rows.sizeAt.get(r) ?? 0));
  }
  return bodyEnd;
}

export function paintVisibleNotes(
  ctx: CanvasRenderingContext2D,
  state: ViewState,
  cols: AxisLayout,
  rows: AxisLayout,
): void {
  const { layout, viewport } = state;
  const sheet = state.data.sheetIndex;
  const scale = renderScale(layout, viewport);
  const rtl = layout.rtl === true;
  const macPt = wholePointAdvances(layout) ? fontPointsToCssPx(1) * scale : null;
  const notes = visibleNotes(state.format.formats);
  if (notes.length === 0) return;
  const bodyEnd = bodyEndOf(rows);
  for (const note of notes) {
    const { addr } = note;
    if (addr.sheet !== sheet) continue;
    if (!isRowVisible(layout, viewport, addr.row) || !isColVisible(layout, viewport, addr.col)) {
      continue;
    }
    const cell = cellRectIn(layout, cols, rows, addr.row, addr.col);
    const pane = paneRect(layout, cols, rows, addr, bodyEnd);
    if (macPt !== null) {
      paintMacNote(ctx, note, cell, pane, macPt, rtl);
      continue;
    }
    const w = NOTE_WIDTH * scale;
    const h = NOTE_HEIGHT * scale;
    const cornerX = rtl ? cell.x : cell.x + cell.w;
    const x = rtl ? cornerX - NOTE_OFFSET_X * scale - w : cornerX + NOTE_OFFSET_X * scale;
    const y = Math.max(pane.y, cell.y - NOTE_OFFSET_Y * scale);
    ctx.save();
    ctx.beginPath();
    ctx.rect(pane.x, pane.y, pane.w, pane.h);
    ctx.clip();
    ctx.strokeStyle = NOTE_BOX_STROKE;
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(cornerX, cell.y);
    ctx.lineTo(rtl ? x + w : x, y + 2 * scale);
    ctx.stroke();
    ctx.fillStyle = NOTE_BOX_FILL;
    ctx.fillRect(x, y, w, h);
    ctx.strokeRect(x + 0.5, y + 0.5, w - 1, h - 1);
    ctx.beginPath();
    ctx.rect(x, y, w, h);
    ctx.clip();
    const fontPx = NOTE_FONT_PX * scale;
    const pad = NOTE_PAD * scale;
    ctx.fillStyle = NOTE_BOX_STROKE;
    ctx.textBaseline = 'top';
    ctx.textAlign = 'left';
    let lineY = y + pad;
    if (note.author) {
      ctx.font = `bold ${fontPx}px Tahoma, sans-serif`;
      ctx.fillText(`${note.author}:`, x + pad, lineY);
      lineY += fontPx * 1.25;
    }
    ctx.font = `${fontPx}px Tahoma, sans-serif`;
    const lines = wrapText(note.text, {
      maxWidth: Math.max(1, w - pad * 2),
      measureText: (text) => ctx.measureText(text).width,
    });
    for (const line of lines) {
      if (lineY > y + h) break;
      ctx.fillText(line, x + pad, lineY);
      lineY += fontPx * 1.25;
    }
    ctx.restore();
  }
}

/** The Mac card: the whole note set bold in black, with no author line. */
function paintMacNote(
  ctx: CanvasRenderingContext2D,
  note: VisibleNote,
  cell: Rect,
  pane: Rect,
  pt: number,
  rtl: boolean,
): void {
  const w = MAC_NOTE.width * pt;
  const h = MAC_NOTE.height * pt;
  const cornerX = rtl ? cell.x : cell.x + cell.w;
  const x = rtl ? cornerX - MAC_NOTE.offsetX * pt - w : cornerX + MAC_NOTE.offsetX * pt;
  const y = Math.max(pane.y, cell.y - MAC_NOTE.offsetY * pt);
  ctx.save();
  ctx.beginPath();
  ctx.rect(pane.x, pane.y, pane.w, pane.h);
  ctx.clip();
  ctx.strokeStyle = MAC_NOTE.frame;
  ctx.lineWidth = pt;
  ctx.beginPath();
  ctx.moveTo(cornerX, cell.y);
  ctx.lineTo(rtl ? x + w : x, y);
  ctx.stroke();
  ctx.fillStyle = NOTE_BOX_FILL;
  ctx.fillRect(x, y, w, h);
  ctx.strokeRect(x + pt / 2, y + pt / 2, w - pt, h - pt);
  ctx.beginPath();
  ctx.rect(x, y, w, h);
  ctx.clip();
  ctx.fillStyle = MAC_NOTE.frame;
  ctx.font = `bold ${MAC_NOTE.fontPt * pt}px Tahoma, sans-serif`;
  ctx.textBaseline = 'alphabetic';
  ctx.textAlign = rtl ? 'right' : 'left';
  const inset = MAC_NOTE.textInset * pt;
  const lines = wrapText(note.text, {
    maxWidth: Math.max(1, w - inset * 2),
    measureText: (text) => ctx.measureText(text).width,
  });
  let lineY = y + MAC_NOTE.baseline * pt;
  for (const line of lines) {
    if (lineY > y + h) break;
    ctx.fillText(line, rtl ? x + w - inset : x + inset, lineY);
    lineY += MAC_NOTE.lineHeight * pt;
  }
  ctx.restore();
}
