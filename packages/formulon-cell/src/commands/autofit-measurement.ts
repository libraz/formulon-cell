/** Canvas text measurement behind the autofit row-height / column-width commands. */
import type { CellValue } from '../engine/types.js';
import { formatCell } from '../engine/value.js';
import type { CellFormat, State } from '../store/store.js';
import { formatNumber } from './format.js';

export function createAutofitMeasureContext(): CanvasRenderingContext2D | null {
  const doc = globalThis.document;
  const canvas = doc?.createElement?.('canvas');
  return canvas?.getContext?.('2d') ?? null;
}

export function computeAutofitColWidth(
  state: State,
  col: number,
  ctx: CanvasRenderingContext2D | null,
): number {
  const sheet = state.data.sheetIndex;
  const padding = 16;
  const minWidth = 48;
  let max = 0;

  for (const [key, cell] of state.data.cells) {
    const parsed = parseCellKey(key);
    if (!parsed || parsed.sheet !== sheet || parsed.col !== col) continue;
    const text = autofitDisplayText(state, key, cell);
    if (!text) continue;
    const fmt = state.format.formats.get(key);
    const fontSize = fmt?.fontSize ?? 13;
    if (ctx) ctx.font = autofitFont(fmt);
    const width =
      maxExplicitLineWidth(text, ctx, fontSize) +
      (isFilterHeaderCell(state, parsed.sheet, parsed.row, parsed.col) ? 28 : 0);
    if (width > max) max = width;
  }

  return Math.max(minWidth, Math.ceil(max) + padding);
}

export function computeAutofitRowHeight(
  state: State,
  row: number,
  ctx: CanvasRenderingContext2D | null,
): number {
  const sheet = state.data.sheetIndex;
  let max = state.layout.defaultRowHeight;

  for (const [key, cell] of state.data.cells) {
    const parsed = parseCellKey(key);
    if (!parsed || parsed.sheet !== sheet || parsed.row !== row) continue;
    const text = autofitDisplayText(state, key, cell);
    if (!text) continue;
    const fmt = state.format.formats.get(key);
    const fontSize = fmt?.fontSize ?? 13;
    if (ctx) ctx.font = autofitFont(fmt);
    const lineHeight = Math.round(fontSize * 1.28);
    const colW = state.layout.colWidths.get(parsed.col) ?? state.layout.defaultColWidth;
    const lines = autofitLineCount(text, fmt?.wrap === true, colW, ctx, fontSize);
    const height = Math.ceil(lines * lineHeight + 8);
    if (height > max) max = height;
  }

  return max;
}

function parseCellKey(key: string): { sheet: number; row: number; col: number } | null {
  const parts = key.split(':');
  if (parts.length !== 3) return null;
  const sheet = Number(parts[0]);
  const row = Number(parts[1]);
  const col = Number(parts[2]);
  if (!Number.isInteger(sheet) || !Number.isInteger(row) || !Number.isInteger(col)) return null;
  return { sheet, row, col };
}

function autofitDisplayText(
  state: State,
  key: string,
  cell: { value: CellValue; formula: string | null },
): string {
  if (state.ui.showFormulas && cell.formula) return cell.formula;
  const fmt = state.format.formats.get(key);
  if (cell.value.kind === 'number' && fmt?.numFmt)
    return formatNumber(cell.value.value, fmt.numFmt);
  return formatCell(cell.value);
}

function isFilterHeaderCell(state: State, sheet: number, row: number, col: number): boolean {
  const fr = state.ui.filterRange;
  return !!fr && fr.sheet === sheet && fr.r0 === row && col >= fr.c0 && col <= fr.c1;
}

function autofitFont(format: CellFormat | undefined): string {
  const styleSlant = format?.italic ? 'italic ' : '';
  const weight = format?.bold ? 700 : 400;
  const size = format?.fontSize ?? 13;
  const family = format?.fontFamily ?? 'system-ui, sans-serif';
  return `${styleSlant}${weight} ${size}px ${fontFamilyCss(family)}`;
}

function fontFamilyCss(family: string): string {
  return family
    .split(',')
    .map((part) => {
      const trimmed = part.trim();
      if (/^["'].*["']$/.test(trimmed) || /^[a-z-]+$/i.test(trimmed)) return trimmed;
      return `"${trimmed.replace(/"/g, '\\"')}"`;
    })
    .join(', ');
}

function maxExplicitLineWidth(
  text: string,
  ctx: CanvasRenderingContext2D | null,
  fontSize: number,
): number {
  let max = 0;
  for (const line of text.split(/\r\n|\r|\n/)) {
    const measured = ctx ? ctx.measureText(line).width : 0;
    const width = measured > 0 ? measured : line.length * fontSize * 0.54;
    if (width > max) max = width;
  }
  return max;
}

function autofitLineCount(
  text: string,
  wrap: boolean,
  colWidth: number,
  ctx: CanvasRenderingContext2D | null,
  fontSize: number,
): number {
  const paragraphs = text.split(/\r\n|\r|\n/);
  if (!wrap) return Math.max(1, paragraphs.length);

  const available = Math.max(1, colWidth - 12);
  let lines = 0;
  for (const paragraph of paragraphs) {
    if (paragraph.length === 0) {
      lines += 1;
      continue;
    }
    lines += wrapAutofitParagraph(paragraph, available, ctx, fontSize);
  }
  return Math.max(1, lines);
}

function wrapAutofitParagraph(
  paragraph: string,
  maxWidth: number,
  ctx: CanvasRenderingContext2D | null,
  fontSize: number,
): number {
  const words = paragraph.split(/(\s+)/);
  let line = '';
  let count = 0;
  for (const word of words) {
    const candidate = line + word;
    if (measureAutofitText(candidate, ctx, fontSize) <= maxWidth || line === '') {
      line = candidate;
    } else {
      count += 1;
      line = word.trimStart();
    }
  }
  return count + (line ? 1 : 0);
}

function measureAutofitText(
  text: string,
  ctx: CanvasRenderingContext2D | null,
  fontSize: number,
): number {
  const measured = ctx ? ctx.measureText(text).width : 0;
  return measured > 0 ? measured : text.length * fontSize * 0.54;
}

/** Resolve which row/col indices to show again from the current selection.
 *  Spreadsheets return visible rows that flank a hidden band; we emulate by
 *  reporting every hidden row inside the selection. */
