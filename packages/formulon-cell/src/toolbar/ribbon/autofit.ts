// Off-screen canvas-based autofit measurement. Used when the user picks the
// "AutoFit row height" / "AutoFit column width" menu items. The default
// rendering pipeline doesn't lay out cells to pixel widths so we re-measure
// here with the same fonts the renderer would use.

import { formatNumber } from '../../commands/format.js';
import { formatCell } from '../../engine/value.js';
import type { SpreadsheetInstance } from '../../mount/types.js';
import type { NumFmt } from '../../store/types.js';

const autofitMeasureCanvas = document.createElement('canvas');
const autofitMeasureCtx = autofitMeasureCanvas.getContext('2d');

export interface AutofitCellFormat {
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  numFmt?: NumFmt;
  wrap?: boolean;
}

const cssFontFamily = (family: string): string =>
  family
    .split(',')
    .map((part) => {
      const trimmed = part.trim();
      if (/^["'].*["']$/.test(trimmed) || /^[a-z-]+$/i.test(trimmed)) return trimmed;
      return `"${trimmed.replace(/"/g, '\\"')}"`;
    })
    .join(', ');

const autofitFont = (format: AutofitCellFormat | undefined): string => {
  const size = format?.fontSize ?? 13;
  const weight = format?.bold ? 700 : 400;
  const slant = format?.italic ? 'italic ' : '';
  const family = cssFontFamily(format?.fontFamily ?? 'system-ui, sans-serif');
  return `${slant}${weight} ${size}px ${family}`;
};

const measureAutofitText = (text: string, fontSize: number): number => {
  const measured = autofitMeasureCtx?.measureText(text).width ?? 0;
  // Fall back to a rough character-width estimate when the canvas API
  // isn't available (jsdom in some test runners).
  return measured > 0 ? measured : text.length * fontSize * 0.54;
};

const cellDisplayText = (
  i: SpreadsheetInstance,
  row: number,
  col: number,
  locale: 'ja' | 'en',
): string => {
  const state = i.store.getState();
  const sheet = state.data.sheetIndex;
  const key = `${sheet}:${row}:${col}`;
  const cell = state.data.cells.get(key);
  if (!cell) return '';
  const formula = i.workbook.cellFormula({ sheet, row, col });
  const value = formula ? i.workbook.getValue({ sheet, row, col }) : cell.value;
  const fmt = state.format.formats.get(key);
  if (value.kind === 'number' && fmt?.numFmt) return formatNumber(value.value, fmt.numFmt);
  return formatCell(value, locale === 'ja' ? 'ja-JP' : 'en-US');
};

export const autofitColWidth = (
  i: SpreadsheetInstance,
  col: number,
  r0: number,
  r1: number,
  locale: 'ja' | 'en',
): number => {
  const state = i.store.getState();
  const sheet = state.data.sheetIndex;
  let width = state.layout.defaultColWidth;
  for (const [key] of state.data.cells) {
    const [s, r, c] = key.split(':').map(Number);
    if (s === undefined || r === undefined || c === undefined) continue;
    if (s !== sheet || c !== col || r < r0 || r > r1) continue;
    const fmt = state.format.formats.get(key);
    const fontSize = fmt?.fontSize ?? 13;
    if (autofitMeasureCtx) autofitMeasureCtx.font = autofitFont(fmt);
    const text = cellDisplayText(i, r, c, locale);
    for (const line of text.split(/\r\n|\r|\n/)) {
      width = Math.max(width, Math.ceil(measureAutofitText(line, fontSize) + 14));
    }
  }
  return Math.max(12, Math.min(512, width));
};

const wrappedLineCount = (text: string, maxWidth: number, fontSize: number): number => {
  const paragraphs = text.split(/\r\n|\r|\n/);
  let total = 0;
  for (const paragraph of paragraphs) {
    if (!paragraph) {
      total += 1;
      continue;
    }
    const words = paragraph.split(/(\s+)/);
    let line = '';
    let count = 0;
    for (const word of words) {
      const next = line + word;
      if (measureAutofitText(next, fontSize) <= maxWidth || line === '') line = next;
      else {
        count += 1;
        line = word.trimStart();
      }
    }
    total += count + (line ? 1 : 0);
  }
  return Math.max(1, total);
};

export const autofitRowHeight = (
  i: SpreadsheetInstance,
  row: number,
  c0: number,
  c1: number,
  locale: 'ja' | 'en',
): number => {
  const state = i.store.getState();
  const sheet = state.data.sheetIndex;
  let height = state.layout.defaultRowHeight;
  for (const [key] of state.data.cells) {
    const [s, r, c] = key.split(':').map(Number);
    if (s === undefined || r === undefined || c === undefined) continue;
    if (s !== sheet || r !== row || c < c0 || c > c1) continue;
    const fmt = state.format.formats.get(key);
    const fontSize = fmt?.fontSize ?? 13;
    const colWidth = state.layout.colWidths.get(c) ?? state.layout.defaultColWidth;
    if (autofitMeasureCtx) autofitMeasureCtx.font = autofitFont(fmt);
    const text = cellDisplayText(i, r, c, locale);
    const lines =
      fmt?.wrap === true
        ? wrappedLineCount(text, Math.max(1, colWidth - 12), fontSize)
        : Math.max(1, text.split(/\r\n|\r|\n/).length);
    height = Math.max(height, Math.ceil(lines * Math.round(fontSize * 1.28) + 8));
  }
  return Math.max(8, Math.min(409, height));
};
