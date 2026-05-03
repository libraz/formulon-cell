/**
 * High-level import/export commands for tabular text formats. The lower-level
 * encoders / parsers live next to clipboard so paste-as-CSV stays trivial.
 */

import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { State } from '../store/store.js';
import { type CSVEncodeOptions, encodeCSV, parseCSV } from './clipboard/csv.js';
import { coerceInput, writeCoerced } from './coerce-input.js';

export interface ImportResult {
  /** Range that received writes. r0/c0 = anchor; r1/c1 = far corner. */
  writtenRange: Range;
  /** Total cells written (including blanks that overwrote prior values). */
  cellsWritten: number;
  /** Rows parsed from the input (== writtenRange.r1 - writtenRange.r0 + 1). */
  rows: number;
}

/**
 * Parse a CSV blob and write the resulting cells starting at `anchor` (or the
 * active cell when omitted). Each value runs through `coerceInput` so a leading
 * `=` becomes a formula, `"42"` becomes a number, etc. — same semantics as
 * paste-from-clipboard, but for full-document loads.
 */
export function importCSV(
  state: State,
  wb: WorkbookHandle,
  text: string,
  anchor?: Addr,
): ImportResult | null {
  if (!text) return null;
  const rows = parseCSV(text);
  if (rows.length === 0) return null;

  const origin: Addr = anchor ?? state.selection.active;
  const sheet = origin.sheet;
  let maxCols = 0;
  let cellsWritten = 0;

  for (let r = 0; r < rows.length; r += 1) {
    const cells = rows[r] ?? [];
    if (cells.length > maxCols) maxCols = cells.length;
    for (let c = 0; c < cells.length; c += 1) {
      const addr: Addr = { sheet, row: origin.row + r, col: origin.col + c };
      writeCoerced(wb, addr, coerceInput(cells[c] ?? ''));
      cellsWritten += 1;
    }
  }

  return {
    writtenRange: {
      sheet,
      r0: origin.row,
      c0: origin.col,
      r1: origin.row + rows.length - 1,
      c1: origin.col + Math.max(0, maxCols - 1),
    },
    cellsWritten,
    rows: rows.length,
  };
}

export interface ExportOptions extends CSVEncodeOptions {
  /** Range to export. Defaults to the current selection. When the selection is
   *  a single cell, the entire used region of the sheet is exported instead. */
  range?: Range;
}

/**
 * Serialize a range to CSV. Values come from the displayed text (formulas
 * collapse to their last computed result) — same rules as TSV copy. When no
 * explicit range is supplied and the selection is a single cell, the export
 * covers the bounding box of populated cells on the active sheet.
 */
export function exportCSV(state: State, opts: ExportOptions = {}): string {
  const range = opts.range ?? selectionOrUsed(state);
  if (!range) return '';
  const { sheet } = range;

  const grid: string[][] = [];
  for (let row = range.r0; row <= range.r1; row += 1) {
    const line: string[] = [];
    for (let col = range.c0; col <= range.c1; col += 1) {
      line.push(displayValue(state, sheet, row, col));
    }
    grid.push(line);
  }
  return encodeCSV(grid, { eol: opts.eol, bom: opts.bom });
}

function selectionOrUsed(state: State): Range | null {
  const sel = state.selection.range;
  const isSingle = sel.r0 === sel.r1 && sel.c0 === sel.c1;
  if (!isSingle) return sel;
  return usedRange(state);
}

/** Bounding box of every populated cell on the active sheet. Returns null
 *  when the sheet is empty. */
function usedRange(state: State): Range | null {
  const sheet = state.data.sheetIndex;
  let r0 = Number.POSITIVE_INFINITY;
  let r1 = Number.NEGATIVE_INFINITY;
  let c0 = Number.POSITIVE_INFINITY;
  let c1 = Number.NEGATIVE_INFINITY;
  let any = false;
  for (const [key, cell] of state.data.cells) {
    if (cell.value.kind === 'blank' && !cell.formula) continue;
    const parts = key.split(':');
    if (parts.length !== 3) continue;
    if (Number(parts[0]) !== sheet) continue;
    const row = Number(parts[1]);
    const col = Number(parts[2]);
    if (row < r0) r0 = row;
    if (row > r1) r1 = row;
    if (col < c0) c0 = col;
    if (col > c1) c1 = col;
    any = true;
  }
  if (!any) return null;
  return { sheet, r0, c0, r1, c1 };
}

function displayValue(state: State, sheet: number, row: number, col: number): string {
  const cell = state.data.cells.get(`${sheet}:${row}:${col}`);
  if (!cell) return '';
  switch (cell.value.kind) {
    case 'number':
      return String(cell.value.value);
    case 'text':
      return cell.value.value;
    case 'bool':
      return cell.value.value ? 'TRUE' : 'FALSE';
    case 'error':
      return cell.value.text;
    default:
      return '';
  }
}
