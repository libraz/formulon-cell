import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';

/** Pivot reports saved by the desktop app (Japanese UI). */
export const FIXTURES = join(
  dirname(fileURLToPath(import.meta.url)),
  '../../../fixtures/pivot-excel',
);

/** Load a fixture under the Japanese Mac profile the facts were measured in. */
export const loadFixture = async (file: string): Promise<WorkbookHandle> => {
  const wb = await WorkbookHandle.loadBytes(new Uint8Array(readFileSync(join(FIXTURES, file))));
  wb.setSpreadsheetProfileId('mac-ja_JP');
  return wb;
};

const a1 = (row: number, col: number): string => `${String.fromCharCode(65 + col)}${row + 1}`;

/** The rendered report of a sheet, A1 ref to value; blank cells are absent. */
export const pivotGrid = (
  wb: WorkbookHandle,
  sheet = 0,
): Record<string, string | number | boolean> => {
  const out: Record<string, string | number | boolean> = {};
  for (const cell of wb.pivotCells(sheet)) {
    const v = cell.value;
    out[a1(cell.addr.row, cell.addr.col)] =
      v.kind === 'error'
        ? v.text
        : v.kind === 'blank'
          ? ''
          : (v.value as string | number | boolean);
  }
  return out;
};

/** Column `col` (letter) of a grid, top to bottom, blanks skipped. */
export const column = (grid: Record<string, unknown>, col: string): unknown[] =>
  Object.entries(grid)
    .filter(([ref]) => ref.startsWith(col) && /^\d+$/.test(ref.slice(col.length)))
    .sort(([a], [b]) => Number(a.slice(col.length)) - Number(b.slice(col.length)))
    .map(([, v]) => v);

/** Number formats of a column's report cells. */
export const columnFormats = (wb: WorkbookHandle, col: number, sheet = 0): string[] =>
  Array.from(wb.pivotCells(sheet))
    .filter((c) => c.addr.col === col)
    .sort((a, b) => a.addr.row - b.addr.row)
    .map((c) => c.numberFormat);
