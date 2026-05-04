import type { CellValue, Range } from './types.js';
import { addrKey } from './workbook-handle.js';

/** Functions whose result is expected to spill. Anchor-cell formulas starting
 *  with one of these (case-insensitive) are candidates for spill detection.
 *  Bare ranges (`=A1:A10`, `=Sheet1!A1:A10`) are also treated as candidates
 *  via the regex below. */
const ARRAY_FUNCS = new Set([
  'FILTER',
  'SORT',
  'SORTBY',
  'UNIQUE',
  'SEQUENCE',
  'RANDARRAY',
  'TRANSPOSE',
  'MUNIT',
  'TEXTSPLIT',
  'WRAPROWS',
  'WRAPCOLS',
  'TOROW',
  'TOCOL',
  'CHOOSEROWS',
  'CHOOSECOLS',
  'TAKE',
  'DROP',
  'EXPAND',
  'HSTACK',
  'VSTACK',
]);

/** A formula text that is just a range reference, with optional sheet prefix.
 *  Matches `=A1:A10`, `=$A$1:$B$5`, `=Sheet1!A1:A10`, `='My Sheet'!A1:A10`. */
const BARE_RANGE_RE =
  /^=\s*(?:'[^']+'|[A-Za-z_][A-Za-z0-9_]*)?!?\$?[A-Za-z]+\$?\d+:\$?[A-Za-z]+\$?\d+\s*$/;

const FUNC_HEAD_RE = /^=\s*([A-Za-z_][A-Za-z0-9_]*)\s*\(/;

/** True when `formula` is heuristically expected to return an array. The
 *  detector errs on the side of producing an outline; a false positive only
 *  paints a thin ring around a single cell. */
export function looksLikeArrayFormula(formula: string): boolean {
  if (!formula.startsWith('=')) return false;
  if (BARE_RANGE_RE.test(formula)) return true;
  const m = FUNC_HEAD_RE.exec(formula);
  if (!m) return false;
  return ARRAY_FUNCS.has((m[1] ?? '').toUpperCase());
}

/** Walk right and down from an anchor cell counting consecutive cells with a
 *  populated value but no formula. The anchor itself is included in the
 *  returned range. Stops at the first cell that is blank, has its own formula,
 *  or doesn't exist. */
export function detectSpillRange(
  cells: ReadonlyMap<string, { value: CellValue; formula: string | null }>,
  sheet: number,
  row: number,
  col: number,
): Range {
  let lastCol = col;
  for (let c = col + 1; c < col + 16384; c += 1) {
    const cell = cells.get(addrKey({ sheet, row, col: c }));
    if (!cell || cell.formula !== null || cell.value.kind === 'blank') break;
    lastCol = c;
  }
  let lastRow = row;
  for (let r = row + 1; r < row + 1_048_576; r += 1) {
    const cell = cells.get(addrKey({ sheet, row: r, col }));
    if (!cell || cell.formula !== null || cell.value.kind === 'blank') break;
    lastRow = r;
  }
  return { sheet, r0: row, c0: col, r1: lastRow, c1: lastCol };
}

/** Scan every populated cell on `sheet` and return the spill rects whose
 *  anchor formula looks like a dynamic-array call. Single-cell results
 *  (1×1) are filtered out — there's nothing to outline. */
export function findSpillRanges(
  cells: ReadonlyMap<string, { value: CellValue; formula: string | null }>,
  sheet: number,
): Range[] {
  const out: Range[] = [];
  for (const [key, cell] of cells) {
    if (!cell.formula || !looksLikeArrayFormula(cell.formula)) continue;
    const [sStr, rStr, cStr] = key.split(':');
    if (sStr === undefined || rStr === undefined || cStr === undefined) continue;
    if (Number.parseInt(sStr, 10) !== sheet) continue;
    const row = Number.parseInt(rStr, 10);
    const col = Number.parseInt(cStr, 10);
    const r = detectSpillRange(cells, sheet, row, col);
    if (r.r0 === r.r1 && r.c0 === r.c1) continue;
    out.push(r);
  }
  return out;
}
