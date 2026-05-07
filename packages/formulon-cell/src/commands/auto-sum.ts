import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { State } from '../store/store.js';

const colLetter = (n: number): string => {
  let v = n;
  let out = '';
  do {
    out = String.fromCharCode(65 + (v % 26)) + out;
    v = Math.floor(v / 26) - 1;
  } while (v >= 0);
  return out;
};

const rangeRef = (r: Range): string =>
  `${colLetter(r.c0)}${r.r0 + 1}:${colLetter(r.c1)}${r.r1 + 1}`;

const isNum = (state: State, sheet: number, row: number, col: number): boolean => {
  if (row < 0 || col < 0) return false;
  const cell = state.data.cells.get(`${sheet}:${row}:${col}`);
  return !!cell && cell.value.kind === 'number';
};

const isEmpty = (state: State, sheet: number, row: number, col: number): boolean => {
  if (row < 0 || col < 0) return false;
  const cell = state.data.cells.get(`${sheet}:${row}:${col}`);
  return !cell || cell.value.kind === 'blank';
};

/**
 * Excel's Σ button. Decides where a SUM formula belongs based on the current
 * selection and writes it. Returns the inserted location + formula (when
 * multiple cells are written, the first one) or null when nothing reasonable
 * can be done.
 *
 * Single-cell selection:
 *   - Active cell is empty: SUM goes IN the active cell. Range = contiguous
 *     numbers above (preferred) or to the left.
 *   - Active cell is a number: SUM goes in the first empty cell at the end
 *     of its contiguous block — directly below the column block (preferred),
 *     otherwise to the right of the row block. The active cell is never
 *     overwritten.
 *
 * Range selection:
 *   - If the user included empty trailing rows or columns inside the
 *     selection, fill those with SUM formulas (per column for trailing
 *     rows, per row for trailing columns) — the SUM range is the
 *     contiguous numeric block right above/left of each target.
 *   - Otherwise, place `=SUM(<range>)` directly below the range, or to
 *     its right when below is occupied.
 */
export function autoSum(state: State, wb: WorkbookHandle): { addr: Addr; formula: string } | null {
  const r = state.selection.range;
  const sheet = state.data.sheetIndex;
  const isSingle = r.r0 === r.r1 && r.c0 === r.c1;
  const a = state.selection.active;

  if (isSingle) {
    if (isEmpty(state, sheet, a.row, a.col)) {
      // Empty active cell: scan up the column for a contiguous block.
      const r1 = a.row - 1;
      let r0 = r1;
      while (r0 - 1 >= 0 && isNum(state, sheet, r0 - 1, a.col)) r0 -= 1;
      if (r1 >= 0 && isNum(state, sheet, r1, a.col)) {
        const formula = `=SUM(${colLetter(a.col)}${r0 + 1}:${colLetter(a.col)}${r1 + 1})`;
        wb.setFormula(a, formula);
        return { addr: a, formula };
      }
      // Fall back to the row to the left.
      const c1 = a.col - 1;
      let c0 = c1;
      while (c0 - 1 >= 0 && isNum(state, sheet, a.row, c0 - 1)) c0 -= 1;
      if (c1 >= 0 && isNum(state, sheet, a.row, c1)) {
        const formula = `=SUM(${colLetter(c0)}${a.row + 1}:${colLetter(c1)}${a.row + 1})`;
        wb.setFormula(a, formula);
        return { addr: a, formula };
      }
      return null;
    }

    if (isNum(state, sheet, a.row, a.col)) {
      // Active cell is a number — find the contiguous block in its column,
      // SUM goes in the empty cell directly after the block.
      let top = a.row;
      let bottom = a.row;
      while (top - 1 >= 0 && isNum(state, sheet, top - 1, a.col)) top -= 1;
      while (isNum(state, sheet, bottom + 1, a.col)) bottom += 1;
      const colTarget: Addr = { sheet, row: bottom + 1, col: a.col };
      if (isEmpty(state, sheet, colTarget.row, colTarget.col)) {
        const formula = `=SUM(${colLetter(a.col)}${top + 1}:${colLetter(a.col)}${bottom + 1})`;
        wb.setFormula(colTarget, formula);
        return { addr: colTarget, formula };
      }
      // Fall back to the row direction.
      let left = a.col;
      let right = a.col;
      while (left - 1 >= 0 && isNum(state, sheet, a.row, left - 1)) left -= 1;
      while (isNum(state, sheet, a.row, right + 1)) right += 1;
      const rowTarget: Addr = { sheet, row: a.row, col: right + 1 };
      if (isEmpty(state, sheet, rowTarget.row, rowTarget.col)) {
        const formula = `=SUM(${colLetter(left)}${a.row + 1}:${colLetter(right)}${a.row + 1})`;
        wb.setFormula(rowTarget, formula);
        return { addr: rowTarget, formula };
      }
    }
    return null;
  }

  // Column-direction: per column in the range, place SUM right below the
  // last numeric cell in that column (when the slot is empty AND still
  // inside the selection). This handles "select numbers + trailing empty
  // rows, press Σ" — the formula lands inside the selection, not past it.
  const colWrites: { addr: Addr; formula: string }[] = [];
  for (let col = r.c0; col <= r.c1; col += 1) {
    let lastNumRow = -1;
    for (let row = r.r1; row >= r.r0; row -= 1) {
      if (isNum(state, sheet, row, col)) {
        lastNumRow = row;
        break;
      }
    }
    if (lastNumRow < 0 || lastNumRow >= r.r1) continue;
    const target: Addr = { sheet, row: lastNumRow + 1, col };
    if (!isEmpty(state, sheet, target.row, target.col)) continue;
    let top = lastNumRow;
    while (top - 1 >= r.r0 && isNum(state, sheet, top - 1, col)) top -= 1;
    const formula = `=SUM(${colLetter(col)}${top + 1}:${colLetter(col)}${lastNumRow + 1})`;
    colWrites.push({ addr: target, formula });
  }
  if (colWrites.length > 0) {
    for (const w of colWrites) wb.setFormula(w.addr, w.formula);
    return colWrites[0] ?? null;
  }

  // Row-direction: same idea horizontally — trailing empty columns inside
  // the selection get SUM formulas.
  const rowWrites: { addr: Addr; formula: string }[] = [];
  for (let row = r.r0; row <= r.r1; row += 1) {
    let lastNumCol = -1;
    for (let col = r.c1; col >= r.c0; col -= 1) {
      if (isNum(state, sheet, row, col)) {
        lastNumCol = col;
        break;
      }
    }
    if (lastNumCol < 0 || lastNumCol >= r.c1) continue;
    const target: Addr = { sheet, row, col: lastNumCol + 1 };
    if (!isEmpty(state, sheet, target.row, target.col)) continue;
    let left = lastNumCol;
    while (left - 1 >= r.c0 && isNum(state, sheet, row, left - 1)) left -= 1;
    const formula = `=SUM(${colLetter(left)}${row + 1}:${colLetter(lastNumCol)}${row + 1})`;
    rowWrites.push({ addr: target, formula });
  }
  if (rowWrites.length > 0) {
    for (const w of rowWrites) wb.setFormula(w.addr, w.formula);
    return rowWrites[0] ?? null;
  }

  const ref = rangeRef(r);
  const candidates: Addr[] = [
    { sheet, row: r.r1 + 1, col: r.c0 }, // directly below
    { sheet, row: r.r0, col: r.c1 + 1 }, // directly to the right
  ];
  for (const t of candidates) {
    if (isEmpty(state, sheet, t.row, t.col)) {
      const formula = `=SUM(${ref})`;
      wb.setFormula(t, formula);
      return { addr: t, formula };
    }
  }
  return null;
}
