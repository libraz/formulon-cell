import { MAX_COL, MAX_ROW } from '../engine/address.js';
import type { Addr, CellValue, Range } from '../engine/types.js';
import type { State } from '../store/store.js';
import { shiftFormulaRefs } from './refs.js';

type RestrictedCellChange = {
  addr: Addr;
  value: CellValue;
  formula?: string | null;
};

const cellAt = (state: State, addr: Addr): { value: CellValue; formula: string | null } => {
  const cell = state.data.cells.get(`${addr.sheet}:${addr.row}:${addr.col}`);
  return { value: cell?.value ?? { kind: 'blank' }, formula: cell?.formula ?? null };
};

/** Build a value/formula-only fill plan before touching the workbook. The
 * restricted path must authorize the complete destination as one batch; the
 * legacy fill command is intentionally left untouched for unrestricted mounts. */
export const restrictedFillChanges = (
  state: State,
  src: Range,
  dest: Range,
  copyOnly: boolean,
): RestrictedCellChange[] => {
  const srcRows = src.r1 - src.r0 + 1;
  const srcCols = src.c1 - src.c0 + 1;
  const down = dest.r1 !== src.r1 || dest.r0 !== src.r0;
  const right = !down && (dest.c1 !== src.c1 || dest.c0 !== src.c0);
  const changes: RestrictedCellChange[] = [];
  const sourceAt = (
    row: number,
    col: number,
  ): { addr: Addr; value: CellValue; formula: string | null } => {
    const addr = { sheet: src.sheet, row, col };
    const cell = cellAt(state, addr);
    return { addr, ...cell };
  };
  const projectedNumber = (
    target: Addr,
    source: { addr: Addr; value: CellValue },
  ): CellValue | null => {
    if (copyOnly || source.value.kind !== 'number') return null;
    if (down && srcRows >= 2) {
      const first = sourceAt(src.r0, target.col).value;
      const second = sourceAt(src.r0 + 1, target.col).value;
      if (first.kind === 'number' && second.kind === 'number') {
        const step = second.value - first.value;
        const offset = target.row - src.r1;
        const last = sourceAt(src.r1, target.col).value;
        return last.kind === 'number'
          ? { kind: 'number', value: last.value + step * offset }
          : null;
      }
    }
    if (right && srcCols >= 2) {
      const first = sourceAt(target.row, src.c0).value;
      const second = sourceAt(target.row, src.c0 + 1).value;
      if (first.kind === 'number' && second.kind === 'number') {
        const step = second.value - first.value;
        const offset = target.col - src.c1;
        const last = sourceAt(target.row, src.c1).value;
        return last.kind === 'number'
          ? { kind: 'number', value: last.value + step * offset }
          : null;
      }
    }
    return null;
  };
  for (let row = dest.r0; row <= dest.r1; row += 1) {
    for (let col = dest.c0; col <= dest.c1; col += 1) {
      if (row >= src.r0 && row <= src.r1 && col >= src.c0 && col <= src.c1) continue;
      const sr = src.r0 + ((((row - src.r0) % srcRows) + srcRows) % srcRows);
      const sc = src.c0 + ((((col - src.c0) % srcCols) + srcCols) % srcCols);
      const source = sourceAt(sr, sc);
      const addr = { sheet: dest.sheet, row, col };
      const projected = projectedNumber(addr, source);
      if (projected) {
        changes.push({ addr, value: projected, formula: null });
      } else if (source.formula) {
        changes.push({
          addr,
          value: source.value,
          formula: shiftFormulaRefs(source.formula, row - source.addr.row, col - source.addr.col),
        });
      } else {
        changes.push({ addr, value: source.value, formula: null });
      }
    }
  }
  return changes;
};

/**
 * Desktop spreadsheets "double-click the fill handle" rule: extend the source range downward
 * by the contiguous-data run of the immediate left-then-right neighbour
 * column. Returns null when neither neighbour has a usable run (so we don't
 * flash a no-op fill).
 *
 * The neighbour run starts at the row just below the source bottom edge
 * (src.r1 + 1) and ends at the last non-blank row in that column. We require
 * at least one non-blank cell at row src.r1 + 1 — without it, spreadsheets don't
 * expand either, so a stray cell ten rows down won't trigger an unexpected
 * fill.
 */
export function autoFillDownExtent(state: State, src: Range): Range | null {
  const sheet = src.sheet;
  const start = src.r1 + 1;
  if (start > MAX_ROW) return null;
  const probeCols: number[] = [];
  if (src.c0 > 0) probeCols.push(src.c0 - 1); // left first (desktop spreadsheets preference)
  if (src.c1 < MAX_COL) probeCols.push(src.c1 + 1);

  let endRow = -1;
  for (const col of probeCols) {
    if (!hasCellAt(state, sheet, start, col)) continue;
    let r = start;
    while (r <= MAX_ROW && hasCellAt(state, sheet, r, col)) r += 1;
    const last = r - 1;
    if (last > endRow) endRow = last;
  }
  if (endRow < start) return null;
  return { sheet, r0: src.r0, c0: src.c0, r1: endRow, c1: src.c1 };
}

function hasCellAt(state: State, sheet: number, row: number, col: number): boolean {
  const cell = state.data.cells.get(`${sheet}:${row}:${col}`);
  if (!cell) return false;
  if (cell.formula) return true;
  return cell.value.kind !== 'blank';
}
