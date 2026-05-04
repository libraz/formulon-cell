import { extractRefs } from '../commands/refs.js';
import type { Addr } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

/**
 * Trace-precedents / trace-dependents helpers. Same-sheet only — cross-sheet
 * refs are out of scope for v1 trace arrows. Returned addrs always carry the
 * caller's sheet index; the regex used by `extractRefs` skips quoted-string
 * literals already, so SUM("A1:B2") is not mistaken for a real ref.
 */

/** True when the ref at `[start, end)` in `text` is sheet-qualified (e.g.
 *  `Sheet1!A1` or `'Sheet 1'!A1`). `extractRefs` includes the sheet prefix
 *  in the matched span, so a `!` anywhere inside the slice signals a
 *  cross-sheet reference. v1 arrows only render same-sheet relationships. */
function isCrossSheetRef(text: string, start: number, end: number): boolean {
  return text.slice(start, end).includes('!');
}

/** Source cells referenced by the formula at `addr`. Returns an empty list
 *  when the cell is not a formula or the formula has no refs. Range refs
 *  (`A1:A5`) are expanded into individual addrs. Sheet-qualified refs are
 *  skipped — only same-sheet precedents are surfaced. */
export function findPrecedents(wb: WorkbookHandle, addr: Addr): Addr[] {
  const formula = wb.cellFormula(addr);
  if (!formula) return [];
  const refs = extractRefs(formula);
  if (refs.length === 0) return [];
  const out: Addr[] = [];
  const seen = new Set<string>();
  for (const ref of refs) {
    if (isCrossSheetRef(formula, ref.start, ref.end)) continue;
    for (let r = ref.r0; r <= ref.r1; r += 1) {
      for (let c = ref.c0; c <= ref.c1; c += 1) {
        // Skip self-reference — Excel doesn't draw an arrow from a cell to
        // itself when the formula is e.g. `=A1+A1` and A1 is the host.
        if (r === addr.row && c === addr.col) continue;
        const key = `${r}:${c}`;
        if (seen.has(key)) continue;
        seen.add(key);
        out.push({ sheet: addr.sheet, row: r, col: c });
      }
    }
  }
  return out;
}

/** Cells whose formulas read from `addr`. Iterates every populated cell on
 *  the same sheet, parses each formula's refs, and pushes the cell when any
 *  ref's rectangle contains `(addr.row, addr.col)`. */
export function findDependents(wb: WorkbookHandle, addr: Addr): Addr[] {
  const out: Addr[] = [];
  for (const cell of wb.cells(addr.sheet)) {
    if (!cell.formula) continue;
    if (cell.addr.row === addr.row && cell.addr.col === addr.col) continue;
    const refs = extractRefs(cell.formula);
    if (refs.length === 0) continue;
    for (const ref of refs) {
      if (isCrossSheetRef(cell.formula, ref.start, ref.end)) continue;
      if (addr.row >= ref.r0 && addr.row <= ref.r1 && addr.col >= ref.c0 && addr.col <= ref.c1) {
        out.push({ sheet: addr.sheet, row: cell.addr.row, col: cell.addr.col });
        break;
      }
    }
  }
  return out;
}
