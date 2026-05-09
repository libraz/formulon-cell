import { extractRefs } from '../commands/refs.js';
import type { Addr } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

/**
 * Trace-precedents / trace-dependents helpers. When the engine exposes
 * `precedents` / `dependents` (5/5 build onward), these wrappers delegate
 * to the engine — that surfaces cross-sheet refs and reflects the live
 * dep graph. Otherwise (stub engine, pre-5/5 vendored builds) we fall
 * back to a same-sheet regex scan over `extractRefs`. The regex skips
 * quoted-string literals so `SUM("A1:B2")` is not mistaken for a real ref.
 */

/** True when the ref at `[start, end)` in `text` is sheet-qualified (e.g.
 *  `Sheet1!A1` or `'Sheet 1'!A1`). `extractRefs` includes the sheet prefix
 *  in the matched span, so a `!` anywhere inside the slice signals a
 *  cross-sheet reference. v1 arrows only render same-sheet relationships. */
function isCrossSheetRef(text: string, start: number, end: number): boolean {
  return text.slice(start, end).includes('!');
}

/** Source cells referenced by the formula at `addr`. Engine-backed when
 *  available — returns the full graph including cross-sheet refs. Falls
 *  back to a same-sheet regex scan when `capabilities.traceArrows` is off,
 *  in which case range refs (`A1:A5`) are expanded into individual addrs
 *  and sheet-qualified refs are skipped. */
export function findPrecedents(wb: WorkbookHandle, addr: Addr): Addr[] {
  const fromEngine = wb.precedents(addr);
  if (fromEngine !== null) return fromEngine;
  return scanPrecedents(wb, addr);
}

/** Cells whose formulas read from `addr`. Engine-backed when available;
 *  same-sheet-only fallback otherwise. */
export function findDependents(wb: WorkbookHandle, addr: Addr): Addr[] {
  const fromEngine = wb.dependents(addr);
  if (fromEngine !== null) return fromEngine;
  return scanDependents(wb, addr);
}

/** Same-sheet regex fallback for `findPrecedents`. Exported for tests; the
 *  engine path is preferred at runtime when available. */
export function scanPrecedents(wb: WorkbookHandle, addr: Addr): Addr[] {
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
        // Skip self-reference — spreadsheets don't draw an arrow from a cell to
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

/** Same-sheet regex fallback for `findDependents`. Exported for tests. */
export function scanDependents(wb: WorkbookHandle, addr: Addr): Addr[] {
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
