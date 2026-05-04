import type { Addr, CellValue } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

/**
 * Resolve the seed text the formula bar / inline editor should display
 * for `cell`. Honours four cases in priority order:
 *
 *   1. Cells with an explicit formula text — show it verbatim.
 *   2. Cells whose cached value is a lambda (no formula stored): pull the
 *      LAMBDA body via `getLambdaText` and prefix `=` so the editor sees
 *      something parseable. Excel does the same — formula bar shows the
 *      lambda definition, not `#CALC!`.
 *   3. Plain values — render them as the editor would type them
 *      (numbers without thousand separators, booleans as TRUE/FALSE,
 *      errors as their canonical sentinel).
 *   4. Anything else (e.g. blanks) — empty string.
 *
 * `wb` and `addr` are optional so callers without engine context (e.g.
 * tests, isolated formatting paths) can still resolve cases 1, 3, 4.
 */
export function formatCellForEdit(
  cell: { value: CellValue; formula: string | null } | undefined,
  wb?: WorkbookHandle,
  addr?: Addr,
): string {
  if (!cell) return '';
  if (cell.formula) return cell.formula;
  if (wb && addr) {
    const lambda = wb.getLambdaText(addr);
    if (lambda) return `=${lambda}`;
  }
  const v = cell.value;
  switch (v.kind) {
    case 'number':
      return String(v.value);
    case 'bool':
      return v.value ? 'TRUE' : 'FALSE';
    case 'text':
      return v.value;
    case 'error':
      return v.text;
    default:
      return '';
  }
}
