import type { CellFormat, SpreadsheetStore } from '../store/store.js';
import { addrKey } from './workbook-handle.js';
import type { WorkbookHandle } from './workbook-handle.js';

/**
 * Map engine validation entries to the per-cell `format.validation` field on
 * `sheet`. Only `type === 'list'` is supported by the UI today; other kinds
 * are silently dropped (their writeable round-trip will land alongside an
 * upstream extension).
 *
 * The engine's `formula1` for list validations is one of:
 *   - an inline literal: `"Yes,No,Maybe"` or `Yes,No,Maybe`
 *   - a range reference: `Sheet1!$A$1:$A$10` (not yet expanded — dropped)
 *
 * Inline literals are split on comma. Surrounding double quotes (Excel's
 * standard wrapping) are stripped.
 */
export function hydrateValidationsFromEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
): void {
  if (!wb.capabilities.dataValidation) return;
  const entries = wb.getValidationsForSheet(sheet);
  if (entries.length === 0) return;

  store.setState((s) => {
    const formats = new Map(s.format.formats);
    for (const v of entries) {
      if (v.type !== 'list') continue;
      const list = parseInlineList(v.formula1);
      if (!list) continue;
      for (const r of v.ranges) {
        if (r.sheet !== sheet) continue;
        for (let row = r.r0; row <= r.r1; row += 1) {
          for (let col = r.c0; col <= r.c1; col += 1) {
            const k = addrKey({ sheet, row, col });
            const cur = formats.get(k);
            const next: CellFormat = {
              ...(cur ?? {}),
              validation: { kind: 'list', source: list },
            };
            formats.set(k, next);
          }
        }
      }
    }
    return { ...s, format: { formats } };
  });
}

/** Parse Excel-style inline list literals. Returns null when the source
 *  string is empty, a range reference, or otherwise unparseable. */
function parseInlineList(formula: string): string[] | null {
  const trimmed = formula.trim().replace(/^=/, '');
  if (!trimmed) return null;
  // Range references contain `!` (sheet-qualified) or `$` (anchored) or
  // `:` (range). Drop those since we'd need formula evaluation to expand.
  if (/[!$:]/.test(trimmed)) return null;
  // Strip outer double quotes if present, then split on comma.
  const inner = trimmed.startsWith('"') && trimmed.endsWith('"') ? trimmed.slice(1, -1) : trimmed;
  const parts = inner
    .split(',')
    .map((s) => s.trim())
    .filter((s) => s.length > 0);
  return parts.length > 0 ? parts : null;
}
