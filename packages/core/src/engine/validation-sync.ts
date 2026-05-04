import type { CellFormat, SpreadsheetStore } from '../store/store.js';
import type { Range } from './types.js';
import { addrKey } from './workbook-handle.js';
import type { WorkbookHandle } from './workbook-handle.js';

/** OOXML validation `type` ordinal for `list`. Mirrors the upstream comment
 *  on `DataValidationEntry`: 0 none, 1 whole, 2 decimal, 3 list, 4 date,
 *  5 time, 6 textLength, 7 custom. */
const DV_TYPE_LIST = 3;

/**
 * Map engine validation entries to the per-cell `format.validation` field on
 * `sheet`. Only list validations (`type === 3`) are surfaced today; other
 * kinds carry richer config (range / op + formula1/2) that the UI does not
 * render yet, so they are silently dropped.
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
      if (v.type !== DV_TYPE_LIST) continue;
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

/**
 * Replace the engine's data-validation rules on `sheet` with whatever
 * FormatSlice currently asserts. Each list-validation source becomes one
 * inline-literal rule (`formula1: "A,B,C"`); cells with the same source list
 * are coalesced into a single rule with multiple ranges to keep the rule
 * count tight. No-op when `capabilities.dataValidation` is off.
 */
export function syncValidationsToEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
): void {
  if (!wb.capabilities.dataValidation) return;
  // Group cells by their source-list signature so cells sharing a list
  // collapse into a single rule with multiple single-cell ranges.
  const buckets = new Map<string, { source: string[]; ranges: Range[] }>();
  const formats = store.getState().format.formats;
  for (const [key, fmt] of formats) {
    if (!fmt.validation || fmt.validation.kind !== 'list') continue;
    const [sStr, rStr, cStr] = key.split(':');
    if (sStr === undefined || rStr === undefined || cStr === undefined) continue;
    const sIdx = Number.parseInt(sStr, 10);
    if (sIdx !== sheet) continue;
    const row = Number.parseInt(rStr, 10);
    const col = Number.parseInt(cStr, 10);
    const sig = JSON.stringify(fmt.validation.source);
    let bucket = buckets.get(sig);
    if (!bucket) {
      bucket = { source: fmt.validation.source, ranges: [] };
      buckets.set(sig, bucket);
    }
    bucket.ranges.push({ sheet, r0: row, c0: col, r1: row, c1: col });
  }
  wb.clearValidations(sheet);
  for (const { source, ranges } of buckets.values()) {
    if (ranges.length === 0) continue;
    wb.addValidationEntry(sheet, {
      type: DV_TYPE_LIST,
      ranges,
      formula1: encodeInlineList(source),
      allowBlank: true,
      showInputMessage: true,
      showErrorMessage: true,
    });
  }
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

/** Encode a UI source list back into the inline-literal form Excel expects.
 *  Always wraps in double quotes for round-trip stability. */
function encodeInlineList(source: string[]): string {
  return `"${source.join(',')}"`;
}
