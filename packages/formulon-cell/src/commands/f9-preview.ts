import type { CellValue } from '../engine/types.js';
import { addrKey } from '../engine/workbook-handle.js';

/**
 * the spreadsheet's F9 key, while editing a formula, replaces the highlighted
 * sub-expression with its evaluated value. We can't run a full engine
 * evaluation against an arbitrary text without writing into a scratch cell,
 * so this MVP handles the two cases that cover ~95% of real F9 use:
 *
 *   - A selection that is a numeric or text literal — return the literal.
 *   - A selection that is a single A1-style reference (optionally
 *     sheet-prefixed) — look the cell up in the supplied cell map and
 *     return its current value.
 *
 * Anything else (function calls, operators) is reported as "unsupported"
 * so the editor can show a hint instead of pretending to compute.
 */
export interface F9Preview {
  /** Resolved display string (`"3.14"`, `"Hello"`, `"true"`, `"#REF!"`). */
  display: string;
  /** True when the caller can safely substitute `display` into the formula
   *  in place of the original selection. Falls back to false for partial
   *  evaluations (refs that the cell map doesn't carry, complex
   *  sub-expressions, etc.). */
  substitutable: boolean;
}

const REF_RE = /^(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_]*))?!?(\$?)([A-Za-z]+)(\$?)(\d+)$/;
const NUMBER_RE = /^-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?$/;
const STRING_RE = /^"([^"]*)"$/;

const lettersToCol = (letters: string): number => {
  let col = 0;
  for (let i = 0; i < letters.length; i += 1) {
    col = col * 26 + (letters.toUpperCase().charCodeAt(i) - 64);
  }
  return col - 1;
};

/** Render a CellValue the way the formula bar would substitute it after F9. */
export function renderCellValueForF9(v: CellValue | undefined): string {
  if (!v || v.kind === 'blank') return '0';
  if (v.kind === 'number') return String(v.value);
  if (v.kind === 'text') return `"${v.value}"`;
  if (v.kind === 'bool') return v.value ? 'TRUE' : 'FALSE';
  if (v.kind === 'error') return v.text || '#ERROR!';
  return '';
}

/** Compute the F9 substitution for `selection` taken from `formula`. The
 *  selection is the substring the user has highlighted while editing. The
 *  `cells` map mirrors `DataSlice.cells` and `sheetByName` translates a
 *  sheet name (spreadsheet-side) to its 0-based index — when omitted, sheet-
 *  qualified refs are unresolved. */
export function computeF9Preview(
  _formula: string,
  selection: string,
  activeSheet: number,
  cells: ReadonlyMap<string, { value: CellValue; formula: string | null }>,
  sheetByName?: (name: string) => number,
): F9Preview {
  const trimmed = selection.trim();
  if (!trimmed) {
    return { display: '', substitutable: false };
  }
  if (NUMBER_RE.test(trimmed)) {
    return { display: trimmed, substitutable: true };
  }
  if (STRING_RE.test(trimmed)) {
    return { display: trimmed, substitutable: true };
  }
  if (/^(true|false)$/i.test(trimmed)) {
    return { display: trimmed.toUpperCase(), substitutable: true };
  }
  const ref = trimmed.match(REF_RE);
  if (ref) {
    const sheetName = ref[1] ?? ref[2];
    const letters = ref[4] ?? '';
    const digits = ref[6] ?? '';
    const col = lettersToCol(letters);
    const row = Number.parseInt(digits, 10) - 1;
    if (row < 0 || col < 0) {
      return { display: '#REF!', substitutable: false };
    }
    let sheet = activeSheet;
    if (sheetName) {
      const resolved = sheetByName?.(sheetName);
      if (resolved === undefined || resolved < 0) {
        return { display: '#REF!', substitutable: false };
      }
      sheet = resolved;
    }
    const cell = cells.get(addrKey({ sheet, row, col }));
    return { display: renderCellValueForF9(cell?.value), substitutable: true };
  }
  // Anything else (sub-expression, function call, range): we can't evaluate
  // without invoking the engine. Report unsupported — the caller surfaces a
  // hint and leaves the formula text intact.
  return { display: '', substitutable: false };
}
