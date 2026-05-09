import type { CellValue } from '../engine/types.js';

/**
 * Toggle helper for the checkbox cell type. Desktop spreadsheets ship an explicit
 * "Insert → Checkbox" cell type that:
 *   - Renders a checkbox glyph in place of TRUE/FALSE.
 *   - Flips the underlying boolean on click or when the user presses Space.
 *   - Treats blank as FALSE (an unchecked box).
 *
 * The function returns the next CellValue. Callers wire this into the
 * pointer / keyboard pipeline; the painter below handles drawing.
 */
export function toggleCheckboxValue(v: CellValue | undefined): CellValue {
  if (!v || v.kind === 'blank') return { kind: 'bool', value: true };
  if (v.kind === 'bool') return { kind: 'bool', value: !v.value };
  if (v.kind === 'number') return { kind: 'bool', value: v.value === 0 };
  // Text or error cells: best-effort – reset to checked-true so the user can
  // recover from accidentally typing a string into a checkbox cell.
  return { kind: 'bool', value: true };
}

/** Reduce a CellValue to a checkbox boolean for paint purposes. Mirrors
 *  the spreadsheet coercion: TRUE / non-zero / non-empty text → checked. */
export function isCheckboxValueChecked(v: CellValue | undefined): boolean {
  if (!v) return false;
  if (v.kind === 'blank') return false;
  if (v.kind === 'bool') return v.value;
  if (v.kind === 'number') return v.value !== 0;
  if (v.kind === 'text') return v.value.length > 0;
  return false;
}
