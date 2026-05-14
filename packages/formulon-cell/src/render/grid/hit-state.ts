import { coerceInput } from '../../commands/coerce-input.js';
import { validateAgainst } from '../../commands/validate.js';
import type { RangeResolver } from '../../engine/range-resolver.js';
import type { CellValue } from '../../engine/types.js';
import type { CellFormat, CellValidation } from '../../store/store.js';
import type { Rect } from '../geometry.js';
import type { OutlineToggleHit } from '../painters.js';

export function isPlainTextOverflowCandidate(input: {
  value: CellValue;
  formula: string | null;
  format?: CellFormat;
  showFormulas: boolean;
  displayOverride: string | null;
  tableHeader: boolean;
  hasIcon: boolean;
  isMergeAnchor: boolean;
}): boolean {
  if (input.tableHeader || input.hasIcon || input.isMergeAnchor) return false;
  if (input.format?.wrap || input.format?.rotation || input.format?.align) return false;
  if (input.showFormulas && input.formula) return false;
  if (input.displayOverride != null) return input.value.kind === 'text';
  return input.value.kind === 'text' && !input.formula;
}

export type ErrorTriangleKind = 'error' | 'validation';

/** Hot-zone of an error/validation triangle painted in this frame. The
 *  click layer (mount.ts) hit-tests these to open the popover menu. */
export interface ErrorTriangleHit {
  rect: Rect;
  addr: { sheet: number; row: number; col: number };
  kind: ErrorTriangleKind;
}

/** Color used for formula-error triangles (spreadsheet-style green). */
export const ERROR_TRIANGLE_COLOR = '#2ea043';
/** Color used for data-validation violation triangles (spreadsheet-style red). */
export const VALIDATION_TRIANGLE_COLOR = '#d24545';

let cachedFillHandleRect: Rect | null = null;
let cachedValidationChevron: { rect: Rect; row: number; col: number } | null = null;
let cachedOutlineToggles: OutlineToggleHit[] = [];
let cachedErrorTriangles: ErrorTriangleHit[] = [];

/** Latest hit-rects of all error / validation triangles painted in this
 *  frame. The mount-level click handler hit-tests these to open the
 *  error-info popover. */
export function getErrorTriangleHits(): ErrorTriangleHit[] {
  return cachedErrorTriangles;
}

export function setErrorTriangleHits(hits: ErrorTriangleHit[]): void {
  cachedErrorTriangles = hits;
}

/** Spreadsheet error sentinels that the renderer surfaces with a green corner
 *  triangle. Engine-typed errors take the `value.kind === 'error'` branch;
 *  the string set covers cases where a custom formatter / passthrough
 *  layer left a string sentinel in the cell. */
const ERROR_SENTINELS: ReadonlySet<string> = new Set([
  '#DIV/0!',
  '#NAME?',
  '#REF!',
  '#VALUE!',
  '#NUM!',
  '#N/A',
  '#NULL!',
  '#CIRCULAR!',
]);

export { normalizeFormatLocale } from '../../format/locale.js';

/** Detect whether `cell.value` should surface an error indicator. */
export function detectErrorKind(value: CellValue): boolean {
  if (value.kind === 'error') return true;
  if (value.kind === 'text') return ERROR_SENTINELS.has(value.value);
  return false;
}

/** Detect whether `cell.value` violates `validation`. Returns false when
 *  validation is missing, when the cell is blank and `allowBlank` is set,
 *  or when the value is itself an error (we surface that as an error
 *  triangle, not a validation triangle). */
export function detectValidationViolation(
  value: CellValue,
  validation: CellValidation | undefined,
  resolveRange?: RangeResolver,
): boolean {
  if (!validation) return false;
  if (value.kind === 'error') return false;
  // Re-use coerceInput by stringifying the value first — same shape the
  // keyboard / formula-bar paths feed validateAgainst.
  let raw: string;
  switch (value.kind) {
    case 'blank':
      raw = '';
      break;
    case 'number':
      raw = String(value.value);
      break;
    case 'bool':
      raw = value.value ? 'TRUE' : 'FALSE';
      break;
    case 'text':
      raw = value.value;
      break;
  }
  const coerced = coerceInput(raw);
  const outcome = validateAgainst(validation, coerced, resolveRange);
  return !outcome.ok;
}

/** Latest hit-rects of all outline +/- toggles painted in this frame. The
 *  pointer layer hit-tests these to route clicks to collapse/expand. */
export function getOutlineToggleHits(): OutlineToggleHit[] {
  return cachedOutlineToggles;
}

export function setOutlineToggles(hits: OutlineToggleHit[]): void {
  cachedOutlineToggles = hits;
}

/** Latest device-space bounds of the fill handle. Hit-tested by the pointer
 *  layer to start a fill drag. Null while the handle is offscreen. */
export function getFillHandleRect(): Rect | null {
  return cachedFillHandleRect;
}

export function setFillHandleRect(r: Rect | null): void {
  cachedFillHandleRect = r;
}

/** Bounds + addr of the active cell's validation chevron, or null when the
 *  active cell has no list validation. */
export function getValidationChevron(): { rect: Rect; row: number; col: number } | null {
  return cachedValidationChevron;
}

export function setValidationChevron(v: { rect: Rect; row: number; col: number } | null): void {
  cachedValidationChevron = v;
}
