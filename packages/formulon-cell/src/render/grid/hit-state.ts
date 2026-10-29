import { cellValueViolatesValidation } from '../../commands/validate.js';
import type { RangeResolver } from '../../engine/range-resolver.js';
import type { CellValue } from '../../engine/types.js';
import type { CellFormat, CellValidation } from '../../store/store.js';
import type { Rect } from '../geometry.js';
import type { OutlineToggleHit } from '../painters.js';
import type { PageBandHit, PageBreakHandle, RulerHandle } from './page-view.js';

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

/** Bounds + cell of the active cell's validation chevron. */
export interface ValidationChevron {
  rect: Rect;
  row: number;
  col: number;
}

/** Bounds + cell of the totals-row function picker painted on the active
 *  cell when it sits in a table's total row. */
export interface TotalsChevron {
  rect: Rect;
  row: number;
  col: number;
  tableId: string;
  column: number;
}

/** Hit targets the latest frame painted on one grid surface. Each mount owns
 *  its own, so two spreadsheets on a page never hit-test each other's frame. */
export interface GridHitState {
  /** Fill handle bounds; null while the handle is offscreen. */
  fillHandle: Rect | null;
  /** The active cell's list-validation chevron, or null. */
  validationChevron: ValidationChevron | null;
  /** The active total-row cell's function picker, or null. */
  totalsChevron: TotalsChevron | null;
  outlineToggles: readonly OutlineToggleHit[];
  errorTriangles: readonly ErrorTriangleHit[];
  pageBands: readonly PageBandHit[];
  rulerHandles: readonly RulerHandle[];
  pageBreakHandles: readonly PageBreakHandle[];
}

const hitStates = new WeakMap<HTMLElement, GridHitState>();

/** Hit targets painted on `surface`, the grid element a `GridRenderer` paints
 *  into and the pointer layers listen on. Empty until the first paint. */
export function gridHitState(surface: HTMLElement): GridHitState {
  let state = hitStates.get(surface);
  if (!state) {
    state = {
      fillHandle: null,
      validationChevron: null,
      totalsChevron: null,
      outlineToggles: [],
      errorTriangles: [],
      pageBands: [],
      rulerHandles: [],
      pageBreakHandles: [],
    };
    hitStates.set(surface, state);
  }
  return state;
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
  return cellValueViolatesValidation(value, validation, resolveRange);
}

export function shouldShowValidationChevron(validation: CellValidation | undefined): boolean {
  return validation?.kind === 'list' && validation.showDropdown !== false;
}
