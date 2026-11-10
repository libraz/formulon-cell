import {
  assertPatternSheet,
  type CellInput,
  compileRangeInput,
  patternContainsAddr,
  type RangeInput,
  type RangePattern,
} from '../engine/range-input.js';
import type { Addr, CellValue, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';

/** Operations that can be authorized by an embedded instance. Keep this list
 * finite: a new command must be classified before a restricted profile can
 * opt into it. */
export const INTERACTION_OPERATIONS = [
  'valueEdit',
  'formulaEdit',
  'clear',
  'paste',
  'fill',
  'moveCells',
  'format',
  'insertRows',
  'deleteRows',
  'insertColumns',
  'deleteColumns',
  'resizeRows',
  'resizeColumns',
  'sort',
  'merge',
  'filter',
  'table',
  'pivot',
  'comment',
  'hyperlink',
  'validation',
  'namedRange',
  'sheet',
  'object',
  'protection',
  'pageSetup',
  'print',
  'export',
] as const;

export type InteractionOperation = (typeof INTERACTION_OPERATIONS)[number];

export type InteractionOrigin =
  | 'editor'
  | 'formulaBar'
  | 'keyboard'
  | 'clipboard'
  | 'fillHandle'
  | 'contextMenu'
  | 'ribbon'
  | 'instanceApi'
  | 'undo'
  | 'redo';

/** Short alias used by hosts when declaring an editable form. */
export type Origin = InteractionOrigin;

export type PermissionCode =
  | 'readOnly'
  | 'operationDenied'
  | 'cellIneligible'
  | 'protected'
  | 'outOfBounds'
  | 'effectTooLarge'
  | 'invalid';

export type PermissionDecision =
  | { allowed: true }
  | { allowed: false; code: PermissionCode; reason?: string; addr?: Addr };

export type BatchDeniedMode = 'reject' | 'skipIneligible';

export interface EditablePredicateContext {
  readonly addr: Addr;
  readonly operation: InteractionOperation;
  readonly origin: InteractionOrigin;
}

/** A range list or a synchronous cell predicate. Ranges may be A1 text
 * (`'B2:B11'`, `'Input!B2:B11'`); text without a sheet name applies to every
 * sheet. The array/function shorthands are accepted by fixedFormPolicy. */
export type EditableCells =
  | { readonly ranges: readonly RangeInput[] }
  | { readonly predicate: (ctx: EditablePredicateContext) => boolean }
  | readonly RangeInput[]
  | ((ctx: EditablePredicateContext) => boolean);

export interface InteractionRestrictionContext {
  readonly intent: OperationIntent;
  readonly decision: PermissionDecision;
  readonly addr?: Addr;
  readonly operation: InteractionOperation;
  readonly origin: InteractionOrigin;
  readonly commandId?: string;
}

export interface InteractionPolicy {
  readonly readOnly?: boolean;
  readonly editable?: EditableCells;
  readonly operations?: Partial<Readonly<Record<InteractionOperation, boolean>>>;
  readonly defaultOperation?: 'allow' | 'deny';
  readonly selection?: boolean;
  readonly copy?: boolean;
  readonly batchDenied?: BatchDeniedMode;
  /** A final synchronous restriction. It may deny a result, but never grant
   * an operation already denied by the policy. */
  readonly restrict?: (
    context: InteractionRestrictionContext,
  ) => PermissionDecision | boolean | void;
}

export type OperationEffect =
  | {
      readonly kind: 'cells';
      readonly cells: readonly Addr[];
      readonly includesFormula?: boolean;
    }
  | { readonly kind: 'range'; readonly range: Range; readonly includesFormula?: boolean }
  /** A sheet-level write (rename, tab color, page setup, protection): decided
   *  for the sheet as a whole, never per cell. */
  | { readonly kind: 'sheet'; readonly sheet: number }
  | { readonly kind: 'workbook' };

export interface OperationIntent {
  readonly operation: InteractionOperation;
  readonly origin: InteractionOrigin;
  readonly commandId?: string;
  readonly effects: readonly OperationEffect[];
}

export type CellChangeInput =
  | { readonly addr: Addr; readonly input: string }
  | { readonly addr: Addr; readonly value: CellValue; readonly formula?: string | null };

/** A trusted host update for `applyChanges`. `addr` may be A1 text (`'B2'`,
 * `'Input!B2'`); text without a sheet name targets the active sheet. */
export type HostCellChangeInput =
  | { readonly addr: CellInput; readonly input: string }
  | { readonly addr: CellInput; readonly value: CellValue; readonly formula?: string | null };

export type CellBatchOperation =
  | 'valueEdit'
  | 'formulaEdit'
  | 'clear'
  | 'paste'
  | 'fill'
  | 'moveCells';

export interface CellBatchCommand {
  readonly type: 'cellBatch';
  readonly operation: CellBatchOperation;
  readonly origin: InteractionOrigin;
  readonly commandId?: string;
  readonly changes: readonly CellChangeInput[];
  readonly denied?: BatchDeniedMode;
}

export interface BatchRejection {
  readonly addr?: Addr;
  readonly code: PermissionCode;
  readonly reason?: string;
}

export interface ChangeBatchResult {
  readonly status: 'applied' | 'rejected' | 'noop';
  readonly applied: readonly Addr[];
  readonly rejected: readonly BatchRejection[];
  readonly revision: number;
}

const MUTATING_OPERATIONS: ReadonlySet<InteractionOperation> = new Set([
  'valueEdit',
  'formulaEdit',
  'clear',
  'paste',
  'fill',
  'moveCells',
  'format',
  'insertRows',
  'deleteRows',
  'insertColumns',
  'deleteColumns',
  'resizeRows',
  'resizeColumns',
  'sort',
  'merge',
  'filter',
  'table',
  'pivot',
  'comment',
  'hyperlink',
  'validation',
  'namedRange',
  'sheet',
  'object',
  'protection',
  'pageSetup',
]);

export const isMutatingInteraction = (operation: InteractionOperation): boolean =>
  MUTATING_OPERATIONS.has(operation);

type NormalizedEditable =
  | { readonly ranges: readonly RangeInput[] }
  | { readonly predicate: (ctx: EditablePredicateContext) => boolean };

const normalizeEditable = (editable: EditableCells): NormalizedEditable => {
  if (Array.isArray(editable)) return { ranges: editable };
  if (typeof editable === 'function') return { predicate: editable };
  const object = editable as
    | { readonly ranges: readonly RangeInput[] }
    | { readonly predicate: (ctx: EditablePredicateContext) => boolean };
  if ('ranges' in object) return { ranges: object.ranges };
  return { predicate: object.predicate };
};

const compiledEditable = new WeakMap<readonly RangeInput[], readonly RangePattern[]>();

const editablePatterns = (ranges: readonly RangeInput[]): readonly RangePattern[] => {
  let patterns = compiledEditable.get(ranges);
  if (!patterns) {
    patterns = ranges.map((range, index) =>
      compileRangeInput(range, `InteractionPolicy.editable[${index}]`),
    );
    compiledEditable.set(ranges, patterns);
  }
  return patterns;
};

/** Throw when the policy's editable ranges hold bad A1 text or name a sheet
 * the workbook does not have. */
export function validateInteractionPolicy(
  policy: InteractionPolicy | undefined,
  wb: WorkbookHandle,
): void {
  if (!policy?.editable) return;
  const editable = normalizeEditable(policy.editable);
  if (!('ranges' in editable)) return;
  editablePatterns(editable.ranges).forEach((pattern, index) => {
    assertPatternSheet(pattern, wb, `InteractionPolicy.editable[${index}]`);
  });
}

/** Return whether a cell is eligible under the host policy. Operation
 * permissions and workbook protection are deliberately handled by the
 * controller, so this helper only answers the host range/predicate question. */
export function isPolicyCellEligible(
  policy: InteractionPolicy,
  addr: Addr,
  operation: InteractionOperation,
  origin: InteractionOrigin,
  wb: WorkbookHandle,
): boolean {
  if (!policy.editable) return true;
  const editable = normalizeEditable(policy.editable);
  if ('ranges' in editable) {
    return editablePatterns(editable.ranges).some((pattern) =>
      patternContainsAddr(pattern, addr, wb),
    );
  }
  return editable.predicate({ addr, operation, origin });
}

/** Whether every cell of a range too large to enumerate is eligible under the
 *  host policy: one editable range must cover it; a predicate cannot. */
export function isPolicyRangeEligible(
  policy: InteractionPolicy,
  range: Range,
  wb: WorkbookHandle,
): boolean {
  if (!policy.editable) return true;
  const editable = normalizeEditable(policy.editable);
  if (!('ranges' in editable)) return false;
  const first: Addr = { sheet: range.sheet, row: range.r0, col: range.c0 };
  const last: Addr = { sheet: range.sheet, row: range.r1, col: range.c1 };
  return editablePatterns(editable.ranges).some(
    (pattern) => patternContainsAddr(pattern, first, wb) && patternContainsAddr(pattern, last, wb),
  );
}

/** The policy's decision for an operation before any cell effect is checked.
 *  The authorizer starts every policy decision here, and `policyPermits` is
 *  this decision's allowed flag, so the two agree by construction. */
export function operationDecision(
  policy: InteractionPolicy,
  operation: InteractionOperation,
): PermissionDecision {
  const readOnly = policy.readOnly === true && isMutatingInteraction(operation);
  if (readOnly) return { allowed: false, code: 'readOnly', reason: 'instance policy is read-only' };
  const configured = policy.operations?.[operation];
  const allowed =
    configured !== undefined ? configured : (policy.defaultOperation ?? 'deny') === 'allow';
  if (allowed) return { allowed: true };
  return { allowed: false, code: 'operationDenied', reason: `operation ${operation} is denied` };
}

export function operationPermission(
  policy: InteractionPolicy,
  operation: InteractionOperation,
): boolean {
  return operationDecision(policy, operation).allowed;
}

/** Whether a dialog or command classified as `operation` may run: always without a policy or
 *  for `null` (nothing to edit), else as the policy decides the operation. */
export const policyPermits = (
  policy: InteractionPolicy | undefined,
  operation: InteractionOperation | null,
): boolean =>
  operation === null || policy === undefined || operationDecision(policy, operation).allowed;

/** Viewer profile: users can select and copy, while every mutation remains
 * denied. Print/export are deliberately left for an explicit host choice. */
export function viewerPolicy(): InteractionPolicy {
  return {
    readOnly: true,
    defaultOperation: 'deny',
    selection: true,
    copy: true,
    batchDenied: 'reject',
  };
}

/** Fixed-form profile: the host declares the input cells and receives a
 * conservative value-entry surface. Formula, format, structure, and object
 * operations stay denied until explicitly configured by the host. */
export function fixedFormPolicy(editable: EditableCells): InteractionPolicy {
  const normalized = normalizeEditable(editable);
  if ('ranges' in normalized) editablePatterns(normalized.ranges);
  return {
    editable: normalized,
    defaultOperation: 'deny',
    operations: {
      valueEdit: true,
      clear: true,
      paste: true,
      fill: true,
    },
    selection: true,
    copy: true,
    batchDenied: 'reject',
  };
}
