import type { Addr, CellValue, Range } from '../engine/types.js';
import { rangeContainsAddr } from '../store/selection-geometry.js';

/** Operations that can be authorized by an embedded instance. Keep this list
 * finite: a new command must be classified before a restricted profile can
 * opt into it. */
export type InteractionOperation =
  | 'valueEdit'
  | 'formulaEdit'
  | 'clear'
  | 'paste'
  | 'fill'
  | 'moveCells'
  | 'format'
  | 'insertRows'
  | 'deleteRows'
  | 'insertColumns'
  | 'deleteColumns'
  | 'resizeRows'
  | 'resizeColumns'
  | 'sort'
  | 'merge'
  | 'filter'
  | 'table'
  | 'comment'
  | 'hyperlink'
  | 'validation'
  | 'namedRange'
  | 'sheet'
  | 'object'
  | 'protection'
  | 'pageSetup'
  | 'print'
  | 'export';

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
  | 'unsupported'
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

/** A range list or a synchronous cell predicate. The array/function shorthands
 * are accepted by fixedFormPolicy for ergonomic form configuration. */
export type EditableCells =
  | { readonly ranges: readonly Range[] }
  | { readonly predicate: (ctx: EditablePredicateContext) => boolean }
  | readonly Range[]
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
  | { readonly ranges: readonly Range[] }
  | { readonly predicate: (ctx: EditablePredicateContext) => boolean };

const normalizeEditable = (editable: EditableCells): NormalizedEditable => {
  if (Array.isArray(editable)) return { ranges: editable };
  if (typeof editable === 'function') return { predicate: editable };
  const object = editable as
    | { readonly ranges: readonly Range[] }
    | { readonly predicate: (ctx: EditablePredicateContext) => boolean };
  if ('ranges' in object) return { ranges: object.ranges };
  return { predicate: object.predicate };
};

/** Return whether a cell is eligible under the host policy. Operation
 * permissions and workbook protection are deliberately handled by the
 * controller, so this helper only answers the host range/predicate question. */
export function isPolicyCellEligible(
  policy: InteractionPolicy,
  addr: Addr,
  operation: InteractionOperation,
  origin: InteractionOrigin,
): boolean {
  if (!policy.editable) return true;
  const editable = normalizeEditable(policy.editable);
  if ('ranges' in editable) return editable.ranges.some((range) => rangeContainsAddr(range, addr));
  return editable.predicate({ addr, operation, origin });
}

export function operationPermission(
  policy: InteractionPolicy,
  operation: InteractionOperation,
): boolean {
  if (policy.readOnly && isMutatingInteraction(operation)) return false;
  const configured = policy.operations?.[operation];
  if (configured !== undefined) return configured;
  return (policy.defaultOperation ?? 'deny') === 'allow';
}

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
  return {
    editable: normalizeEditable(editable),
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
