import type { CellAddress } from '@libraz/formulon';
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
import { deny, type PermissionReasonKey } from './permission-reasons.js';

/** Operations that can be authorized by an embedded instance. Keep this list
 * finite: a new command must be classified before a restricted profile can
 * opt into it. */
export const INTERACTION_OPERATIONS = [
  'valueEdit',
  'formulaEdit',
  'clear',
  'paste',
  'fill',
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
  | {
      allowed: false;
      code: PermissionCode;
      /** English text; host-supplied when the refusal has no `reasonKey`. */
      reason?: string;
      /** Set on built-in refusals; UI surfaces localize through it. */
      reasonKey?: PermissionReasonKey;
      addr?: Addr;
    };

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

/**
 * Limits on what an end user can do through the spreadsheet UI: editing, structural and
 * object operations, selection and copy.
 *
 * InteractionPolicy and sheet protection govern end-user interaction; the Workbook/Sheet API is
 * host code and is not gated by them.
 */
export interface InteractionPolicy {
  readonly readOnly?: boolean;
  readonly editable?: EditableCells;
  readonly operations?: Partial<Readonly<Record<InteractionOperation, boolean>>>;
  readonly defaultOperation?: 'allow' | 'deny';
  readonly selection?: boolean;
  readonly copy?: boolean;
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

export type CellBatchOperation = 'valueEdit' | 'formulaEdit' | 'clear' | 'paste' | 'fill';

export interface CellBatchCommand {
  readonly type: 'cellBatch';
  readonly operation: CellBatchOperation;
  readonly origin: InteractionOrigin;
  readonly commandId?: string;
  readonly changes: readonly CellChangeInput[];
}

/** The refusal that decided a batch: the cell it fell on (`null` when it concerns no single
 *  cell), its code and why. */
export interface BatchRejection {
  readonly sheet: number | null;
  readonly at: CellAddress | null;
  readonly code: PermissionCode;
  readonly reason: string;
  /** Set when the refusal is built in; see `PermissionDecision`. */
  readonly reasonKey?: PermissionReasonKey;
}

/** A cell a batch wrote. */
export interface BatchCell {
  readonly sheet: number;
  readonly at: CellAddress;
}

/**
 * The outcome of every write of many cells, all or nothing: `applied`, or `refused` with nothing
 * written and the one refusal that decided it in `rejections`.
 */
export interface ChangeBatchResult {
  readonly status: 'applied' | 'refused';
  /**
   * The cells whose content actually changed, bounded by the data. A cell written with the
   * content it already had lies inside `range` but is not listed here.
   */
  readonly applied: readonly BatchCell[];
  readonly rejections: readonly BatchRejection[];
  /**
   * The rectangle the batch wrote: the target of a paste or fill, whole rows or columns
   * included; the bounding box of the cells a sparse batch targeted. `null` when refused, when
   * nothing was targeted, or when the targets span more than one sheet.
   */
  readonly range: Range | null;
}

/** A refusal before it is reported: the cell it fell on, if any. */
export interface CellRefusal {
  readonly addr?: Addr;
  readonly code: PermissionCode;
  readonly reason?: string;
  readonly reasonKey?: PermissionReasonKey;
}

/** The rectangle around `cells`, or null when there are none or they span sheets. */
export const boundingRange = (cells: readonly Addr[]): Range | null => {
  const [first] = cells;
  if (!first) return null;
  const range = { sheet: first.sheet, r0: first.row, c0: first.col, r1: first.row, c1: first.col };
  for (const { sheet, row, col } of cells) {
    if (sheet !== range.sheet) return null;
    range.r0 = Math.min(range.r0, row);
    range.c0 = Math.min(range.c0, col);
    range.r1 = Math.max(range.r1, row);
    range.c1 = Math.max(range.c1, col);
  }
  return range;
};

/** The result of a batch over `range` that changed `cells`. */
export const appliedBatch = (cells: readonly Addr[], range: Range | null): ChangeBatchResult => ({
  status: 'applied',
  applied: cells.map(({ sheet, row, col }) => ({ sheet, at: { row, col } })),
  rejections: [],
  range,
});

/** The result of a batch `refusal` refused: nothing written. */
export const refusedBatch = (refusal: CellRefusal): ChangeBatchResult => ({
  status: 'refused',
  applied: [],
  rejections: [
    {
      sheet: refusal.addr?.sheet ?? null,
      at: refusal.addr ? { row: refusal.addr.row, col: refusal.addr.col } : null,
      code: refusal.code,
      reason: refusal.reason ?? refusal.code,
      ...(refusal.reasonKey ? { reasonKey: refusal.reasonKey } : {}),
    },
  ],
  range: null,
});

/** The result of a batch refused by `decision`. */
export const refusedByDecision = (
  decision: Extract<PermissionDecision, { allowed: false }>,
): ChangeBatchResult =>
  refusedBatch({
    addr: decision.addr,
    code: decision.code,
    reason: decision.reason,
    reasonKey: decision.reasonKey,
  });

const MUTATING_OPERATIONS: ReadonlySet<InteractionOperation> = new Set([
  'valueEdit',
  'formulaEdit',
  'clear',
  'paste',
  'fill',
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
  if (readOnly) return deny('readOnly', 'readOnlyPolicy');
  const configured = policy.operations?.[operation];
  const allowed =
    configured !== undefined ? configured : (policy.defaultOperation ?? 'deny') === 'allow';
  if (allowed) return { allowed: true };
  return deny('operationDenied', 'operationDenied', { reason: `operation ${operation} is denied` });
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

/** Whether `policy` lets a write carry formulas and formats wherever it lands: it limits no cell
 *  and allows pasting, formulas and formats. Writes that move whole cells (the clipboard, a drag
 *  move, a fill drag's menu) ask the policy only for the operation, so they need this. */
export const fullFidelityCellWrites = (policy: InteractionPolicy | undefined): boolean => {
  if (policy === undefined) return true;
  if (policy.editable !== undefined) return false;
  return (['paste', 'formulaEdit', 'format'] as const).every((operation) =>
    policyPermits(policy, operation),
  );
};

/** Whether the clipboard carries values alone under `policy` (see {@link fullFidelityCellWrites}).
 *  A values-only clipboard has no move, so it cannot cut. */
export const valuesOnlyClipboard = (policy: InteractionPolicy | undefined): boolean =>
  !fullFidelityCellWrites(policy);

/** Viewer profile: users can select and copy, while every mutation remains
 * denied. Print/export are deliberately left for an explicit host choice. */
export function viewerPolicy(): InteractionPolicy {
  return {
    readOnly: true,
    defaultOperation: 'deny',
    selection: true,
    copy: true,
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
  };
}
