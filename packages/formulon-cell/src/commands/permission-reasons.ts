// Built-in refusal reasons. The public `reason` of a decision stays English
// with a stable `code`; every built-in refusal also carries a `reasonKey` so UI
// surfaces can show localized text through `localizedDecisionReason`. A reason
// without a built-in key is host-supplied text and is shown unchanged.

import type { Strings } from '../i18n/strings/types.js';
import type { PermissionCode, PermissionDecision } from './interaction-policy.js';

/** English text of every built-in refusal, keyed by its stable reason key. */
export const PERMISSION_REASONS = {
  sheetProtected: 'sheet is protected',
  workbookStructureProtected: 'workbook structure is protected',
  cellProtected: 'cell is protected',
  cellHoldsNote: 'the cell holds a note',
  allowEditRangesUnavailable: 'Allow Edit Ranges is not on the Mac desktop',
  selectionOverlapsTable: 'selection overlaps a table',
  dialogNoChoiceAllowed: 'no choice of the dialog is allowed',
  valuesOnlyClipboardCannotCut: 'a values-only clipboard cannot cut',
  notBuiltInCommand: 'not a built-in command',
  selectionDisabled: 'selection is disabled',
  copyDisabled: 'copy is disabled',
  controllerDisposed: 'controller is disposed',
  workbookUnavailable: 'workbook is unavailable',
  hyperlinkNeedsOneCell: 'Automatic hyperlink authorization needs one cell and no formula effect',
  hyperlinkNeedsPermission:
    'Automatic hyperlink needs protected-sheet format and hyperlink permission',
  mergedRangeTooLarge: 'merged cell range exceeds the authorization bound',
  cellOutsideBounds: 'cell is outside the configured bounds',
  cellOutsideWorkbook: 'cell address is outside the workbook',
  cellNotEditable: 'cell is outside editable cells',
  sheetHasNoCells: 'the sheet has no cells',
  rangeOutsideBounds: 'range is outside the configured bounds',
  rangeOutsideWorkbook: 'range is outside the workbook',
  rangeNotEditable: 'range is outside editable cells',
  sheetOutsideBounds: 'sheet is outside the configured bounds',
  sheetNotInWorkbook: 'sheet is not in the workbook',
  restrictionHookAsync: 'restriction hook must be synchronous',
  restrictionCallbackFailed: 'restriction callback failed',
  restrictedByHost: 'restricted by host',
  readOnlyPolicy: 'instance policy is read-only',
  operationDenied: 'operation is denied',
  valueNotWritten: 'the value could not be written',
  cellsNotFilled: 'the selected cells could not be filled',
  affectedRangeMalformed: 'affected range is malformed',
  affectedRangeTooLarge: 'affected range exceeds the authorization bound',
  hostCopyDisabled: 'Copy is disabled by the host policy.',
  boundedViewportStructural: 'This structural operation is unavailable with a bounded viewport.',
  restrictedEmbedding: 'Unavailable in restricted embedding',
  codeReadOnly: 'read-only',
  codeOperationDenied: 'operation is denied',
  codeCellIneligible: 'cell is outside editable cells',
  codeProtected: 'protected',
  codeOutOfBounds: 'outside the allowed bounds',
  codeEffectTooLarge: 'effect is too large',
  codeInvalid: 'invalid',
} as const;

export type PermissionReasonKey = keyof typeof PERMISSION_REASONS;

/** Every built-in reason key. */
export const PERMISSION_REASON_KEYS = Object.keys(PERMISSION_REASONS) as PermissionReasonKey[];

/** Reason keys whose localized text is an existing dictionary entry. */
export type ReferencedPermissionReasonKey =
  | 'sheetProtected'
  | 'cellProtected'
  | 'workbookStructureProtected';

/** Keys the `permissionReasons` dictionary section defines itself. */
export type OwnPermissionReasonKey = Exclude<PermissionReasonKey, ReferencedPermissionReasonKey>;

const CODE_KEYS: Readonly<Record<PermissionCode, PermissionReasonKey>> = {
  readOnly: 'codeReadOnly',
  operationDenied: 'codeOperationDenied',
  cellIneligible: 'codeCellIneligible',
  protected: 'codeProtected',
  outOfBounds: 'codeOutOfBounds',
  effectTooLarge: 'codeEffectTooLarge',
  invalid: 'codeInvalid',
};

export type RefusedDecision = Extract<PermissionDecision, { allowed: false }>;

/** A built-in refusal: English `reason` (or `reason` override carrying an
 *  interpolated detail) plus its `reasonKey`. */
export const deny = (
  code: PermissionCode,
  reasonKey: PermissionReasonKey,
  extra: { readonly addr?: RefusedDecision['addr']; readonly reason?: string } = {},
): RefusedDecision => ({
  allowed: false,
  code,
  reason: extra.reason ?? PERMISSION_REASONS[reasonKey],
  reasonKey,
  ...(extra.addr ? { addr: extra.addr } : {}),
});

/** Localized text of one built-in reason key. */
export const permissionReasonText = (strings: Strings, key: PermissionReasonKey): string => {
  switch (key) {
    case 'sheetProtected':
    case 'cellProtected':
      return strings.protection.protectedCellEdit;
    case 'workbookStructureProtected':
      return strings.ribbonMenu.workbookStructureProtectedBlocked;
    default:
      return strings.permissionReasons[key];
  }
};

const isReasonKey = (key: unknown): key is PermissionReasonKey =>
  typeof key === 'string' && Object.hasOwn(PERMISSION_REASONS, key);

/** The refusal fields a UI projection reads; decisions and batch rejections fit. */
export interface RefusalLike {
  readonly code: PermissionCode;
  readonly reason?: string;
  readonly reasonKey?: PermissionReasonKey;
}

/** The text a UI surface shows for a refusal: the localized built-in reason,
 *  the host's own text unchanged, or the localized generic text of its code. */
export const localizedDecisionReason = (refusal: RefusalLike, strings: Strings): string => {
  if (isReasonKey(refusal.reasonKey)) return permissionReasonText(strings, refusal.reasonKey);
  if (refusal.reason) return refusal.reason;
  return permissionReasonText(strings, CODE_KEYS[refusal.code] ?? 'codeInvalid');
};
