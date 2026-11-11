import type { OwnPermissionReasonKey } from '../../../commands/permission-reasons.js';

/** Localized text for built-in refusal reasons. The wording is written for
 *  this package, not taken from the desktop spreadsheet's strings (its
 *  protected-sheet alert is reused through `protection.protectedCellEdit`). */
export interface PermissionReasonsStrings {
  permissionReasons: Record<OwnPermissionReasonKey, string>;
}
