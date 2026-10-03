// Contract between the Format Cells orchestrator (`format-dialog.ts`) and its
// per-tab controllers. A controller registers its tab's listeners through
// `on` (so the dialog shell removes them on detach), edits the shared draft,
// and projects the draft back onto the refs its tab builder returned.

import type { Strings } from '../../i18n/strings.js';
import type { DialogShell } from '../dialog-shell.js';
import type { DraftState } from '../format-dialog-model.js';
import type { FormatDialogField } from '../format-dialog-state.js';

export interface FormatTabContext {
  /** The dialog's draft; one object for the dialog's lifetime. */
  readonly draft: DraftState;
  readonly t: Strings['formatDialog'];
  /** Shell-tracked listener registration. */
  readonly on: DialogShell['on'];
  /** Record a user edit; clears the mixed marker on the field's controls. */
  touch(...fields: FormatDialogField[]): void;
  /** True while a field differs across the selection and is still untouched. */
  isMixed(field: FormatDialogField): boolean;
  /** Re-project the whole draft onto every tab. */
  syncControls(): void;
  renderPreview(): void;
  getLocale(): string;
}

export interface FormatTabController {
  /** Project the draft onto this tab's controls. */
  sync(): void;
  /** Blank and flag this tab's controls whose field is mixed. */
  syncMixed(): void;
}

/** Flag a control as showing a mixed value; `field` ties it to `touch`. */
export const markMixed = (
  element: HTMLElement,
  mixed: boolean,
  field?: FormatDialogField,
): void => {
  if (field !== undefined) element.dataset.fcMixedField = field;
  if (mixed) element.dataset.fcMixed = 'true';
  else delete element.dataset.fcMixed;
};

/** Checkbox form of `markMixed`: mixed shows as indeterminate. */
export const markMixedCheck = (
  input: HTMLInputElement,
  mixed: boolean,
  field: FormatDialogField,
): void => {
  input.indeterminate = mixed;
  markMixed(input, input.indeterminate, field);
};
