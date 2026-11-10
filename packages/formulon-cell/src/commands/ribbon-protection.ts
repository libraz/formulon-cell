// Shared "Allow Users to Edit Ranges" ribbon split-button — adds a new allowed
// range from the active selection, or wipes every allowed range on the sheet.
// The mutation is one call into protection.ts; the dialog payload is what the
// hosts actually care about, so we hand it back as a [[RibbonProtectionReport]]
// they can drop straight onto their ribbon-report shell.

import { formatA1Range } from '../engine/address.js';
import type { Strings } from '../i18n/strings.js';
import { isSheetProtected } from './protection.js';
import {
  addAllowedEditRangeAction,
  clearAllowedEditRangesAction,
  type ProtectionActionContext,
} from './protection-actions.js';

type ProtectionStrings = Pick<
  Strings['ribbonMenu'],
  | 'allowEditRangesDialogTitle'
  | 'allowEditRangesCommand'
  | 'allowEditRangesClearCommand'
  | 'allowedEditRangeAddedStatus'
  | 'allowedEditRangesClearedStatus'
  | 'allowEditRangesRequiresUnprotected'
>;

export type RibbonProtectionAction = 'allow-edit-range' | 'clear-allowed-edit-ranges';

export interface RibbonProtectionReport {
  title: string;
  items: { severity: 'info' | 'warning'; label: string; detail: string }[];
}

export interface ExecuteRibbonProtectionActionDeps {
  /** The store, workbook and entry point the range change is authorized and written under. */
  context: ProtectionActionContext;
  action: RibbonProtectionAction;
  strings: ProtectionStrings;
}

/** Mutate the store and return the dialog payload the host should surface.
 *  `allow-edit-range` adds the current selection as a titled range; everything
 *  else clears every range on the active sheet. */
export const executeRibbonProtectionAction = (
  deps: ExecuteRibbonProtectionActionDeps,
): RibbonProtectionReport => {
  const { context, action, strings } = deps;
  const { store } = context;
  const state = store.getState();
  const range = state.selection.range;
  const sheet = action === 'allow-edit-range' ? range.sheet : state.data.sheetIndex;
  if (isSheetProtected(state, sheet)) {
    return {
      title: strings.allowEditRangesDialogTitle,
      items: [
        {
          severity: 'warning',
          label: strings.allowEditRangesCommand,
          detail: strings.allowEditRangesRequiresUnprotected,
        },
      ],
    };
  }
  if (action === 'allow-edit-range') {
    const rangeText = formatA1Range(range);
    addAllowedEditRangeAction(context, range, { title: rangeText });
    return {
      title: strings.allowEditRangesDialogTitle,
      items: [
        {
          severity: 'info',
          label: strings.allowEditRangesCommand,
          detail: strings.allowedEditRangeAddedStatus.replace('{range}', rangeText),
        },
      ],
    };
  }
  clearAllowedEditRangesAction(context, state.data.sheetIndex);
  return {
    title: strings.allowEditRangesDialogTitle,
    items: [
      {
        severity: 'info',
        label: strings.allowEditRangesClearCommand,
        detail: strings.allowedEditRangesClearedStatus,
      },
    ],
  };
};
