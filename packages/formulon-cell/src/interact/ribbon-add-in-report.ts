// Builds the "Add-Ins" split-button report dialog every ribbon host shows
// when the action is informational (get / my / manage / fallback). The host
// wrapper is responsible for the *real* "open my add-in" flow via its own
// callback — this helper only owns the i18n table.

import type { Strings } from '../i18n/strings.js';

type CellMenuStrings = Pick<
  Strings['ribbonMenu'],
  | 'addInGet'
  | 'addInMy'
  | 'addInManage'
  | 'addInManagedStatus'
  | 'addInStoreLabel'
  | 'addInStoreDetail'
  | 'addInBuiltInLabel'
  | 'addInBuiltInDetail'
  | 'addInExternalLabel'
  | 'addInExternalDetail'
>;

export type RibbonAddInAction = 'get' | 'my' | 'manage' | 'launch';

export interface RibbonAddInReportItem {
  severity: 'info' | 'warning';
  label: string;
  detail: string;
}

export interface RibbonAddInReport {
  title: string;
  items: RibbonAddInReportItem[];
}

export interface BuildRibbonAddInReportStrings {
  cellMenu: CellMenuStrings;
  addInDefaultTitle: string;
}

/** Resolve a `get`/`my`/`manage` add-in click into the report dialog payload.
 *  Returns `null` for `'launch'`, signalling the host should call its
 *  `onAddIn` callback (or, when no callback is wired, treat the "no callback"
 *  branch by calling [[buildRibbonAddInDefaultReport]]). */
export const buildRibbonAddInReport = (
  action: RibbonAddInAction,
  strings: BuildRibbonAddInReportStrings,
): RibbonAddInReport | null => {
  const { cellMenu, addInDefaultTitle } = strings;
  if (action === 'get') {
    return {
      title: cellMenu.addInGet,
      items: [
        { severity: 'info', label: cellMenu.addInStoreLabel, detail: cellMenu.addInStoreDetail },
        {
          severity: 'info',
          label: cellMenu.addInBuiltInLabel,
          detail: cellMenu.addInBuiltInDetail,
        },
      ],
    };
  }
  if (action === 'manage') {
    return {
      title: cellMenu.addInManage,
      items: [
        {
          severity: 'info',
          label: cellMenu.addInManagedStatus,
          detail: cellMenu.addInExternalDetail,
        },
      ],
    };
  }
  if (action === 'my') {
    return {
      title: cellMenu.addInMy,
      items: [
        {
          severity: 'info',
          label: cellMenu.addInBuiltInLabel,
          detail: cellMenu.addInBuiltInDetail,
        },
        {
          severity: 'info',
          label: cellMenu.addInExternalLabel,
          detail: cellMenu.addInExternalDetail,
        },
      ],
    };
  }
  return {
    title: addInDefaultTitle,
    items: [
      { severity: 'info', label: cellMenu.addInBuiltInLabel, detail: cellMenu.addInBuiltInDetail },
      {
        severity: 'info',
        label: cellMenu.addInExternalLabel,
        detail: cellMenu.addInExternalDetail,
      },
    ],
  };
};
