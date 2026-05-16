// PDF split-button helper — same shape every ribbon host needs: route
// "Preferences" to the Page Setup dialog, otherwise kick the host's print
// pipeline and (for `create` / `share`) show a confirmation report.

import type { Strings } from '../i18n/strings.js';

type CellMenuStrings = Pick<
  Strings['ribbonMenu'],
  'pdfCreate' | 'pdfCreateReady' | 'pdfShare' | 'pdfShareReady'
>;

export type RibbonPdfAction = 'create' | 'share' | 'preferences';

export interface RibbonPdfReportItem {
  severity: 'info' | 'warning';
  label: string;
  detail: string;
}

export interface RibbonPdfReport {
  title: string;
  items: RibbonPdfReportItem[];
}

export type RibbonPdfActionResult =
  | { kind: 'open-page-setup' }
  | { kind: 'print'; report?: RibbonPdfReport };

export interface ResolveRibbonPdfActionStrings {
  cellMenu: CellMenuStrings;
  pdfTitle: string;
}

/** Convert a click on the PDF split-button into a host-routable verdict.
 *  Hosts handle `open-page-setup` by calling `instance.openPageSetup()` and
 *  `print` by calling `instance.print('pdf')` (then surfacing `report` when
 *  set, so the user sees the matching "ready" confirmation). */
export const resolveRibbonPdfAction = (
  action: RibbonPdfAction,
  strings: ResolveRibbonPdfActionStrings,
): RibbonPdfActionResult => {
  if (action === 'preferences') return { kind: 'open-page-setup' };
  const { cellMenu, pdfTitle } = strings;
  if (action === 'create') {
    return {
      kind: 'print',
      report: {
        title: pdfTitle,
        items: [{ severity: 'info', label: cellMenu.pdfCreate, detail: cellMenu.pdfCreateReady }],
      },
    };
  }
  return {
    kind: 'print',
    report: {
      title: pdfTitle,
      items: [{ severity: 'info', label: cellMenu.pdfShare, detail: cellMenu.pdfShareReady }],
    },
  };
};
