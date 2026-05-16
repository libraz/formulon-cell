// Shared "Link" ribbon split-button — translates the action into a host-routable
// directive: open the hyperlink dialog, open the external-links dialog, open
// a URL (the host owns `window.open` to satisfy CSP and tracking concerns),
// clear the hyperlink on the active cell, or surface a "no link here" report.

import type { History } from '../commands/history.js';
import { recordFormatChange } from '../commands/history.js';
import { clearHyperlink, hyperlinkAt } from '../commands/hyperlinks.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import type { SpreadsheetStore } from '../store/store.js';

type HyperlinkStrings = Pick<Strings['ribbonMenu'], 'linkOpen' | 'linkNoHyperlink'>;

export type RibbonHyperlinkAction = 'edit' | 'external' | 'open' | 'clear';

export interface RibbonHyperlinkReport {
  title: string;
  items: { severity: 'info' | 'warning'; label: string; detail: string }[];
}

export type RibbonHyperlinkActionResult =
  | { kind: 'open-hyperlink-dialog' }
  | { kind: 'open-external-dialog' }
  | { kind: 'open-url'; url: string }
  | { kind: 'cleared' }
  | { kind: 'report'; report: RibbonHyperlinkReport };

export interface ExecuteRibbonHyperlinkActionDeps {
  store: SpreadsheetStore;
  workbook: WorkbookHandle;
  history: History;
  action: RibbonHyperlinkAction;
  strings: HyperlinkStrings;
}

/** Convert a hyperlink-menu click into a host-routable verdict. The host
 *  handles `open-hyperlink-dialog` / `open-external-dialog` by calling the
 *  matching instance method, `open-url` by calling `window.open` with safe
 *  flags, and `report` by displaying the ribbon-report shell. `cleared`
 *  signals the mutation already ran via the supplied history. */
export const executeRibbonHyperlinkAction = (
  deps: ExecuteRibbonHyperlinkActionDeps,
): RibbonHyperlinkActionResult => {
  const { store, workbook, history, action, strings } = deps;
  if (action === 'edit') return { kind: 'open-hyperlink-dialog' };
  if (action === 'external') return { kind: 'open-external-dialog' };
  const state = store.getState();
  const target = hyperlinkAt(state, state.selection.active);
  if (!target) {
    return {
      kind: 'report',
      report: {
        title: strings.linkOpen,
        items: [{ severity: 'info', label: strings.linkNoHyperlink, detail: '' }],
      },
    };
  }
  if (action === 'open') return { kind: 'open-url', url: target };
  recordFormatChange(history, store, () => {
    clearHyperlink(store, state.selection.active, workbook);
  });
  return { kind: 'cleared' };
};
