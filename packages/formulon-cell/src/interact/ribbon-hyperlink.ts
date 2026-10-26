// Shared "Link" ribbon split-button — translates the action into a host-routable
// directive: open the hyperlink dialog, open the external-links dialog, open
// a URL (the host owns `window.open` to satisfy CSP and tracking concerns),
// clear the hyperlink on the active cell, or surface a "no link here" report.

import type { History } from '../commands/history.js';
import {
  clearHyperlink,
  hyperlinkAt,
  isFollowableHyperlinkTarget,
} from '../commands/hyperlinks.js';
import { deferForAllowedRangeUnlock } from '../commands/protection.js';
import { recordFormatChange } from '../commands/slice-history.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import type { SpreadsheetStore } from '../store/store.js';

type HyperlinkStrings = Pick<Strings['ribbonMenu'], 'linkOpen' | 'linkNoHyperlink'> &
  Partial<Pick<Strings['ribbonMenu'], 'linkUnsupportedTarget'>>;

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
  | { kind: 'deferred' }
  | { kind: 'refused' }
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
 *  signals the mutation already ran via the supplied history, `deferred` that
 *  it waits for a range password, and `refused` that protection blocked it. */
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
  if (action === 'open') {
    const url = target.trim();
    if (!isFollowableHyperlinkTarget(url)) {
      return {
        kind: 'report',
        report: {
          title: strings.linkOpen,
          items: [
            {
              severity: 'warning',
              label: strings.linkUnsupportedTarget ?? strings.linkNoHyperlink,
              detail: '',
            },
          ],
        },
      };
    }
    return { kind: 'open-url', url };
  }
  const { active } = state.selection;
  const cell = {
    sheet: active.sheet,
    r0: active.row,
    c0: active.col,
    r1: active.row,
    c1: active.col,
  };
  const clear = (): void => {
    recordFormatChange(history, store, () => {
      clearHyperlink(store, active, workbook);
    });
  };
  // A password prompt makes the removal asynchronous; the caller hears `deferred`.
  if (deferForAllowedRangeUnlock(store, [cell], clear)) return { kind: 'deferred' };
  clear();
  return hyperlinkAt(store.getState(), active) === null ? { kind: 'cleared' } : { kind: 'refused' };
};
