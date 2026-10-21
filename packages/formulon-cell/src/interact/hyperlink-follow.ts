import { isInDocumentHyperlinkTarget, isSafeHyperlinkTarget } from '../commands/hyperlinks.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { resolveAreaText } from '../mount/ref-utils.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import {
  clampNavigationAddr,
  clampNavigationRange,
  scrollActiveIntoView,
} from './navigation-policy.js';

export interface HyperlinkFollowDeps {
  store: SpreadsheetStore;
  wb: WorkbookHandle;
  /** Make `sheet` the active sheet and hydrate it; false when it cannot be shown. */
  activateSheet?: (sheet: number) => boolean;
  /** Called when an in-document target does not resolve. */
  onInvalidTarget?: () => void;
  /** Opens an external URL; defaults to a `noopener` new tab. */
  openExternal?: (url: string) => void;
}

const openInNewTab = (url: string): void => {
  if (typeof window === 'undefined' || typeof window.open !== 'function') return;
  window.open(url, '_blank', 'noopener,noreferrer');
};

/**
 * Follow a hyperlink target: an in-document `#Sheet2!A1` / `#Name` selects the
 * destination and scrolls it into view, a safe external URL opens in a new
 * tab. Returns false when nothing was followed.
 */
export function followHyperlink(deps: HyperlinkFollowDeps, target: string): boolean {
  const { store, wb } = deps;
  const text = target.trim();
  if (isInDocumentHyperlinkTarget(text)) {
    const area = resolveAreaText(wb, text.slice(1), store.getState().data.sheetIndex);
    const range = area ? clampNavigationRange(store, area) : null;
    const anchor = range
      ? clampNavigationAddr(store, { sheet: range.sheet, row: range.r0, col: range.c0 })
      : null;
    if (!range || !anchor || range.sheet >= wb.sheetCount) {
      deps.onInvalidTarget?.();
      return false;
    }
    if (range.sheet !== store.getState().data.sheetIndex) {
      if (deps.activateSheet) {
        if (!deps.activateSheet(range.sheet)) return false;
      } else {
        mutators.setSheetIndex(store, range.sheet);
        mutators.replaceCells(store, wb.cells(range.sheet));
      }
      if (store.getState().data.sheetIndex !== range.sheet) return false;
    }
    mutators.setActive(store, anchor);
    if (range.r1 > range.r0 || range.c1 > range.c0) mutators.setRange(store, range);
    scrollActiveIntoView(store, anchor);
    return true;
  }
  if (!isSafeHyperlinkTarget(text)) return false;
  (deps.openExternal ?? openInNewTab)(text);
  return true;
}
