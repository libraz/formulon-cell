import { copy } from '../commands/clipboard/copy.js';
import { cut } from '../commands/clipboard/cut.js';
import { encodeHtml } from '../commands/clipboard/html.js';
import { pasteTSV } from '../commands/clipboard/paste.js';
import { type ClipboardSnapshot, captureSnapshot } from '../commands/clipboard/snapshot.js';
import { parseTSV } from '../commands/clipboard/tsv.js';
import { applyUnmerge } from '../commands/merge.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SpreadsheetStore } from '../store/store.js';

export interface ClipboardDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  wb: WorkbookHandle;
  /** Refresh the cached cell map after a write — same contract as the
   *  inline editor. */
  onAfterCommit: () => void;
}

export interface ClipboardHandle {
  /** Module-level structured snapshot — set by copy/cut, read by Paste Special.
   *  Cleared when the user copies from outside (system-clipboard-only events). */
  getSnapshot(): ClipboardSnapshot | null;
  detach(): void;
}

/**
 * Hook the host's `copy` / `cut` / `paste` events into the corresponding
 * commands. The host element must be focusable (tabindex) for the browser
 * to emit these events when the grid is the active region.
 */
export function attachClipboard(deps: ClipboardDeps): ClipboardHandle {
  const { host, store, wb } = deps;

  let snapshot: ClipboardSnapshot | null = null;

  const onCopy = (e: ClipboardEvent): void => {
    const s = store.getState();
    if (s.ui.editor.kind !== 'idle') return; // let the input handle it
    const r = copy(s);
    if (!r || !e.clipboardData) return;
    e.clipboardData.setData('text/plain', r.tsv);
    e.clipboardData.setData('text/html', encodeHtml(s, r.range));
    snapshot = captureSnapshot(s, r.range);
    e.preventDefault();
  };

  const onCut = (e: ClipboardEvent): void => {
    const s = store.getState();
    if (s.ui.editor.kind !== 'idle') return;
    snapshot = captureSnapshot(s, s.selection.range);
    const r = cut(s, wb);
    if (!r || !e.clipboardData) return;
    e.clipboardData.setData('text/plain', r.tsv);
    e.clipboardData.setData('text/html', encodeHtml(s, r.range));
    e.preventDefault();
    deps.onAfterCommit();
  };

  const onPaste = (e: ClipboardEvent): void => {
    const s = store.getState();
    if (s.ui.editor.kind !== 'idle') return;
    const text = e.clipboardData?.getData('text/plain') ?? '';
    if (!text) return;
    // Excel parity: any merge that intersects the destination range gets
    // unmerged before the paste — a textual paste cannot tear merged cells.
    const rows = parseTSV(text);
    if (rows.length > 0) {
      const origin = s.selection.active;
      let maxCols = 0;
      for (const row of rows) if (row.length > maxCols) maxCols = row.length;
      applyUnmerge(store, wb, null, {
        sheet: origin.sheet,
        r0: origin.row,
        c0: origin.col,
        r1: origin.row + rows.length - 1,
        c1: origin.col + Math.max(0, maxCols - 1),
      });
    }
    const r = pasteTSV(s, wb, text);
    e.preventDefault();
    if (r) deps.onAfterCommit();
  };

  host.addEventListener('copy', onCopy);
  host.addEventListener('cut', onCut);
  host.addEventListener('paste', onPaste);

  return {
    getSnapshot: () => snapshot,
    detach() {
      host.removeEventListener('copy', onCopy);
      host.removeEventListener('cut', onCut);
      host.removeEventListener('paste', onPaste);
    },
  };
}
