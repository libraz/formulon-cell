import { copy } from '../commands/clipboard/copy.js';
import { cut } from '../commands/clipboard/cut.js';
import { encodeHtml } from '../commands/clipboard/html.js';
import { pasteTSV } from '../commands/clipboard/paste.js';
import { type ClipboardSnapshot, captureSnapshot } from '../commands/clipboard/snapshot.js';
import { parseTSV } from '../commands/clipboard/tsv.js';
import { applyUnmerge } from '../commands/merge.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';

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
  /** Shortcut-driven equivalents of the `copy`/`cut`/`paste` events. The
   *  browser only dispatches those events when focus is on an editable
   *  element or a real text selection is present; our canvas-backed grid
   *  satisfies neither, so the keyboard router routes Mod+C/X/V here. */
  runShortcut(kind: 'copy' | 'cut' | 'paste'): Promise<void>;
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
    if (!r || !e.clipboardData) {
      mutators.setCopyRange(store, null);
      return;
    }
    e.clipboardData.setData('text/plain', r.tsv);
    e.clipboardData.setData('text/html', encodeHtml(s, r.range));
    snapshot = captureSnapshot(s, r.range);
    if (r.ranges) mutators.setCopyRanges(store, r.ranges);
    else mutators.setCopyRange(store, r.range);
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
    mutators.setCopyRange(store, r.range);
    e.preventDefault();
    deps.onAfterCommit();
  };

  const onPaste = (e: ClipboardEvent): void => {
    const s = store.getState();
    if (s.ui.editor.kind !== 'idle') return;
    const text = e.clipboardData?.getData('text/plain') ?? '';
    if (!text) return;
    // Spreadsheet parity: any merge that intersects the destination range gets
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
    if (r) {
      mutators.setCopyRange(store, null);
      mutators.setRange(store, r.writtenRange);
      deps.onAfterCommit();
    }
  };

  host.addEventListener('copy', onCopy);
  host.addEventListener('cut', onCut);
  host.addEventListener('paste', onPaste);

  const writeClipboardText = async (tsv: string): Promise<void> => {
    try {
      await navigator.clipboard?.writeText(tsv);
    } catch (err) {
      console.warn('formulon-cell: clipboard write failed', err);
    }
  };

  const runShortcut = async (kind: 'copy' | 'cut' | 'paste'): Promise<void> => {
    const s = store.getState();
    if (s.ui.editor.kind !== 'idle') return;
    if (kind === 'copy') {
      const r = copy(s);
      if (!r) {
        mutators.setCopyRange(store, null);
        return;
      }
      snapshot = captureSnapshot(s, r.range);
      if (r.ranges) mutators.setCopyRanges(store, r.ranges);
      else mutators.setCopyRange(store, r.range);
      await writeClipboardText(r.tsv);
      return;
    }
    if (kind === 'cut') {
      snapshot = captureSnapshot(s, s.selection.range);
      const r = cut(s, wb);
      if (!r) return;
      mutators.setCopyRange(store, r.range);
      await writeClipboardText(r.tsv);
      deps.onAfterCommit();
      return;
    }
    // paste
    let text = '';
    try {
      text = (await navigator.clipboard?.readText()) ?? '';
    } catch (err) {
      console.warn('formulon-cell: clipboard read failed', err);
      return;
    }
    if (!text) return;
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
    if (r) {
      mutators.setCopyRange(store, null);
      mutators.setRange(store, r.writtenRange);
      deps.onAfterCommit();
    }
  };

  return {
    getSnapshot: () => snapshot,
    runShortcut,
    detach() {
      host.removeEventListener('copy', onCopy);
      host.removeEventListener('cut', onCut);
      host.removeEventListener('paste', onPaste);
    },
  };
}
