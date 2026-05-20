import { copy } from '../commands/clipboard/copy.js';
import { cut } from '../commands/clipboard/cut.js';
import { encodeHtml } from '../commands/clipboard/html.js';
import { pasteTSV } from '../commands/clipboard/paste.js';
import { pasteSpecial } from '../commands/clipboard/paste-special.js';
import { type ClipboardSnapshot, captureSnapshot } from '../commands/clipboard/snapshot.js';
import { parseTSV } from '../commands/clipboard/tsv.js';
import type { History } from '../commands/history.js';
import { recordFormatChange } from '../commands/history.js';
import { applyUnmerge } from '../commands/merge.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { mutators, type SpreadsheetStore, type State } from '../store/store.js';
import type { PasteOptionsActivation } from './paste-options.js';

export interface ClipboardDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  wb: WorkbookHandle;
  /** Bundles multi-cell clipboard mutations into a single undo step. */
  history?: History | null;
  /** Refresh the cached cell map after a write — same contract as the
   *  inline editor. */
  onAfterCommit: () => void;
  onPasteOptions?: (activation: PasteOptionsActivation) => void;
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
  const { history = null, host, store, wb } = deps;
  if (history) wb.attachHistory(history);

  let snapshot: ClipboardSnapshot | null = null;
  let snapshotText: string | null = null;

  const snapshotDestRange = (state: State, snap: ClipboardSnapshot): Range => ({
    sheet: state.selection.active.sheet,
    r0: state.selection.active.row,
    c0: state.selection.active.col,
    r1: state.selection.active.row + snap.rows - 1,
    c1: state.selection.active.col + snap.cols - 1,
  });
  const clearFormatsInRanges = (ranges: readonly Range[]): void => {
    store.setState((s) => {
      const formats = new Map(s.format.formats);
      for (const range of ranges) {
        for (let row = range.r0; row <= range.r1; row += 1) {
          for (let col = range.c0; col <= range.c1; col += 1) {
            formats.delete(`${range.sheet}:${row}:${col}`);
          }
        }
      }
      return { ...s, format: { ...s.format, formats } };
    });
  };

  const pasteFromClipboardText = (
    state: State,
    text: string,
  ): { result: { writtenRange: Range } | null; activation: PasteOptionsActivation | null } => {
    if (snapshot && snapshotText === text) {
      const source = snapshot;
      const before = captureSnapshot(state, snapshotDestRange(state, source));
      let result: { writtenRange: Range } | null = null;
      recordFormatChange(history, store, () => {
        result = pasteSpecial(state, store, wb, source, {
          what: 'all',
          operation: 'none',
          skipBlanks: false,
          transpose: false,
        });
      });
      const applied = result as { writtenRange: Range } | null;
      return {
        result: applied,
        activation: applied && before ? { source, before, range: { ...applied.writtenRange } } : null,
      };
    }
    snapshot = null;
    snapshotText = null;
    return { result: pasteTSV(state, wb, text), activation: null };
  };

  const onCopy = (e: ClipboardEvent): void => {
    const s = store.getState();
    if (s.ui.editor.kind !== 'idle') return; // let the input handle it
    const r = copy(s);
    if (!r || !e.clipboardData) {
      snapshot = null;
      snapshotText = null;
      mutators.setCopyRange(store, null);
      return;
    }
    e.clipboardData.setData('text/plain', r.tsv);
    e.clipboardData.setData('text/html', encodeHtml(s, r.range));
    snapshot = captureSnapshot(s, r.range);
    snapshotText = r.tsv;
    if (r.ranges) mutators.setCopyRanges(store, r.ranges);
    else mutators.setCopyRange(store, r.range);
    e.preventDefault();
  };

  const onCut = (e: ClipboardEvent): void => {
    const s = store.getState();
    if (s.ui.editor.kind !== 'idle') return;
    snapshot = captureSnapshot(s, s.selection.range);
    if (history) history.begin();
    let r: ReturnType<typeof cut> = null;
    try {
      r = cut(s, wb);
      if (r) {
        const ranges = r.payloadRanges ?? r.ranges ?? [r.range];
        recordFormatChange(history, store, () => clearFormatsInRanges(ranges));
      }
    } finally {
      if (history) history.end();
    }
    if (!r || !e.clipboardData) {
      snapshot = null;
      snapshotText = null;
      return;
    }
    e.clipboardData.setData('text/plain', r.tsv);
    e.clipboardData.setData('text/html', encodeHtml(s, r.range));
    snapshotText = r.tsv;
    mutators.setCopyRange(store, r.range);
    e.preventDefault();
    deps.onAfterCommit();
  };

  const onPaste = (e: ClipboardEvent): void => {
    const s = store.getState();
    if (s.ui.editor.kind !== 'idle') return;
    const text = e.clipboardData?.getData('text/plain') ?? '';
    if (!text) return;
    if (history) history.begin();
    let r: { writtenRange: Range } | null = null;
    let activation: PasteOptionsActivation | null = null;
    try {
      // Spreadsheet parity: any merge that intersects the destination range gets
      // unmerged before the paste — a textual paste cannot tear merged cells.
      const rows = parseTSV(text);
      if (rows.length > 0) {
        const origin = s.selection.active;
        let maxCols = 0;
        for (const row of rows) if (row.length > maxCols) maxCols = row.length;
        applyUnmerge(store, wb, history, {
          sheet: origin.sheet,
          r0: origin.row,
          c0: origin.col,
          r1: origin.row + rows.length - 1,
          c1: origin.col + Math.max(0, maxCols - 1),
        });
      }
      ({ result: r, activation } = pasteFromClipboardText(s, text));
    } finally {
      if (history) history.end();
    }
    e.preventDefault();
    if (r) {
      mutators.setCopyRange(store, null);
      mutators.setRange(store, r.writtenRange);
      deps.onAfterCommit();
      if (activation) deps.onPasteOptions?.(activation);
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
        snapshot = null;
        snapshotText = null;
        mutators.setCopyRange(store, null);
        return;
      }
      snapshot = captureSnapshot(s, r.range);
      snapshotText = r.tsv;
      if (r.ranges) mutators.setCopyRanges(store, r.ranges);
      else mutators.setCopyRange(store, r.range);
      await writeClipboardText(r.tsv);
      return;
    }
    if (kind === 'cut') {
      snapshot = captureSnapshot(s, s.selection.range);
      if (history) history.begin();
      let r: ReturnType<typeof cut> = null;
      try {
        r = cut(s, wb);
        if (r) {
          const ranges = r.payloadRanges ?? r.ranges ?? [r.range];
          recordFormatChange(history, store, () => clearFormatsInRanges(ranges));
        }
      } finally {
        if (history) history.end();
      }
      if (!r) {
        snapshot = null;
        snapshotText = null;
        return;
      }
      snapshotText = r.tsv;
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
    if (history) history.begin();
    let r: { writtenRange: Range } | null = null;
    let activation: PasteOptionsActivation | null = null;
    try {
      const rows = parseTSV(text);
      if (rows.length > 0) {
        const origin = s.selection.active;
        let maxCols = 0;
        for (const row of rows) if (row.length > maxCols) maxCols = row.length;
        applyUnmerge(store, wb, history, {
          sheet: origin.sheet,
          r0: origin.row,
          c0: origin.col,
          r1: origin.row + rows.length - 1,
          c1: origin.col + Math.max(0, maxCols - 1),
        });
      }
      ({ result: r, activation } = pasteFromClipboardText(s, text));
    } finally {
      if (history) history.end();
    }
    if (r) {
      mutators.setCopyRange(store, null);
      mutators.setRange(store, r.writtenRange);
      deps.onAfterCommit();
      if (activation) deps.onPasteOptions?.(activation);
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
