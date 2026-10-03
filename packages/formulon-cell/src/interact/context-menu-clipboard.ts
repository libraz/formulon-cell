import { copy } from '../commands/clipboard/copy.js';
import { cut } from '../commands/clipboard/cut.js';
import { pasteTSV } from '../commands/clipboard/paste.js';
import {
  type PasteWhat,
  pasteSpecial,
  resolvePasteDestination,
} from '../commands/clipboard/paste-special.js';
import {
  type ClipboardSnapshot,
  captureSnapshotFromCopyResult,
} from '../commands/clipboard/snapshot.js';
import { parseTSV } from '../commands/clipboard/tsv.js';
import type { History } from '../commands/history.js';
import { addrKey, MAX_COL, MAX_ROW } from '../engine/address.js';
import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { mutators, type SpreadsheetStore, type State } from '../store/store.js';
import type { ContextMenuDeps } from './context-menu.js';
import type { ItemId } from './context-menu-spec.js';

export interface ContextMenuClipboardContext {
  readonly store: SpreadsheetStore;
  readonly wb: WorkbookHandle;
  readonly history: History | null;
  /** Host clipboard hooks, read at call time like the rest of `deps`. */
  readonly deps: Pick<ContextMenuDeps, 'getClipboardSnapshot' | 'onClipboardShortcut'>;
  /** Forwards to the host's `onAfterCommit`, looked up at call time. */
  readonly afterCommit: () => void;
  readonly hasExplicitPolicy: () => boolean;
  readonly canPasteToRange: (range: Range | null) => boolean;
  /** Restricted-mode paste: routes TSV text through the interaction controller. */
  readonly executeClipboardText: (text: string) => void;
}

/** The menu's own copy/cut snapshot plus the copy, cut, paste and quick-paste
 *  items that read or write it. */
export interface ContextMenuClipboard {
  /** Host snapshot when wired, otherwise the menu's own copy while its
   *  marquee is still the current one. */
  snapshot(): ClipboardSnapshot | null;
  /** Returns false for any item other than copy, cut, paste and the quick
   *  Paste Special entries. */
  run(id: ItemId, state: State): boolean;
}

const pasteDestinationRange = (origin: Addr, rows: number, cols: number): Range | null => {
  if (rows <= 0 || cols <= 0) return null;
  const r1 = origin.row + rows - 1;
  const c1 = origin.col + cols - 1;
  if (r1 > MAX_ROW || c1 > MAX_COL) return null;
  return {
    sheet: origin.sheet,
    r0: origin.row,
    c0: origin.col,
    r1,
    c1,
  };
};

export const hasPastePayload = (
  text: string,
  snap: ClipboardSnapshot | null | undefined,
): boolean => text.length > 0 || snap != null;

/** A cut can only be pasted once, so its marquee is consumed by the paste.
 *  A copy marquee stays up for repeat pastes, exactly like the desktop app. */
export const consumeCutMarquee = (store: SpreadsheetStore): void => {
  if (store.getState().ui.copyMode !== 'cut') return;
  mutators.setCopyRange(store, null);
  mutators.setCopyRanges(store, null);
};

export function createContextMenuClipboard(ctx: ContextMenuClipboardContext): ContextMenuClipboard {
  const { store, wb, history, deps } = ctx;
  let localSnapshot: ClipboardSnapshot | null = null;
  let localSnapshotText: string | null = null;
  let localCopyRevision: number | null = null;
  const clipboardSnapshot = (): ClipboardSnapshot | null => {
    if (deps.getClipboardSnapshot) return deps.getClipboardSnapshot();
    const state = store.getState();
    const ui = state.ui;
    if (
      !localSnapshot ||
      localCopyRevision !== (ui.copyRevision ?? 0) ||
      ui.copyMode !== localSnapshot.mode ||
      !(ui.copyRange || ui.copyRanges?.length)
    )
      return null;
    if (localSnapshot.mode === 'copy' && ui.copyRange) {
      const sourceState: State = {
        ...state,
        data: {
          ...state.data,
          sheetIndex: ui.copyRange.sheet,
          cells: new Map(
            Array.from(wb.cells(ui.copyRange.sheet), (cell) => [
              addrKey(cell.addr),
              { value: cell.value, formula: cell.formula },
            ]),
          ),
        },
        selection: { ...state.selection, range: ui.copyRange, extraRanges: [] },
      };
      const materialized = copy(sourceState);
      localSnapshot = materialized
        ? captureSnapshotFromCopyResult(sourceState, materialized, 'copy')
        : null;
    }
    return localSnapshot;
  };

  function runPasteSpecial(what: PasteWhat, transpose: boolean): void {
    if (ctx.hasExplicitPolicy()) {
      void readClipboard().then((text) => {
        if (text) ctx.executeClipboardText(text);
      });
      return;
    }
    const snap = clipboardSnapshot();
    if (!snap) return;
    const state = store.getState();
    const destination = resolvePasteDestination(state, snap, transpose);
    if (!ctx.canPasteToRange(destination)) return;
    if (history) history.begin();
    try {
      pasteSpecial(
        store.getState(),
        store,
        wb,
        snap,
        { what, operation: 'none', skipBlanks: false, transpose },
        history,
      );
    } catch (err) {
      console.warn('formulon-cell: paste special failed', err);
    } finally {
      if (history) history.end();
    }
    ctx.afterCommit();
  }

  const run = (id: ItemId, state: State): boolean => {
    switch (id) {
      case 'copy': {
        if (deps.onClipboardShortcut) {
          deps.onClipboardShortcut('copy');
          return true;
        }
        const r = copy(state);
        if (r) {
          localSnapshot = captureSnapshotFromCopyResult(state, r, 'copy');
          localSnapshotText = r.tsv;
          if (r.ranges) mutators.setCopyRanges(store, r.ranges);
          else mutators.setCopyRange(store, r.range);
          localCopyRevision = store.getState().ui.copyRevision ?? 0;
          void writeClipboard(r.tsv);
        } else {
          localSnapshot = null;
          mutators.setCopyRange(store, null);
        }
        return true;
      }
      case 'cut': {
        if (deps.onClipboardShortcut) {
          deps.onClipboardShortcut('cut');
          return true;
        }
        const r = cut(state, wb);
        if (r) {
          localSnapshot = captureSnapshotFromCopyResult(state, r, 'cut');
          if (!localSnapshot) return true;
          localSnapshotText = r.tsv;
          mutators.setCopyRange(store, r.range, 'cut');
          localCopyRevision = store.getState().ui.copyRevision ?? 0;
          void writeClipboard(r.tsv);
        }
        return true;
      }
      case 'paste': {
        if (deps.onClipboardShortcut) {
          deps.onClipboardShortcut('paste');
          return true;
        }
        void readClipboard().then((text) => {
          const available = clipboardSnapshot();
          const snap =
            available && (!text || deps.getClipboardSnapshot || text === localSnapshotText)
              ? available
              : null;
          if (!hasPastePayload(text, snap)) return;
          if (ctx.hasExplicitPolicy()) {
            if (text) {
              ctx.executeClipboardText(text);
            }
            return;
          }
          const pasteState = store.getState();
          const tsvRows = snap ? [] : parseTSV(text);
          const rows = snap ? snap.rows : tsvRows.length;
          const cols = snap
            ? snap.cols
            : tsvRows.reduce((max, row) => Math.max(max, row.length), 0);
          const destination = snap
            ? resolvePasteDestination(pasteState, snap)
            : pasteDestinationRange(pasteState.selection.active, rows, cols);
          if (!ctx.canPasteToRange(destination)) {
            return;
          }
          if (history) history.begin();
          let r: ReturnType<typeof pasteTSV> | ReturnType<typeof pasteSpecial> = null;
          try {
            r = snap
              ? pasteSpecial(
                  store.getState(),
                  store,
                  wb,
                  snap,
                  { what: 'all', operation: 'none', skipBlanks: false, transpose: false },
                  history,
                )
              : text
                ? pasteTSV(store.getState(), wb, text)
                : null;
          } finally {
            if (history) history.end();
          }
          if (r) {
            consumeCutMarquee(store);
            mutators.setRange(store, r.writtenRange);
          }
          ctx.afterCommit();
        });
        return true;
      }
      case 'pasteAll':
        runPasteSpecial('all', false);
        return true;
      case 'pasteFormulas':
        runPasteSpecial('formulas', false);
        return true;
      case 'pasteFormulasNumFmt':
        runPasteSpecial('formulas-and-numfmt', false);
        return true;
      case 'pasteValues':
        runPasteSpecial('values', false);
        return true;
      case 'pasteValuesNumFmt':
        runPasteSpecial('values-and-numfmt', false);
        return true;
      case 'pasteFormatsOnly':
        runPasteSpecial('formats', false);
        return true;
      case 'pasteTranspose':
        runPasteSpecial('all', true);
        return true;
      default:
        return false;
    }
  };

  return { snapshot: clipboardSnapshot, run };
}

export function canReadClipboard(): boolean {
  return typeof navigator !== 'undefined' && typeof navigator.clipboard?.readText === 'function';
}

async function writeClipboard(text: string): Promise<void> {
  if (typeof navigator === 'undefined' || typeof navigator.clipboard?.writeText !== 'function') {
    return;
  }
  try {
    await navigator.clipboard.writeText(text);
  } catch (err) {
    console.warn('formulon-cell: clipboard write failed', err);
  }
}

export async function readClipboard(): Promise<string> {
  if (!canReadClipboard()) return '';
  try {
    return await navigator.clipboard.readText();
  } catch (err) {
    console.warn('formulon-cell: clipboard read failed', err);
    return '';
  }
}
