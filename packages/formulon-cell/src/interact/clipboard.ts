import { type CopyResult, copy } from '../commands/clipboard/copy.js';
import { cut } from '../commands/clipboard/cut.js';
import { encodeHtml } from '../commands/clipboard/html.js';
import { pasteTSV } from '../commands/clipboard/paste.js';
import {
  materializedPasteCells,
  pasteSpecial,
  resolvePasteDestination,
} from '../commands/clipboard/paste-special.js';
import {
  type ClipboardSnapshot,
  captureSnapshotFromCopyResult,
} from '../commands/clipboard/snapshot.js';
import { encodeTSV, parseTSV } from '../commands/clipboard/tsv.js';
import type { History } from '../commands/history.js';
import { interactionControllerFor } from '../commands/interaction-controller.js';
import { addrKey } from '../engine/address.js';
import type { Range } from '../engine/types.js';
import { formatCell } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { mutators, type SpreadsheetStore, type State } from '../store/store.js';
import { navigationBoundsFor } from './navigation-policy.js';
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
  let payloadRevision: number | null = null;
  let disposed = false;

  const isRestricted = (): boolean => interactionControllerFor(store)?.policy !== undefined;

  /** Restricted internal copies intentionally retain cell positions and values
   * only. Re-pasting them must not carry formulas, formats, or cut semantics. */
  const valueOnlySnapshot = (source: ClipboardSnapshot): ClipboardSnapshot => ({
    ...source,
    mode: 'copy',
    rowHeights: new Map(),
    colWidths: new Map(),
    cells: source.cells.map((row) =>
      row.map((cell) => ({ formula: null, value: cell.value, format: undefined })),
    ),
  });

  const valueOnlyText = (source: ClipboardSnapshot): string =>
    encodeTSV(source.cells.map((row) => row.map((cell) => formatCell(cell.value))));

  const isWithinNavigationBounds = (range: Range | null): boolean => {
    if (!range) return false;
    const bounds = navigationBoundsFor(store);
    if (!bounds) return true;
    return (
      range.sheet === bounds.sheet &&
      range.r0 >= bounds.r0 &&
      range.c0 >= bounds.c0 &&
      range.r1 <= bounds.r1 &&
      range.c1 <= bounds.c1
    );
  };

  const restrictedPaste = (state: State, text: string): { writtenRange: Range } | null => {
    const controller = interactionControllerFor(store);
    if (!controller || controller.policy === undefined) return null;
    const internal =
      snapshot && snapshotText === text && hasLiveInternalPayload(state) ? snapshot : null;
    const parsed = internal ? null : parseTSV(text);
    const destination = internal
      ? resolvePasteDestination(state, internal)
      : pasteDestinationRange(state, text);
    if (!destination || !isWithinNavigationBounds(destination)) return null;
    const changes: Array<
      | {
          addr: { sheet: number; row: number; col: number };
          value: import('../engine/types.js').CellValue;
        }
      | { addr: { sheet: number; row: number; col: number }; input: string }
    > = [];
    const destinationRows = destination.r1 - destination.r0 + 1;
    const destinationCols = destination.c1 - destination.c0 + 1;
    const scalarMerge =
      internal && internal.rows === 1 && internal.cols === 1
        ? [...state.merges.byAnchor.values()].find(
            (merge) =>
              merge.sheet === destination.sheet &&
              merge.r0 === destination.r0 &&
              merge.c0 === destination.c0 &&
              merge.r1 === destination.r1 &&
              merge.c1 === destination.c1,
          )
        : undefined;
    if (internal) {
      const bandCells = materializedPasteCells(internal, destination);
      if (bandCells) {
        for (const cell of bandCells) {
          if (scalarMerge && (cell.row !== scalarMerge.r0 || cell.col !== scalarMerge.c0)) {
            continue;
          }
          changes.push({
            addr: { sheet: destination.sheet, row: cell.row, col: cell.col },
            value: internal.cells[cell.sourceRowIndex]?.[cell.sourceColIndex]?.value ?? {
              kind: 'blank',
            },
          });
        }
      } else {
        for (let row = 0; row < destinationRows; row += 1) {
          const sourceRow = row % internal.rows;
          for (let col = 0; col < destinationCols; col += 1) {
            if (
              scalarMerge &&
              (destination.r0 + row !== scalarMerge.r0 || destination.c0 + col !== scalarMerge.c0)
            ) {
              continue;
            }
            const sourceCol = col % internal.cols;
            const addr = {
              sheet: destination.sheet,
              row: destination.r0 + row,
              col: destination.c0 + col,
            };
            changes.push({
              addr,
              value: internal.cells[sourceRow]?.[sourceCol]?.value ?? { kind: 'blank' },
            });
          }
        }
      }
    } else {
      for (let row = 0; row < (parsed?.length ?? 0); row += 1) {
        const line = parsed?.[row] ?? [];
        for (let col = 0; col < line.length; col += 1) {
          const addr = {
            sheet: destination.sheet,
            row: destination.r0 + row,
            col: destination.c0 + col,
          };
          changes.push({ addr, input: line[col] ?? '' });
        }
      }
    }
    if (changes.length === 0) return null;
    const result = controller.execute({
      type: 'cellBatch',
      operation: 'paste',
      origin: 'clipboard',
      changes,
      // Keep the matrix coordinates stable when a fixed form opts into
      // positional partial application. The controller's policy supplies the
      // default when this is omitted; an explicit value here would silently
      // turn `skipIneligible` back into an atomic rejection.
      denied: controller.policy.batchDenied,
    });
    if (result.status === 'rejected' || result.status === 'noop') return null;
    return { writtenRange: destination };
  };

  const hasLiveInternalPayload = (state: State): boolean => {
    const ui = state.ui;
    if (
      snapshotText === null ||
      payloadRevision === null ||
      payloadRevision !== (ui.copyRevision ?? 0) ||
      !ui.copyMode ||
      (!ui.copyRange && !ui.copyRanges?.length)
    ) {
      return false;
    }
    // Structural edits move the live copy marquee without changing the
    // clipboard revision. Refresh a copy snapshot from that live range so
    // formulas, formats, and merge topology come from the cells' new
    // coordinates. Whole-row/column copies intentionally use a trimmed
    // snapshot; `copy` re-materializes that payload instead of attempting to
    // capture the million-cell marquee itself. A cut payload is kept frozen
    // until its deferred move succeeds.
    if (
      snapshot?.mode === 'copy' &&
      ui.copyMode === 'copy' &&
      ui.copyRange &&
      (!ui.copyRanges || ui.copyRanges.length <= 1)
    ) {
      // Rebuild the source-sheet cache from the engine first. The visible
      // store cache can still belong to the active destination sheet after a
      // sheet switch, while the internal copy remains anchored elsewhere.
      const sourceSheet = ui.copyRange.sheet;
      const sourceCells = new Map(state.data.cells);
      for (const key of sourceCells.keys()) {
        if (key.startsWith(`${sourceSheet}:`)) sourceCells.delete(key);
      }
      for (const cell of wb.cells(sourceSheet)) {
        sourceCells.set(addrKey(cell.addr), { value: cell.value, formula: cell.formula });
      }
      const sourceIsActive = sourceSheet === state.data.sheetIndex;
      const sourceLogical = { ...ui.copyRange };
      const sourceMerges = sourceIsActive
        ? state.merges
        : (() => {
            const engineMerges = wb.getMerges(sourceSheet);
            const merges =
              engineMerges.length > 0
                ? engineMerges
                : (snapshot?.merges ?? []).map((merge) => ({
                    sheet: sourceSheet,
                    r0: (snapshot?.range.r0 ?? 0) + merge.r0,
                    c0: (snapshot?.range.c0 ?? 0) + merge.c0,
                    r1: (snapshot?.range.r0 ?? 0) + merge.r1,
                    c1: (snapshot?.range.c0 ?? 0) + merge.c1,
                  }));
            const byAnchor = new Map<string, Range>();
            for (const merge of merges) {
              byAnchor.set(addrKey({ sheet: sourceSheet, row: merge.r0, col: merge.c0 }), merge);
            }
            return { byAnchor, byCell: new Map<string, string>() };
          })();
      const sourceLayout = sourceIsActive
        ? state.layout
        : (() => {
            const rowHeights = new Map<number, number>();
            const colWidths = new Map<number, number>();
            for (const [offset, height] of snapshot?.rowHeights ?? []) {
              rowHeights.set(sourceLogical.r0 + offset, height);
            }
            for (const [offset, width] of snapshot?.colWidths ?? []) {
              colWidths.set(sourceLogical.c0 + offset, width);
            }
            return { ...state.layout, rowHeights, colWidths };
          })();
      const sourceState: State = {
        ...state,
        data: { ...state.data, sheetIndex: sourceSheet, cells: sourceCells },
        merges: sourceMerges,
        layout: sourceLayout,
        selection: { ...state.selection, range: { ...ui.copyRange }, extraRanges: [] },
      };
      const materialized = copy(sourceState);
      const refreshed = materialized
        ? captureSnapshotFromCopyResult(sourceState, materialized, 'copy')
        : null;
      // Keep snapshotText untouched: it is the system clipboard token that
      // proves this is still the workbook's internal payload. Only the
      // structured cell snapshot is re-materialized.
      snapshot = refreshed ? (isRestricted() ? valueOnlySnapshot(refreshed) : refreshed) : null;
    }
    return snapshot === null || snapshot.mode === ui.copyMode;
  };

  const activeInternalText = (): string | null => {
    const state = store.getState();
    return hasLiveInternalPayload(state) ? snapshotText : null;
  };

  const snapshotDestRange = (
    state: State,
    snap: ClipboardSnapshot,
    transpose = false,
  ): Range | null => resolvePasteDestination(state, snap, transpose);
  const pasteDestinationRange = (state: State, text: string): Range | null => {
    const internal =
      snapshot && snapshotText === text && hasLiveInternalPayload(state) ? snapshot : null;
    if (internal) return snapshotDestRange(state, internal);
    const rows = parseTSV(text);
    if (rows.length === 0) return null;
    let maxCols = 0;
    for (const row of rows) maxCols = Math.max(maxCols, row.length);
    return {
      sheet: state.selection.active.sheet,
      r0: state.selection.active.row,
      c0: state.selection.active.col,
      r1: state.selection.active.row + rows.length - 1,
      c1: state.selection.active.col + Math.max(0, maxCols - 1),
    };
  };
  const materializedRanges = (result: CopyResult): Range[] =>
    result.payloadRanges ?? [result.range];
  const captureMaterializedSnapshot = (
    state: State,
    result: CopyResult,
    mode: 'copy' | 'cut',
  ): ClipboardSnapshot | null => {
    return captureSnapshotFromCopyResult(state, result, mode);
  };
  const encodeMaterializedHtml = (state: State, result: CopyResult): string =>
    materializedRanges(result)
      .map((range) => encodeHtml(state, range))
      .join('');
  const hasPastePayload = (text: string): boolean =>
    text.length > 0 ||
    (snapshot !== null && snapshotText === text && hasLiveInternalPayload(store.getState()));

  const captureDestinationSnapshot = (state: State, range: Range): ClipboardSnapshot | null => {
    const captureState: State = {
      ...state,
      selection: {
        ...state.selection,
        active: { sheet: range.sheet, row: range.r0, col: range.c0 },
        anchor: { sheet: range.sheet, row: range.r0, col: range.c0 },
        range: { ...range },
        extraRanges: [],
      },
    };
    const result = copy(captureState);
    return result ? captureSnapshotFromCopyResult(captureState, result) : null;
  };

  /** A cut can only be pasted once, so its marquee is consumed by the paste.
   *  A copy marquee stays up for repeat pastes, exactly like the desktop app. */
  const consumeCutMarquee = (): void => {
    if (store.getState().ui.copyMode === 'cut') {
      mutators.setCopyRange(store, null);
      mutators.setCopyRanges(store, null);
    } else if (snapshot?.mode !== 'cut') {
      return;
    }
    snapshot = null;
    snapshotText = null;
    payloadRevision = null;
  };

  const pasteFromClipboardText = (
    state: State,
    text: string,
  ): { result: { writtenRange: Range } | null; activation: PasteOptionsActivation | null } => {
    if (snapshot && snapshotText === text && hasLiveInternalPayload(state)) {
      const source = snapshot;
      const beforeRange = snapshotDestRange(state, source);
      const before = beforeRange ? captureDestinationSnapshot(state, beforeRange) : null;
      let result: { writtenRange: Range } | null = null;
      result = pasteSpecial(
        state,
        store,
        wb,
        source,
        {
          what: 'all',
          operation: 'none',
          skipBlanks: false,
          transpose: false,
        },
        history,
      );
      const applied = result as { writtenRange: Range } | null;
      return {
        result: applied,
        activation:
          applied && before && source.mode === 'copy'
            ? { source, before, range: { ...applied.writtenRange } }
            : null,
      };
    }
    snapshot = null;
    snapshotText = null;
    payloadRevision = null;
    return { result: pasteTSV(state, wb, text), activation: null };
  };

  const onCopy = (e: ClipboardEvent): void => {
    const s = store.getState();
    if (s.ui.editor.kind !== 'idle') return; // let the input handle it
    const controller = interactionControllerFor(store);
    if (controller?.policy?.copy === false) {
      e.preventDefault();
      return;
    }
    const r = copy(s);
    if (!r || !e.clipboardData) {
      snapshot = null;
      snapshotText = null;
      payloadRevision = null;
      mutators.setCopyRange(store, null);
      return;
    }
    const captured = captureMaterializedSnapshot(s, r, 'copy');
    snapshot = isRestricted() && captured ? valueOnlySnapshot(captured) : captured;
    snapshotText = snapshot && isRestricted() ? valueOnlyText(snapshot) : r.tsv;
    e.clipboardData.setData('text/plain', snapshotText ?? r.tsv);
    if (!isRestricted()) e.clipboardData.setData('text/html', encodeMaterializedHtml(s, r));
    if (r.ranges) mutators.setCopyRanges(store, r.ranges);
    else mutators.setCopyRange(store, r.range);
    payloadRevision = store.getState().ui.copyRevision ?? 0;
    e.preventDefault();
  };

  const onCut = (e: ClipboardEvent): void => {
    const s = store.getState();
    if (s.ui.editor.kind !== 'idle') return;
    if (isRestricted()) {
      // Cut is a move operation. The restricted route has no atomic move
      // executor yet, so fail closed before touching data or the marquee.
      e.preventDefault();
      return;
    }
    const r = cut(s, wb);
    if (!r || !e.clipboardData) {
      snapshot = null;
      snapshotText = null;
      payloadRevision = null;
      return;
    }
    e.clipboardData.setData('text/plain', r.tsv);
    e.clipboardData.setData('text/html', encodeMaterializedHtml(s, r));
    snapshot = captureMaterializedSnapshot(s, r, 'cut');
    snapshotText = r.tsv;
    mutators.setCopyRange(store, r.range, 'cut');
    payloadRevision = store.getState().ui.copyRevision ?? 0;
    e.preventDefault();
  };

  const onPaste = (e: ClipboardEvent): void => {
    const s = store.getState();
    if (s.ui.editor.kind !== 'idle') return;
    const text = e.clipboardData?.getData('text/plain') ?? '';
    if (!hasPastePayload(text)) return;
    if (isRestricted()) {
      e.preventDefault();
      const r = restrictedPaste(s, text);
      if (r) {
        consumeCutMarquee();
        mutators.setRange(store, r.writtenRange);
        deps.onAfterCommit();
      }
      return;
    }
    if (!isWithinNavigationBounds(pasteDestinationRange(s, text))) {
      e.preventDefault();
      return;
    }
    if (history) history.begin();
    let r: { writtenRange: Range } | null = null;
    let activation: PasteOptionsActivation | null = null;
    try {
      ({ result: r, activation } = pasteFromClipboardText(s, text));
    } finally {
      if (history) history.end();
    }
    e.preventDefault();
    if (r) {
      consumeCutMarquee();
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
    if (disposed) return;
    const s = store.getState();
    if (s.ui.editor.kind !== 'idle') return;
    if (kind === 'copy') {
      const controller = interactionControllerFor(store);
      if (controller?.policy?.copy === false) return;
      const r = copy(s);
      if (!r) {
        snapshot = null;
        snapshotText = null;
        payloadRevision = null;
        mutators.setCopyRange(store, null);
        return;
      }
      const captured = captureMaterializedSnapshot(s, r, 'copy');
      snapshot = isRestricted() && captured ? valueOnlySnapshot(captured) : captured;
      snapshotText = snapshot && isRestricted() ? valueOnlyText(snapshot) : r.tsv;
      if (r.ranges) mutators.setCopyRanges(store, r.ranges);
      else mutators.setCopyRange(store, r.range);
      payloadRevision = store.getState().ui.copyRevision ?? 0;
      await writeClipboardText(snapshotText ?? r.tsv);
      return;
    }
    if (kind === 'cut') {
      if (isRestricted()) return;
      const r = cut(s, wb);
      if (!r) {
        snapshot = null;
        snapshotText = null;
        payloadRevision = null;
        return;
      }
      snapshot = captureMaterializedSnapshot(s, r, 'cut');
      snapshotText = r.tsv;
      mutators.setCopyRange(store, r.range, 'cut');
      payloadRevision = store.getState().ui.copyRevision ?? 0;
      await writeClipboardText(r.tsv);
      return;
    }
    // paste
    const requestedRevision = s.ui.copyRevision ?? 0;
    let text = '';
    try {
      if (!navigator.clipboard?.readText) {
        const internal = activeInternalText();
        if (internal === null) return;
        text = internal;
      } else {
        text = (await navigator.clipboard.readText()) ?? '';
      }
    } catch (err) {
      if ((store.getState().ui.copyRevision ?? 0) !== requestedRevision) return;
      const internal = activeInternalText();
      if (internal === null) {
        console.warn('formulon-cell: clipboard read failed', err);
        return;
      }
      // Browsers may deny system clipboard reads even for a copy made in
      // this workbook. Its live marquee identifies the available local copy.
      text = internal;
    }
    if ((store.getState().ui.copyRevision ?? 0) !== requestedRevision) return;
    if (disposed) return;
    const pasteState = store.getState();
    if (pasteState.ui.editor.kind !== 'idle') return;
    if (!hasPastePayload(text)) return;
    if (isRestricted()) {
      const r = restrictedPaste(pasteState, text);
      if (!r) return;
      consumeCutMarquee();
      mutators.setRange(store, r.writtenRange);
      deps.onAfterCommit();
      return;
    }
    if (!isWithinNavigationBounds(pasteDestinationRange(pasteState, text))) return;
    if (history) history.begin();
    let r: { writtenRange: Range } | null = null;
    let activation: PasteOptionsActivation | null = null;
    try {
      ({ result: r, activation } = pasteFromClipboardText(pasteState, text));
    } finally {
      if (history) history.end();
    }
    if (r) {
      consumeCutMarquee();
      mutators.setRange(store, r.writtenRange);
      deps.onAfterCommit();
      if (activation) deps.onPasteOptions?.(activation);
    }
  };

  return {
    getSnapshot: () => {
      const state = store.getState();
      return snapshot && hasLiveInternalPayload(state) ? snapshot : null;
    },
    runShortcut,
    detach() {
      disposed = true;
      host.removeEventListener('copy', onCopy);
      host.removeEventListener('cut', onCut);
      host.removeEventListener('paste', onPaste);
    },
  };
}
