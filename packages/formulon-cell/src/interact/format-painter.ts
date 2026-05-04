import { type History, recordFormatChange } from '../commands/history.js';
import type { Range } from '../engine/types.js';
import { addrKey } from '../engine/workbook-handle.js';
import { hitTest } from '../render/geometry.js';
import { type CellFormat, type SpreadsheetStore, mutators } from '../store/store.js';

export interface FormatPainterDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  /** Optional shared history. When provided, paint operations push a single
   *  format-snapshot entry per paste so Cmd+Z reverts the painted formats. */
  history?: History | null;
}

export interface FormatPainterHandle {
  /** Capture the current selection's format snapshot and arm paint mode.
   *  When `sticky` is true, paint repeats until `deactivate()` or Esc. */
  activate(sticky?: boolean): void;
  deactivate(): void;
  isActive(): boolean;
  /** Subscribe to active/sticky state transitions. Returns an unsubscribe fn. */
  subscribe(cb: (active: boolean, sticky: boolean) => void): () => void;
  detach(): void;
}

interface Snapshot {
  range: Range;
  // 2D (rows × cols) of CellFormat — undefined entries mean "no format on source cell".
  pattern: (CellFormat | undefined)[][];
}

/**
 * Excel-style "Format Painter" interaction.
 *
 * Click to copy the current selection's formatting to a single destination
 * (single-cell click pastes the whole source pattern; drag-pastes onto a
 * larger destination by tiling). Double-click to enable sticky mode — the
 * pattern repeats until Esc / deactivate.
 *
 * The source snapshot is captured at activation time, not at paint time,
 * which mirrors Excel: editing the source mid-paint does not affect the
 * pattern about to be pasted.
 */
export function attachFormatPainter(deps: FormatPainterDeps): FormatPainterHandle {
  const { host, store } = deps;
  const history = deps.history ?? null;

  let snapshot: Snapshot | null = null;
  let sticky = false;
  let dragging = false;
  let dragStart: { row: number; col: number } | null = null;

  const HOST_CLASS = 'fc-host--paintbrush';
  const listeners = new Set<(active: boolean, sticky: boolean) => void>();

  const isActive = (): boolean => snapshot !== null;

  const fire = (): void => {
    const a = isActive();
    for (const cb of listeners) cb(a, sticky);
  };

  const captureFromSelection = (): Snapshot | null => {
    const s = store.getState();
    const r = s.selection.range;
    const sheet = r.sheet;
    const rows = r.r1 - r.r0 + 1;
    const cols = r.c1 - r.c0 + 1;
    if (rows <= 0 || cols <= 0) return null;
    // Excel caps Format Painter source at the visible used range. We accept
    // anything up to ~10000 cells; beyond that the snapshot grows too large.
    if (rows * cols > 100_000) return null;
    const pattern: (CellFormat | undefined)[][] = [];
    for (let dr = 0; dr < rows; dr += 1) {
      const row: (CellFormat | undefined)[] = [];
      for (let dc = 0; dc < cols; dc += 1) {
        const f = s.format.formats.get(addrKey({ sheet, row: r.r0 + dr, col: r.c0 + dc }));
        row.push(f ? { ...f, borders: f.borders ? { ...f.borders } : undefined } : undefined);
      }
      pattern.push(row);
    }
    return { range: r, pattern };
  };

  const activate = (sticky_ = false): void => {
    const cap = captureFromSelection();
    if (!cap) return;
    snapshot = cap;
    sticky = sticky_;
    host.classList.add(HOST_CLASS);
    fire();
  };

  const deactivate = (): void => {
    if (!snapshot) return;
    snapshot = null;
    sticky = false;
    dragging = false;
    dragStart = null;
    host.classList.remove(HOST_CLASS);
    fire();
  };

  /** Tile `snapshot.pattern` onto `dest` starting at its top-left, replacing
   *  any existing format in the destination range. */
  const apply = (dest: Range): void => {
    const snap = snapshot;
    if (!snap) return;
    const { pattern } = snap;
    const sheet = dest.sheet;
    const rows = pattern.length;
    const cols = pattern[0]?.length ?? 0;
    if (!rows || !cols) return;

    // Operate directly on the format map for atomicity. Wrap in history so
    // Cmd+Z reverts the entire paste.
    recordFormatChange(history, store, () => {
      store.setState((s) => {
        const formats = new Map(s.format.formats);
        for (let r = dest.r0; r <= dest.r1; r += 1) {
          for (let c = dest.c0; c <= dest.c1; c += 1) {
            const sr = (r - dest.r0) % rows;
            const sc = (c - dest.c0) % cols;
            const src = pattern[sr]?.[sc];
            const key = addrKey({ sheet, row: r, col: c });
            if (src) {
              // Wholesale replace — Format Painter does not merge.
              formats.set(key, {
                ...src,
                borders: src.borders ? { ...src.borders } : undefined,
              });
            } else {
              formats.delete(key);
            }
          }
        }
        return { ...s, format: { formats } };
      });
    });

    // Move selection to the painted range, matching Excel behavior.
    mutators.setActive(store, { sheet, row: dest.r0, col: dest.c0 });
    if (dest.r0 !== dest.r1 || dest.c0 !== dest.c1) {
      mutators.extendRangeTo(store, { sheet, row: dest.r1, col: dest.c1 });
    }
  };

  const localXY = (e: PointerEvent): { x: number; y: number } => {
    // pointer.ts uses the grid as host; we attach to the same parent (`fc-host`).
    // Walk to the nearest grid surface so coordinates map correctly.
    const grid = host.querySelector('.fc-host__grid') as HTMLElement | null;
    const ref = grid ?? host;
    const rect = ref.getBoundingClientRect();
    return { x: e.clientX - rect.left, y: e.clientY - rect.top };
  };

  const onPointerDown = (e: PointerEvent): void => {
    if (!snapshot || e.button !== 0) return;
    const { x, y } = localXY(e);
    const s = store.getState();
    const cell = hitTest(s.layout, s.viewport, x, y);
    if (!cell) return;
    e.preventDefault();
    e.stopPropagation();
    dragging = true;
    dragStart = cell;
    // Single-click default: paste source-sized chunk anchored at the click cell.
    const rows = snapshot.pattern.length;
    const cols = snapshot.pattern[0]?.length ?? 1;
    mutators.setActive(store, { sheet: s.data.sheetIndex, row: cell.row, col: cell.col });
    mutators.extendRangeTo(store, {
      sheet: s.data.sheetIndex,
      row: cell.row + rows - 1,
      col: cell.col + cols - 1,
    });
    host.setPointerCapture(e.pointerId);
  };

  const onPointerMove = (e: PointerEvent): void => {
    if (!snapshot || !dragging || !dragStart) return;
    const { x, y } = localXY(e);
    const s = store.getState();
    const cell = hitTest(s.layout, s.viewport, x, y);
    if (!cell) return;
    e.preventDefault();
    e.stopPropagation();
    // Drag overrides the auto-sized destination — extend from anchor to current.
    mutators.setActive(store, { sheet: s.data.sheetIndex, row: dragStart.row, col: dragStart.col });
    mutators.extendRangeTo(store, { sheet: s.data.sheetIndex, row: cell.row, col: cell.col });
  };

  const onPointerUp = (e: PointerEvent): void => {
    if (!snapshot || !dragging) return;
    if (host.hasPointerCapture(e.pointerId)) host.releasePointerCapture(e.pointerId);
    dragging = false;
    dragStart = null;
    const s = store.getState();
    e.preventDefault();
    e.stopPropagation();
    apply(s.selection.range);
    if (!sticky) deactivate();
  };

  const onKey = (e: KeyboardEvent): void => {
    if (!snapshot) return;
    if (e.key === 'Escape') {
      e.preventDefault();
      deactivate();
    }
  };

  // Capture-phase listeners so we run before pointer.ts and the editor.
  host.addEventListener('pointerdown', onPointerDown, true);
  host.addEventListener('pointermove', onPointerMove, true);
  host.addEventListener('pointerup', onPointerUp, true);
  host.addEventListener('pointercancel', onPointerUp, true);
  document.addEventListener('keydown', onKey, true);

  return {
    activate,
    deactivate,
    isActive,
    subscribe(cb) {
      listeners.add(cb);
      return () => listeners.delete(cb);
    },
    detach() {
      deactivate();
      listeners.clear();
      host.removeEventListener('pointerdown', onPointerDown, true);
      host.removeEventListener('pointermove', onPointerMove, true);
      host.removeEventListener('pointerup', onPointerUp, true);
      host.removeEventListener('pointercancel', onPointerUp, true);
      document.removeEventListener('keydown', onKey, true);
    },
  };
}
