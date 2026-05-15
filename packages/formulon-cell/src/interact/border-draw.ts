import { type History, recordFormatChange } from '../commands/history.js';
import { addrKey } from '../engine/address.js';
import { cellRect, hitTest } from '../render/geometry.js';
import {
  type CellBorderSide,
  type CellBorderStyle,
  type CellBorders,
  mutators,
  type SpreadsheetStore,
} from '../store/store.js';

export type BorderDrawMode = 'draw' | 'grid' | 'erase';

export interface BorderDrawDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  history?: History | null;
}

export interface BorderDrawHandle {
  /** Enter draw mode. While armed, pointer events on the grid edit cell
   *  borders instead of selecting cells. */
  activate(mode: BorderDrawMode, style?: CellBorderStyle, color?: string): void;
  deactivate(): void;
  isActive(): boolean;
  getMode(): BorderDrawMode | null;
  /** Update the brush style/color without re-arming. No-op when inactive. */
  setStyle(style: CellBorderStyle): void;
  setColor(color: string | undefined): void;
  subscribe(cb: (mode: BorderDrawMode | null) => void): () => void;
  detach(): void;
}

type Edge = 'top' | 'right' | 'bottom' | 'left';

interface EdgeHit {
  row: number;
  col: number;
  edge: Edge;
}

const EDGE_TOL = 6;
const HOST_CLASS_DRAW = 'fc-host--border-draw';
const HOST_CLASS_GRID = 'fc-host--border-grid';
const HOST_CLASS_ERASE = 'fc-host--border-erase';

/**
 * Excel-parity border drawing modes. Activated from the borders dropdown:
 *
 * - `draw`: click a cell edge to toggle a single side; drag to paint the
 *   nearest edge under the pointer continuously.
 * - `grid`: drag to select a rectangle, then apply all-borders to it on
 *   pointerup.
 * - `erase`: click/drag to clear the side under the pointer.
 *
 * The host class is set so CSS can swap the cursor. Escape (or re-clicking
 * the menu entry) exits.
 */
export function attachBorderDraw(deps: BorderDrawDeps): BorderDrawHandle {
  const { host, store } = deps;
  const history = deps.history ?? null;

  let mode: BorderDrawMode | null = null;
  let style: CellBorderStyle = 'thin';
  let color: string | undefined;
  let dragging = false;
  let dragAnchor: { row: number; col: number } | null = null;
  let lastEdgeKey: string | null = null;
  const listeners = new Set<(m: BorderDrawMode | null) => void>();

  const fire = (): void => {
    for (const cb of listeners) cb(mode);
  };

  const setHostClass = (next: BorderDrawMode | null): void => {
    host.classList.remove(HOST_CLASS_DRAW, HOST_CLASS_GRID, HOST_CLASS_ERASE);
    if (next === 'draw') host.classList.add(HOST_CLASS_DRAW);
    else if (next === 'grid') host.classList.add(HOST_CLASS_GRID);
    else if (next === 'erase') host.classList.add(HOST_CLASS_ERASE);
  };

  const activate = (
    next: BorderDrawMode,
    nextStyle?: CellBorderStyle,
    nextColor?: string,
  ): void => {
    mode = next;
    if (nextStyle) style = nextStyle;
    if (nextColor !== undefined) color = nextColor;
    setHostClass(mode);
    fire();
  };

  const deactivate = (): void => {
    if (mode === null) return;
    mode = null;
    dragging = false;
    dragAnchor = null;
    lastEdgeKey = null;
    setHostClass(null);
    fire();
  };

  const localXY = (e: PointerEvent): { x: number; y: number } => {
    const grid = host.querySelector('.fc-host__grid') as HTMLElement | null;
    const ref = grid ?? host;
    const rect = ref.getBoundingClientRect();
    return { x: e.clientX - rect.left, y: e.clientY - rect.top };
  };

  /** Pick the nearest cell edge for a pointer in the data area. Returns null
   *  when the pointer is inside a cell with no edge within EDGE_TOL. */
  const edgeAt = (x: number, y: number): EdgeHit | null => {
    const s = store.getState();
    const cell = hitTest(s.layout, s.viewport, x, y);
    if (!cell) return null;
    const r = cellRect(s.layout, s.viewport, cell.row, cell.col);
    const dx0 = x - r.x;
    const dx1 = r.x + r.w - x;
    const dy0 = y - r.y;
    const dy1 = r.y + r.h - y;
    const min = Math.min(dx0, dx1, dy0, dy1);
    if (min > EDGE_TOL) return null;
    if (min === dy0) return { row: cell.row, col: cell.col, edge: 'top' };
    if (min === dy1) return { row: cell.row, col: cell.col, edge: 'bottom' };
    if (min === dx0) return { row: cell.row, col: cell.col, edge: 'left' };
    return { row: cell.row, col: cell.col, edge: 'right' };
  };

  const sideValue = (): CellBorderSide => (color !== undefined ? { style, color } : { style });

  /** Apply a single-edge change to one cell, merging with any existing borders. */
  const writeEdge = (hit: EdgeHit, erase: boolean): void => {
    const s = store.getState();
    const sheet = s.data.sheetIndex;
    const addr = { sheet, row: hit.row, col: hit.col };
    const cur = s.format.formats.get(addrKey(addr))?.borders ?? {};
    const next: CellBorders = { ...cur };
    if (erase) {
      next[hit.edge] = false;
    } else {
      next[hit.edge] = sideValue();
    }
    mutators.setCellFormat(store, addr, { borders: next });
  };

  /** Apply all-borders to a rectangular range. Used by the grid sub-mode. */
  const writeGrid = (
    anchor: { row: number; col: number },
    current: { row: number; col: number },
  ): void => {
    const s = store.getState();
    const sheet = s.data.sheetIndex;
    const r0 = Math.min(anchor.row, current.row);
    const r1 = Math.max(anchor.row, current.row);
    const c0 = Math.min(anchor.col, current.col);
    const c1 = Math.max(anchor.col, current.col);
    const side = sideValue();
    for (let r = r0; r <= r1; r += 1) {
      for (let c = c0; c <= c1; c += 1) {
        mutators.setCellFormat(
          store,
          { sheet, row: r, col: c },
          { borders: { top: side, right: side, bottom: side, left: side } },
        );
      }
    }
  };

  const onPointerDown = (e: PointerEvent): void => {
    if (mode === null || e.button !== 0) return;
    const { x, y } = localXY(e);
    if (mode === 'grid') {
      const s = store.getState();
      const cell = hitTest(s.layout, s.viewport, x, y);
      if (!cell) return;
      e.preventDefault();
      e.stopPropagation();
      dragging = true;
      dragAnchor = cell;
      host.setPointerCapture(e.pointerId);
      return;
    }
    const hit = edgeAt(x, y);
    if (!hit) return;
    e.preventDefault();
    e.stopPropagation();
    dragging = true;
    lastEdgeKey = `${hit.row}:${hit.col}:${hit.edge}`;
    recordFormatChange(history, store, () => writeEdge(hit, mode === 'erase'));
    host.setPointerCapture(e.pointerId);
  };

  const onPointerMove = (e: PointerEvent): void => {
    if (mode === null || !dragging) return;
    const { x, y } = localXY(e);
    if (mode === 'grid') {
      // Selection rectangle is finalized on pointerup; mid-drag we don't
      // mutate. Could preview here but Excel doesn't preview either.
      e.preventDefault();
      return;
    }
    const hit = edgeAt(x, y);
    if (!hit) return;
    const key = `${hit.row}:${hit.col}:${hit.edge}`;
    if (key === lastEdgeKey) return;
    e.preventDefault();
    e.stopPropagation();
    lastEdgeKey = key;
    recordFormatChange(history, store, () => writeEdge(hit, mode === 'erase'));
  };

  const onPointerUp = (e: PointerEvent): void => {
    if (mode === null || !dragging) return;
    if (host.hasPointerCapture(e.pointerId)) host.releasePointerCapture(e.pointerId);
    if (mode === 'grid' && dragAnchor) {
      const { x, y } = localXY(e);
      const s = store.getState();
      const cell = hitTest(s.layout, s.viewport, x, y) ?? dragAnchor;
      e.preventDefault();
      e.stopPropagation();
      const anchor = dragAnchor;
      recordFormatChange(history, store, () => writeGrid(anchor, cell));
    }
    dragging = false;
    dragAnchor = null;
    lastEdgeKey = null;
  };

  const onKey = (e: KeyboardEvent): void => {
    if (mode === null) return;
    if (e.key === 'Escape') {
      e.preventDefault();
      deactivate();
    }
  };

  // Capture-phase so we run before pointer.ts / editor / format-painter.
  host.addEventListener('pointerdown', onPointerDown, true);
  host.addEventListener('pointermove', onPointerMove, true);
  host.addEventListener('pointerup', onPointerUp, true);
  host.addEventListener('pointercancel', onPointerUp, true);
  document.addEventListener('keydown', onKey, true);

  return {
    activate,
    deactivate,
    isActive: () => mode !== null,
    getMode: () => mode,
    setStyle(next) {
      style = next;
    },
    setColor(next) {
      color = next;
    },
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
