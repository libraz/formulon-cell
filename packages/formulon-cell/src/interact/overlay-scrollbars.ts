import { gridOriginX, gridOriginY, layoutForView } from '../render/geometry.js';
import type { SpreadsheetStore, State } from '../store/store.js';
import { lastUsedCell } from './data-edge-navigation.js';

/** How long the thumbs stay after the last scroll before they fade out. */
const IDLE_HIDE_MS = 1000;

/** Length and offset of an overlay thumb along a track of `track` px. The
 *  scrollable extent is the used rows or columns, or the view when it reaches
 *  past them, so a short sheet shows a full-length thumb. */
export function overlayThumb(input: {
  track: number;
  start: number;
  visible: number;
  used: number;
}): { length: number; offset: number } {
  const total = Math.max(input.used, input.start + input.visible);
  if (total <= 0) return { length: input.track, offset: 0 };
  return {
    length: (input.track * input.visible) / total,
    offset: (input.track * input.start) / total,
  };
}

export interface OverlayScrollbarsHandle {
  detach(): void;
}

/**
 * The Mac desktop's overlay scroll bars: a vertical and a horizontal thumb
 * over the grid, shown while the view scrolls and faded out when it rests.
 * They only show the position; they take no pointer input.
 */
export function attachOverlayScrollbars(deps: {
  grid: HTMLElement;
  store: SpreadsheetStore;
}): OverlayScrollbarsHandle {
  const { grid, store } = deps;
  const thumb = (axis: 'v' | 'h'): HTMLDivElement => {
    const el = document.createElement('div');
    el.className = `fc-scrollthumb fc-scrollthumb--${axis}`;
    el.setAttribute('aria-hidden', 'true');
    grid.appendChild(el);
    return el;
  };
  const vertical = thumb('v');
  const horizontal = thumb('h');

  let extentKey: readonly unknown[] = [];
  let extent = { row: 0, col: 0 };
  const usedExtent = (s: State): { row: number; col: number } => {
    const key = [s.data.cells, s.format.formats, s.data.sheetIndex];
    if (key.some((part, i) => part !== extentKey[i])) {
      extentKey = key;
      extent = lastUsedCell(s, s.data.sheetIndex);
    }
    return extent;
  };

  const place = (s: State): void => {
    const layout = layoutForView(s);
    const top = gridOriginY(layout);
    const left = gridOriginX(layout);
    const used = usedExtent(s);
    const v = overlayThumb({
      track: Math.max(0, grid.clientHeight - top),
      start: s.viewport.rowStart,
      visible: s.viewport.rowCount,
      used: used.row + 1,
    });
    const h = overlayThumb({
      track: Math.max(0, grid.clientWidth - left),
      start: s.viewport.colStart,
      visible: s.viewport.colCount,
      used: used.col + 1,
    });
    vertical.style.top = `${top + v.offset}px`;
    vertical.style.height = `${v.length}px`;
    horizontal.style.left = `${left + h.offset}px`;
    horizontal.style.width = `${h.length}px`;
  };

  let hideTimer: ReturnType<typeof setTimeout> | null = null;
  const show = (s: State): void => {
    place(s);
    vertical.dataset.visible = 'true';
    horizontal.dataset.visible = 'true';
    if (hideTimer) clearTimeout(hideTimer);
    hideTimer = setTimeout(() => {
      hideTimer = null;
      delete vertical.dataset.visible;
      delete horizontal.dataset.visible;
    }, IDLE_HIDE_MS);
  };

  let last = store.getState().viewport;
  const unsubscribe = store.subscribe(() => {
    const s = store.getState();
    const next = s.viewport;
    if (next.rowStart === last.rowStart && next.colStart === last.colStart) {
      last = next;
      return;
    }
    last = next;
    show(s);
  });

  return {
    detach() {
      unsubscribe();
      if (hideTimer) clearTimeout(hideTimer);
      vertical.remove();
      horizontal.remove();
    },
  };
}
