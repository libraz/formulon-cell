import { gridOriginX, gridOriginY, layoutForView } from '../render/geometry.js';
import { mutators, type SpreadsheetStore, type State } from '../store/store.js';
import { lastUsedCell } from './data-edge-navigation.js';

/** How long the thumbs stay after the last scroll before they fade out. */
const IDLE_HIDE_MS = 1000;

/** Scrollable extent in rows or columns: the used count, or the view when it reaches past it. */
const thumbTotal = (used: number, start: number, visible: number): number =>
  Math.max(used, start + visible);

/** Length and offset of an overlay thumb along a track of `track` px. The
 *  scrollable extent is the used rows or columns, or the view when it reaches
 *  past them, so a short sheet shows a full-length thumb. */
export function overlayThumb(input: {
  track: number;
  start: number;
  visible: number;
  used: number;
}): { length: number; offset: number } {
  const total = thumbTotal(input.used, input.start, input.visible);
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
 * While shown, a click on the track beside a thumb scrolls one screenful
 * (visible count - 1) toward the click and a thumb drag scrolls in proportion.
 * Hidden bars take no pointer input, so clicks reach the cells beneath.
 */
export function attachOverlayScrollbars(deps: {
  grid: HTMLElement;
  store: SpreadsheetStore;
}): OverlayScrollbarsHandle {
  const { grid, store } = deps;
  const bar = (kind: 'thumb' | 'track', axis: 'v' | 'h'): HTMLDivElement => {
    const el = document.createElement('div');
    el.className = `fc-scroll${kind} fc-scroll${kind}--${axis}`;
    el.setAttribute('aria-hidden', 'true');
    grid.appendChild(el);
    return el;
  };
  const verticalTrack = bar('track', 'v');
  const horizontalTrack = bar('track', 'h');
  const vertical = bar('thumb', 'v');
  const horizontal = bar('thumb', 'h');

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

  /** Offset of each thumb along its axis, in px from the grid edge, as last placed. */
  const thumbStart = { v: 0, h: 0 };

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
    thumbStart.v = top + v.offset;
    thumbStart.h = left + h.offset;
    vertical.style.top = `${top + v.offset}px`;
    vertical.style.height = `${v.length}px`;
    horizontal.style.left = `${left + h.offset}px`;
    horizontal.style.width = `${h.length}px`;
    verticalTrack.style.top = `${top}px`;
    horizontalTrack.style.left = `${left}px`;
  };

  let hideTimer: ReturnType<typeof setTimeout> | null = null;
  let dragging = false;
  const hide = (): void => {
    hideTimer = null;
    for (const el of [vertical, horizontal, verticalTrack, horizontalTrack]) {
      delete el.dataset.visible;
    }
  };
  const armHide = (): void => {
    if (hideTimer) clearTimeout(hideTimer);
    hideTimer = dragging ? null : setTimeout(hide, IDLE_HIDE_MS);
  };
  const show = (s: State): void => {
    place(s);
    for (const el of [vertical, horizontal, verticalTrack, horizontalTrack]) {
      el.dataset.visible = 'true';
    }
    armHide();
  };

  const axisOf = (el: HTMLElement): 'v' | 'h' =>
    el.classList.contains('fc-scroll--v') ? 'v' : 'h';
  const pos = (e: PointerEvent, axis: 'v' | 'h'): number => (axis === 'v' ? e.clientY : e.clientX);
  const thumbStartOnScreen = (axis: 'v' | 'h'): number => {
    const box = grid.getBoundingClientRect();
    return (axis === 'v' ? box.top : box.left) + thumbStart[axis];
  };

  const onTrackDown = (e: PointerEvent): void => {
    const el = e.currentTarget as HTMLElement;
    const axis = axisOf(el);
    if (el.dataset.visible !== 'true' || e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    const s = store.getState();
    const page = Math.max(1, (axis === 'v' ? s.viewport.rowCount : s.viewport.colCount) - 1);
    const dir = pos(e, axis) < thumbStartOnScreen(axis) ? -1 : 1;
    if (axis === 'v') mutators.scrollBy(store, dir * page, 0);
    else mutators.scrollBy(store, 0, dir * page);
  };

  const onThumbDown = (e: PointerEvent): void => {
    const el = e.currentTarget as HTMLElement;
    const axis = axisOf(el);
    if (el.dataset.visible !== 'true' || e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    const s = store.getState();
    const layout = layoutForView(s);
    const used = usedExtent(s);
    const v = axis === 'v';
    const start0 = v ? s.viewport.rowStart : s.viewport.colStart;
    const visible = v ? s.viewport.rowCount : s.viewport.colCount;
    const total = thumbTotal((v ? used.row : used.col) + 1, start0, visible);
    const origin = v ? gridOriginY(layout) : gridOriginX(layout);
    const trackLen = Math.max(1, (v ? grid.clientHeight : grid.clientWidth) - origin);
    const from = pos(e, axis);
    dragging = true;
    armHide();
    try {
      el.setPointerCapture(e.pointerId);
    } catch {
      /* synthetic pointers have no capture */
    }
    const onMove = (m: PointerEvent): void => {
      const target = start0 + Math.round(((pos(m, axis) - from) * total) / trackLen);
      const cur = v ? store.getState().viewport.rowStart : store.getState().viewport.colStart;
      if (v) mutators.scrollBy(store, target - cur, 0);
      else mutators.scrollBy(store, 0, target - cur);
    };
    const onUp = (): void => {
      el.removeEventListener('pointermove', onMove);
      el.removeEventListener('pointerup', onUp);
      el.removeEventListener('pointercancel', onUp);
      dragging = false;
      armHide();
    };
    el.addEventListener('pointermove', onMove);
    el.addEventListener('pointerup', onUp);
    el.addEventListener('pointercancel', onUp);
  };

  for (const el of [vertical, verticalTrack]) el.classList.add('fc-scroll--v');
  for (const el of [horizontal, horizontalTrack]) el.classList.add('fc-scroll--h');
  verticalTrack.addEventListener('pointerdown', onTrackDown);
  horizontalTrack.addEventListener('pointerdown', onTrackDown);
  vertical.addEventListener('pointerdown', onThumbDown);
  horizontal.addEventListener('pointerdown', onThumbDown);

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
      verticalTrack.remove();
      horizontalTrack.remove();
    },
  };
}
