import { mutators, type SpreadsheetStore } from '../store/store.js';
import type { ChartSourceHighlight, SelectedObject } from '../store/types.js';

/** A drawing object's frame in grid pixels. */
export interface DrawingFrame {
  x: number;
  y: number;
  w: number;
  h: number;
}

/** The sizing handles of a selected chart or picture, by compass point. */
export const DRAWING_HANDLES = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'] as const;
export type DrawingHandle = (typeof DRAWING_HANDLES)[number];

/** Pointer travel before a press on an object becomes a move. */
const DRAG_THRESHOLD = 3;

/** Elements that keep a press to themselves instead of starting a move. */
export const INTERACTIVE = 'button, input, select, textarea, option, [contenteditable="true"]';

/** The frame `start` takes when `handle` is dragged by (dx, dy); a corner
 *  dragged with `keepRatio` scales the frame without changing its shape. */
export const resizedFrame = (
  start: DrawingFrame,
  handle: DrawingHandle,
  dx: number,
  dy: number,
  keepRatio: boolean,
  min: { w: number; h: number },
): DrawingFrame => {
  const west = handle.endsWith('w');
  const east = handle.endsWith('e');
  const north = handle.startsWith('n');
  const south = handle.startsWith('s');
  let w = Math.max(min.w, east ? start.w + dx : west ? start.w - dx : start.w);
  let h = Math.max(min.h, south ? start.h + dy : north ? start.h - dy : start.h);
  if (keepRatio && (east || west) && (north || south)) {
    const scale = Math.max(w / start.w, h / start.h);
    w = Math.max(min.w, start.w * scale);
    h = Math.max(min.h, start.h * scale);
  }
  return {
    x: west ? start.x + start.w - w : start.x,
    y: north ? start.y + start.h - h : start.y,
    w,
    h,
  };
};

/** `frame` with the edges `handle` drags moved to their nearest gridlines. */
const snappedFrame = (
  frame: DrawingFrame,
  handle: DrawingHandle,
  snap: (axis: 'x' | 'y', value: number) => number,
  min: { w: number; h: number },
): DrawingFrame => {
  let { x, y, w, h } = frame;
  if (handle.endsWith('e')) w = Math.max(min.w, snap('x', x + w) - x);
  if (handle.endsWith('w')) {
    const right = x + w;
    x = Math.min(snap('x', x), right - min.w);
    w = right - x;
  }
  if (handle.startsWith('s')) h = Math.max(min.h, snap('y', y + h) - y);
  if (handle.startsWith('n')) {
    const bottom = y + h;
    y = Math.min(snap('y', y), bottom - min.h);
    h = bottom - y;
  }
  return { x, y, w, h };
};

/** Append the eight sizing handles to `panel`; the stylesheet shows them
 *  only while the object is selected. */
export const appendDrawingHandles = (
  panel: HTMLElement,
  onPointerDown: (e: PointerEvent, handle: DrawingHandle) => void,
): void => {
  for (const handle of DRAWING_HANDLES) {
    const grip = document.createElement('div');
    grip.className = 'fc-drawing-handle';
    grip.dataset.handle = handle;
    grip.setAttribute('aria-hidden', 'true');
    grip.addEventListener('pointerdown', (e) => {
      e.stopPropagation();
      onPointerDown(e, handle);
    });
    panel.appendChild(grip);
  }
};

/**
 * Drag an object's frame from `start`: the whole frame when `handle` is
 * omitted, else one sizing handle. `preview` follows the pointer once it
 * passes the drag threshold, and `commit` runs once on release with the last
 * frame, so a drag is one change to the workbook.
 */
export const trackFrameDrag = (
  e: PointerEvent,
  options: {
    start: DrawingFrame;
    handle?: DrawingHandle;
    min: { w: number; h: number };
    /** Whether a corner keeps the frame's shape for this pointer state. */
    keepRatio: (move: PointerEvent) => boolean;
    /** The gridline nearest a value; with Alt held the dragged edges snap to it. */
    snap?: (axis: 'x' | 'y', value: number) => number;
    preview: (patch: Partial<DrawingFrame>) => void;
    commit: (patch: Partial<DrawingFrame>) => void;
  },
): void => {
  const { start, handle } = options;
  const startX = e.clientX;
  const startY = e.clientY;
  let latest: Partial<DrawingFrame> | null = null;
  const onMove = (move: PointerEvent): void => {
    const dx = move.clientX - startX;
    const dy = move.clientY - startY;
    if (!latest && Math.hypot(dx, dy) < DRAG_THRESHOLD) return;
    move.preventDefault();
    const snap = move.altKey ? options.snap : undefined;
    if (!handle) {
      const x = start.x + dx;
      const y = start.y + dy;
      // Like the spreadsheet, Alt docks the object's top-left corner to a cell.
      latest = snap ? { x: snap('x', x), y: snap('y', y) } : { x, y };
    } else {
      let frame = resizedFrame(start, handle, dx, dy, options.keepRatio(move), options.min);
      if (snap) frame = snappedFrame(frame, handle, snap, options.min);
      latest = {
        ...(frame.x === start.x ? {} : { x: frame.x }),
        ...(frame.y === start.y ? {} : { y: frame.y }),
        w: frame.w,
        h: frame.h,
      };
    }
    options.preview(latest);
  };
  const onUp = (): void => {
    window.removeEventListener('pointermove', onMove);
    window.removeEventListener('pointerup', onUp);
    if (latest) options.commit(latest);
  };
  window.addEventListener('pointermove', onMove);
  window.addEventListener('pointerup', onUp, { once: true });
};

/** Move `panel` to a previewed frame without touching the store. */
export const previewFrame = (panel: HTMLElement, patch: Partial<DrawingFrame>): void => {
  if (patch.x !== undefined) panel.style.left = `${patch.x}px`;
  if (patch.y !== undefined) panel.style.top = `${patch.y}px`;
  if (patch.w !== undefined) panel.style.width = `${patch.w}px`;
  if (patch.h !== undefined) panel.style.height = `${patch.h}px`;
};

/**
 * One overlay's share of the sheet's selected object. Each overlay publishes
 * its own selection and clears only what it published, so a chart overlay
 * leaving its chart does not drop a picture another overlay selected.
 */
export interface ObjectSelectionPublisher {
  publish(object: SelectedObject | null, sourceRefs?: ChartSourceHighlight[]): void;
  /** True while a publish is updating the store, so the overlay can skip
   *  re-rendering for a change it made itself. */
  readonly publishing: boolean;
}

export const createObjectSelectionPublisher = (
  store: SpreadsheetStore,
): ObjectSelectionPublisher => {
  let publishedId: string | null = null;
  let publishing = false;
  return {
    get publishing() {
      return publishing;
    },
    publish(object, sourceRefs = []) {
      const { ui } = store.getState();
      const current = ui.selectedObject;
      if (!object && current?.id !== publishedId) return;
      if (
        current?.id === object?.id &&
        current?.name === object?.name &&
        sameHighlights(ui.chartSourceRefs, sourceRefs)
      ) {
        return;
      }
      publishedId = object?.id ?? null;
      publishing = true;
      try {
        mutators.setSelectedObject(store, object, sourceRefs);
      } finally {
        publishing = false;
      }
    },
  };
};

const sameHighlights = (
  a: readonly ChartSourceHighlight[],
  b: readonly ChartSourceHighlight[],
): boolean =>
  a.length === b.length &&
  a.every(
    (ref, i) =>
      ref.sheet === b[i]?.sheet &&
      ref.r0 === b[i]?.r0 &&
      ref.c0 === b[i]?.c0 &&
      ref.r1 === b[i]?.r1 &&
      ref.c1 === b[i]?.c1 &&
      ref.role === b[i]?.role,
  );
