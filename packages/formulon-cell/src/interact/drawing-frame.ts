import { rotatePoint } from '../engine/drawing-geometry.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import type { ChartSourceHighlight, SelectedObject, State } from '../store/types.js';
import { focusSheet } from './dialog-focus.js';

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

const TEXT_ENTRY = 'input, textarea, select, [contenteditable="true"]';
const MODIFIER_KEYS = new Set(['Alt', 'Control', 'Meta', 'Shift']);
/** The ribbon and its menus, which act on the selected object rather than leave it. */
const RIBBON_SURFACE = '.fc-tb__ribbon-shell, .fc-tb__menu, .fc-tb__submenu, .fc-tb__rb-dd';

/** Whether a press on, or focus moving to, `target` keeps a selected object selected. */
export const keepsObjectSelected = (target: EventTarget | null): boolean =>
  target instanceof Element && target.closest(RIBBON_SURFACE) !== null;

/**
 * Deselect on a press or focus move outside `root`, `menu`, the ribbon and any
 * `extraInside` match, and keep keys pressed in `root` from the grid. Returns
 * the disposer.
 */
export const bindObjectDeselect = (
  root: HTMLElement,
  menu: { contains(target: Node): boolean },
  deselect: () => void,
  extraInside?: string,
): (() => void) => {
  const inside = (target: EventTarget | null): boolean =>
    target instanceof Node &&
    (root.contains(target) ||
      menu.contains(target) ||
      keepsObjectSelected(target) ||
      (extraInside !== undefined &&
        target instanceof Element &&
        target.closest(extraInside) !== null));
  // Like Excel, a press anywhere else on the sheet leaves the object.
  const onDocumentPointerDown = (e: PointerEvent): void => {
    if (!inside(e.target)) deselect();
  };
  const onFocusOut = (e: FocusEvent): void => {
    if (e.relatedTarget !== null && !inside(e.relatedTarget)) deselect();
  };
  document.addEventListener('pointerdown', onDocumentPointerDown, true);
  root.addEventListener('focusout', onFocusOut);
  root.addEventListener('keydown', keepKeyFromGrid);
  return () => {
    document.removeEventListener('pointerdown', onDocumentPointerDown, true);
    root.removeEventListener('focusout', onFocusOut);
    root.removeEventListener('keydown', keepKeyFromGrid);
  };
};

/**
 * Keep a key pressed on a selected object from the grid under it, which would
 * otherwise type into or move the active cell. Shortcuts the sheet owns still
 * reach it: Undo, and Copy, Cut and Paste, which act on the selected object.
 */
export const keepKeyFromGrid = (event: KeyboardEvent): void => {
  if (MODIFIER_KEYS.has(event.key)) return;
  if (event.target instanceof Element && event.target.closest(TEXT_ENTRY)) return;
  if ((event.ctrlKey || event.metaKey) && !isDuplicateShortcut(event)) return;
  event.stopPropagation();
};

/** Ctrl+D, or Cmd+D on a Mac: duplicate the selected object. */
export const isDuplicateShortcut = (event: KeyboardEvent): boolean =>
  (event.ctrlKey || event.metaKey) &&
  !event.altKey &&
  !event.shiftKey &&
  event.key.toLowerCase() === 'd';

/** The `data-*` id of the panel in `root` holding focus, before `root` is
 *  rebuilt; undefined when focus is outside `root`. */
export const focusedPanelId = (
  root: HTMLElement,
  selector: string,
  key: string,
): string | undefined => {
  const active = document.activeElement;
  if (!(active instanceof HTMLElement) || !root.contains(active)) return undefined;
  return active.closest<HTMLElement>(selector)?.dataset[key] ?? '';
};

/** Focus the panel of object `id` in a rebuilt `root`, or `host` when the
 *  object is gone, so the keyboard stays on the sheet. */
export const focusPanel = (
  root: HTMLElement,
  selector: string,
  key: string,
  id: string,
  host: HTMLElement,
): void => {
  const panel = Array.from(root.querySelectorAll<HTMLElement>(selector)).find(
    (candidate) => candidate.dataset[key] === id,
  );
  if (panel) panel.focus({ preventScroll: true });
  else focusSheet(host, { preventScroll: true });
};

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

/** Like {@link resizedFrame} for a frame turned by `rotationDeg` about its centre,
 *  with (dx, dy) in screen pixels: the delta is taken into the frame's own axes,
 *  and the box is moved so the edge or corner opposite `handle` stays put on screen. */
export const rotatedResizedFrame = (
  start: DrawingFrame,
  handle: DrawingHandle,
  dx: number,
  dy: number,
  keepRatio: boolean,
  min: { w: number; h: number },
  rotationDeg: number,
): DrawingFrame => {
  if (rotationDeg === 0) return resizedFrame(start, handle, dx, dy, keepRatio, min);
  const [lx, ly] = rotatePoint(dx, dy, -rotationDeg);
  const next = resizedFrame(start, handle, lx, ly, keepRatio, min);
  // The fixed point as a fraction of the box: the side opposite the dragged one, else the middle.
  const fx = handle.endsWith('w') ? 1 : handle.endsWith('e') ? 0 : 0.5;
  const fy = handle.startsWith('n') ? 1 : handle.startsWith('s') ? 0 : 0.5;
  const onScreen = (frame: DrawingFrame): [number, number] => {
    const [ox, oy] = rotatePoint(frame.w * (fx - 0.5), frame.h * (fy - 0.5), rotationDeg);
    return [frame.x + frame.w / 2 + ox, frame.y + frame.h / 2 + oy];
  };
  const [bx, by] = onScreen(start);
  const [ax, ay] = onScreen(next);
  return { ...next, x: next.x + bx - ax, y: next.y + by - ay };
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

/** Append the rotation handle drawn above the top middle sizing handle. */
export const appendRotationHandle = (panel: HTMLElement): void => {
  const grip = document.createElement('div');
  grip.className = 'fc-drawing-rotate';
  grip.setAttribute('aria-hidden', 'true');
  panel.appendChild(grip);
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
    /** The frame's rotation in degrees; a handle then drags along the frame's own axes. */
    rotationDeg?: number;
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
      const angle = options.rotationDeg ?? 0;
      let frame = rotatedResizedFrame(
        start,
        handle,
        dx,
        dy,
        options.keepRatio(move),
        options.min,
        angle,
      );
      // Gridlines are screen-aligned, so a turned frame does not snap.
      if (snap && angle === 0) frame = snappedFrame(frame, handle, snap, options.min);
      latest = {
        ...(frame.x === start.x ? {} : { x: frame.x }),
        ...(frame.y === start.y ? {} : { y: frame.y }),
        w: frame.w,
        h: frame.h,
      };
    }
    options.preview(latest);
  };
  const stop = (): void => {
    window.removeEventListener('pointermove', onMove);
    window.removeEventListener('pointerup', onUp);
    window.removeEventListener('pointercancel', onCancel);
  };
  const onUp = (): void => {
    stop();
    if (latest) options.commit(latest);
  };
  const onCancel = (): void => stop();
  window.addEventListener('pointermove', onMove);
  window.addEventListener('pointerup', onUp);
  window.addEventListener('pointercancel', onCancel);
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

/**
 * Re-render an object overlay on store changes, skipping its own selection
 * publishes and notifications that only moved the cell selection, which an
 * overlay's render never reads.
 */
export const subscribeObjectOverlay = (
  store: SpreadsheetStore,
  selection: ObjectSelectionPublisher,
  render: () => void,
): (() => void) => {
  let last = store.getState();
  return store.subscribe(() => {
    const prev = last;
    last = store.getState();
    if (selection.publishing) return;
    const changed = (Object.keys(last) as (keyof State)[]).some(
      (key) => key !== 'selection' && last[key] !== prev[key],
    );
    if (changed) render();
  });
};
