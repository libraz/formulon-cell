import type { History } from '../commands/history.js';
import type { IllustrationMutationAuthorization } from '../commands/session-illustration.js';
import {
  arrangeSessionIllustration,
  canDeleteIllustration,
  clearSessionIllustration,
  duplicateSessionIllustration,
  updateSessionIllustration,
} from '../commands/session-illustration.js';
import {
  shapeCornerRadius,
  shapeDashStroke,
  shapeFillColor,
  shapeOutlineColor,
  shapeOutlineDash,
  shapeOutlineWeight,
} from '../commands/shape-format.js';
import {
  drawingClipPath,
  nearestGridline,
  sessionPosition,
  viewportZoom,
} from '../engine/drawing-geometry.js';
import { DEFAULT_ILLUSTRATION_FRAME } from '../engine/shapes-sync.js';
import { argbFromCssColor } from '../engine/theme-color.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SessionIllustration, SpreadsheetStore } from '../store/store.js';
import type { SessionShapeKind, State } from '../store/types.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import {
  appendDrawingHandles,
  createObjectSelectionPublisher,
  type DrawingFrame,
  type DrawingHandle,
  focusedPanelId,
  focusPanel,
  isDuplicateShortcut,
  keepKeyFromGrid,
  keepsObjectSelected,
  previewFrame,
  trackFrameDrag,
} from './drawing-frame.js';
import { createObjectMenu } from './object-menu.js';

export interface SessionIllustrationsHandle {
  refresh(): void;
  setLabels(next: Partial<SessionIllustrationLabels>): void;
  detach(): void;
}

export interface SessionIllustrationLabels {
  /** How assistive technology names a picture and a shape. */
  picture: string;
  shape: string;
  /** Name box texts for a selected picture or shape, as Excel names them;
   *  `{n}` is the object's number among its kind on the sheet. */
  pictureName: string;
  rectangleName: string;
  roundedRectangleName: string;
  ovalName: string;
  triangleName: string;
  diamondName: string;
  lineName: string;
  arrowName: string;
  /** The object's right-click menu. */
  bringToFront: string;
  sendToBack: string;
}

const DEFAULT_LABELS: SessionIllustrationLabels = {
  picture: 'Picture',
  shape: 'Shape',
  pictureName: 'Picture {n}',
  rectangleName: 'Rectangle {n}',
  roundedRectangleName: 'Rectangle: Rounded Corners {n}',
  ovalName: 'Oval {n}',
  triangleName: 'Isosceles Triangle {n}',
  diamondName: 'Diamond {n}',
  lineName: 'Straight Connector {n}',
  arrowName: 'Straight Arrow Connector {n}',
  bringToFront: 'Bring to Front',
  sendToBack: 'Send to Back',
};

const SHAPE_NAME_LABEL: Record<SessionShapeKind, keyof SessionIllustrationLabels> = {
  rectangle: 'rectangleName',
  'rounded-rectangle': 'roundedRectangleName',
  oval: 'ovalName',
  triangle: 'triangleName',
  diamond: 'diamondName',
  line: 'lineName',
  arrow: 'arrowName',
};

/** Excel's name for a picture or shape, numbered among its kind on the sheet. */
const objectName = (
  state: State,
  item: SessionIllustration,
  labels: SessionIllustrationLabels,
): string => {
  const sameKind = (entry: SessionIllustration): boolean =>
    entry.sheet === item.sheet && entry.kind === item.kind && entry.shape === item.shape;
  const index = state.illustrations.illustrations.filter(sameKind).indexOf(item);
  const template =
    item.kind === 'image'
      ? labels.pictureName
      : labels[SHAPE_NAME_LABEL[item.shape ?? 'rectangle']];
  return template.replace('{n}', String(index + 1));
};

const SVG_NS = 'http://www.w3.org/2000/svg';
const MIN_W = 48;
const MIN_H = 32;
const KEYBOARD_STEP = 8;

const illustrationAuthorization = (
  commandId: string,
  origin: IllustrationMutationAuthorization['origin'] = 'contextMenu',
): IllustrationMutationAuthorization => ({ origin, commandId });

const renderedPosition = (
  store: SpreadsheetStore,
  item: SessionIllustration,
): { x: number; y: number; w: number; h: number } =>
  item.nativeAnchor
    ? sessionPosition(store, item.nativeAnchor)
    : {
        x: item.x ?? DEFAULT_ILLUSTRATION_FRAME.x,
        y: item.y ?? DEFAULT_ILLUSTRATION_FRAME.y,
        w: item.w ?? DEFAULT_ILLUSTRATION_FRAME.w,
        h: item.h ?? DEFAULT_ILLUSTRATION_FRAME.h,
      };

/** A moved origin; an anchored object may sit scrolled off, so only a free
 *  object is kept on the grid. */
const movedOrigin = (item: SessionIllustration, value: number): number =>
  item.nativeAnchor ? value : Math.max(0, value);

const appendSvg = <K extends keyof SVGElementTagNameMap>(
  parent: Element,
  tag: K,
  attrs: Record<string, string | number>,
): SVGElementTagNameMap[K] => {
  const el = document.createElementNS(SVG_NS, tag);
  for (const [key, value] of Object.entries(attrs)) el.setAttribute(key, String(value));
  parent.appendChild(el);
  return el;
};

/** CSS pixels per point. */
const PX_PER_PT = 96 / 72;

/** Draw the shape filling a `w` x `h` box, its outline centred on the edge the
 *  way the spreadsheet draws it. */
const drawShape = (svg: SVGSVGElement, item: SessionIllustration, w: number, h: number): void => {
  svg.replaceChildren();
  svg.setAttribute('viewBox', `0 0 ${w} ${h}`);
  if (!item.shape) return;
  const outline = shapeOutlineColor(item);
  const width = (shapeOutlineWeight(item) ?? 0) * PX_PER_PT;
  const dash = shapeDashStroke(shapeOutlineDash(item), width);
  const stroke = {
    stroke: outline ?? 'none',
    'stroke-width': width,
    'stroke-dasharray': dash.dasharray,
    'stroke-linecap': dash.linecap,
  };
  if (item.shape === 'line' || item.shape === 'arrow') {
    // An arrowhead three line widths long ends the line at its tip.
    const head = item.shape === 'arrow' ? Math.max(6, width * 3) : 0;
    const length = Math.hypot(w, h) || 1;
    const [ux, uy] = [w / length, h / length];
    // A line runs from the top-left corner of its box to the bottom-right, as the file stores it.
    const [bx, by] = [w - ux * head, h - uy * head];
    appendSvg(svg, 'line', { x1: 0, y1: 0, x2: bx, y2: by, ...stroke });
    if (head > 0 && outline) {
      const [px, py] = [-uy * (head / 2), ux * (head / 2)];
      appendSvg(svg, 'polygon', {
        points: `${w},${h} ${bx + px},${by + py} ${bx - px},${by - py}`,
        fill: outline,
      });
    }
    return;
  }
  const paint = {
    fill: shapeFillColor(item) ?? 'none',
    'fill-opacity': item.opacity ?? 1,
    ...stroke,
    'stroke-linejoin': 'miter',
  };
  if (item.shape === 'oval') {
    appendSvg(svg, 'ellipse', { cx: w / 2, cy: h / 2, rx: w / 2, ry: h / 2, ...paint });
    return;
  }
  if (item.shape === 'triangle') {
    appendSvg(svg, 'polygon', { points: `${w / 2},0 ${w},${h} 0,${h}`, ...paint });
    return;
  }
  if (item.shape === 'diamond') {
    appendSvg(svg, 'polygon', {
      points: `${w / 2},0 ${w},${h / 2} ${w / 2},${h} 0,${h / 2}`,
      ...paint,
    });
    return;
  }
  const rx = shapeCornerRadius(item, w, h);
  appendSvg(svg, 'rect', { x: 0, y: 0, width: w, height: h, rx, ...paint });
};

/** Text colour that reads on `fill`; black on no fill. */
const textColorOn = (fill: string | null): string => {
  const argb = fill ? argbFromCssColor(fill) : null;
  if (argb === null) return '#000';
  const [r, g, b] = [(argb >>> 16) & 0xff, (argb >>> 8) & 0xff, argb & 0xff];
  return 0.299 * r + 0.587 * g + 0.114 * b < 140 ? '#fff' : '#000';
};

/** Centre the shape's text lines in its `w` x `h` box. */
const drawShapeText = (svg: SVGSVGElement, item: SessionIllustration, w: number, h: number) => {
  if (!item.text || item.shape === 'line' || item.shape === 'arrow') return;
  const lines = item.text.split('\n');
  const size = 14;
  const text = appendSvg(svg, 'text', {
    x: w / 2,
    y: h / 2 - ((lines.length - 1) * size * 1.2) / 2,
    'text-anchor': 'middle',
    'dominant-baseline': 'central',
    'font-size': size,
    fill: textColorOn(shapeFillColor(item)),
  });
  for (const [index, line] of lines.entries()) {
    const span = appendSvg(text, 'tspan', { x: w / 2, dy: index === 0 ? 0 : size * 1.2 });
    span.textContent = line;
  }
};

const renderShape = (svg: SVGSVGElement, item: SessionIllustration, w: number, h: number): void => {
  drawShape(svg, item, w, h);
  drawShapeText(svg, item, w, h);
};

const appendImage = (panel: HTMLElement, item: SessionIllustration): void => {
  const img = document.createElement('img');
  img.src = item.src ?? '';
  img.alt = item.alt ?? '';
  img.draggable = false;
  img.style.cssText = 'display:block;width:100%;height:100%;object-fit:contain;';
  panel.appendChild(img);
};

export function attachSessionIllustrations(deps: {
  host: HTMLElement;
  /** The element the pictures and shapes are drawn over, whose top-left corner is the
   *  grid's; defaults to `host`. */
  layer?: HTMLElement;
  store: SpreadsheetStore;
  history?: History | null;
  labels?: Partial<SessionIllustrationLabels>;
  getWb?: () => WorkbookHandle | null;
}): SessionIllustrationsHandle {
  const { host, store, history = null, getWb = (): WorkbookHandle | null => null } = deps;
  let labels: SessionIllustrationLabels = { ...DEFAULT_LABELS, ...deps.labels };
  const root = document.createElement('div');
  root.className = 'fc-illustrations';
  root.style.cssText = 'position:absolute;inset:0;pointer-events:none;';
  (deps.layer ?? host).appendChild(root);
  let selectedId: string | null = null;
  const selection = createObjectSelectionPublisher(store);

  /** Publish the selected object to the name box, as Excel does. */
  const syncSelection = (): void => {
    const state = store.getState();
    const item =
      selectedId === null
        ? undefined
        : state.illustrations.illustrations.find((entry) => entry.id === selectedId);
    selection.publish(item ? { id: item.id, name: objectName(state, item, labels) } : null);
  };

  const refreshSelection = (): void => {
    syncSelection();
    const panels = Array.from(root.querySelectorAll<HTMLElement>('.fc-illustration'));
    for (const [index, panel] of panels.entries()) {
      const selected = panel.dataset.illustrationId === selectedId;
      panel.classList.toggle('fc-illustration--selected', selected);
      panel.setAttribute('aria-selected', selected ? 'true' : 'false');
      panel.style.zIndex = selected ? String(panels.length + 1) : String(index + 1);
    }
  };

  const select = (id: string, panel: HTMLElement): void => {
    selectedId = id;
    refreshSelection();
    if (document.activeElement !== panel) panel.focus({ preventScroll: true });
  };

  const deselect = (): void => {
    if (selectedId === null) return;
    selectedId = null;
    refreshSelection();
  };

  const menu = createObjectMenu(host);
  /** Whether `target` is part of an object, its menu or the ribbon, which keep it selected. */
  const insideObjects = (target: EventTarget | null): boolean =>
    target instanceof Node &&
    (root.contains(target) || menu.contains(target) || keepsObjectSelected(target));

  // Like Excel, a press anywhere else on the sheet leaves the object.
  const onDocumentPointerDown = (e: PointerEvent): void => {
    if (!insideObjects(e.target)) deselect();
  };
  document.addEventListener('pointerdown', onDocumentPointerDown, true);
  const onFocusOut = (e: FocusEvent): void => {
    if (e.relatedTarget !== null && !insideObjects(e.relatedTarget)) deselect();
  };
  root.addEventListener('focusout', onFocusOut);
  root.addEventListener('keydown', keepKeyFromGrid);

  const isEditable = (item: SessionIllustration): boolean =>
    item.nativeObjectId === undefined || item.nativeGeometryEditable === true;

  const place = (
    item: SessionIllustration,
    patch: Partial<DrawingFrame>,
    commandId: 'sessionIllustration.move' | 'sessionIllustration.resize',
    origin?: IllustrationMutationAuthorization['origin'],
  ): void => {
    updateSessionIllustration(
      store,
      item.id,
      {
        ...patch,
        ...(patch.x === undefined ? {} : { x: movedOrigin(item, patch.x) }),
        ...(patch.y === undefined ? {} : { y: movedOrigin(item, patch.y) }),
      },
      history,
      getWb() ?? undefined,
      illustrationAuthorization(commandId, origin),
    );
  };

  /** Drag the whole object (`handle` omitted) or one of its sizing handles.
   *  The frame follows the pointer and the object changes once, on release. */
  const applyDrag = (
    e: PointerEvent,
    item: SessionIllustration,
    panel: HTMLElement,
    handle?: DrawingHandle,
  ): void => {
    if (e.button !== 0 || !isEditable(item)) return;
    e.preventDefault();
    trackFrameDrag(e, {
      start: renderedPosition(store, item),
      ...(handle ? { handle } : {}),
      min: { w: MIN_W, h: MIN_H },
      snap: (axis, value) => nearestGridline(store, axis, value),
      // Excel locks a picture's aspect ratio; a shape keeps it with Shift.
      keepRatio: (move) => item.kind === 'image' || move.shiftKey,
      preview: (patch) => previewFrame(panel, patch),
      commit: (patch) => {
        place(item, patch, handle ? 'sessionIllustration.resize' : 'sessionIllustration.move');
        // A refused change leaves the store as it was; put the frame back.
        render();
      },
    });
  };

  /** The object's right-click menu, in the spreadsheet's order. */
  const openMenu = (item: SessionIllustration, panel: HTMLElement, x: number, y: number): void => {
    const state = store.getState();
    const sheetItems = state.illustrations.illustrations.filter(
      (entry) => entry.sheet === item.sheet,
    );
    const rank = sheetItems.indexOf(item);
    const arrange = (action: 'bring-front' | 'send-back'): void => {
      arrangeSessionIllustration(
        store,
        item.id,
        action,
        history,
        illustrationAuthorization('sessionIllustration.arrange'),
        getWb() ?? undefined,
      );
    };
    const editable = isEditable(item);
    // The engine orders pictures only; a saved shape keeps its slot.
    const orderable = editable && !(item.kind === 'shape' && item.nativeObjectId !== undefined);
    menu.open(
      x,
      y,
      [
        {
          id: 'bringToFront',
          label: labels.bringToFront,
          disabled: !orderable || rank === sheetItems.length - 1,
          run: () => arrange('bring-front'),
        },
        {
          id: 'sendToBack',
          label: labels.sendToBack,
          disabled: !orderable || rank === 0,
          run: () => arrange('send-back'),
        },
      ],
      { label: objectName(state, item, labels), restoreFocusTo: panel },
    );
  };

  const applyKeyboard = (event: KeyboardEvent, item: SessionIllustration, panel: HTMLElement) => {
    if (event.key === 'Escape') {
      event.preventDefault();
      deselect();
      host.focus({ preventScroll: true });
      return;
    }
    if (event.key === 'ContextMenu' || (event.key === 'F10' && event.shiftKey)) {
      event.preventDefault();
      const rect = panel.getBoundingClientRect();
      openMenu(item, panel, rect.left + 8, rect.top + 8);
      return;
    }
    if (event.key === 'Delete' || event.key === 'Backspace') {
      event.preventDefault();
      if (!canDeleteIllustration(item)) return;
      clearSessionIllustration(
        store,
        item.id,
        history,
        getWb() ?? undefined,
        illustrationAuthorization('sessionIllustration.delete', 'keyboard'),
      );
      host.focus({ preventScroll: true });
      return;
    }
    if (!isEditable(item)) return;
    if (isDuplicateShortcut(event)) {
      event.preventDefault();
      const copy = duplicateSessionIllustration(
        store,
        item.id,
        history,
        getWb() ?? undefined,
        illustrationAuthorization('sessionIllustration.duplicate', 'keyboard'),
      );
      // Like the spreadsheet, the copy takes the selection.
      const copyPanel = Array.from(root.querySelectorAll<HTMLElement>('.fc-illustration')).find(
        (candidate) => candidate.dataset.illustrationId === copy?.id,
      );
      if (copy && copyPanel) select(copy.id, copyPanel);
      return;
    }
    const deltas: Record<string, [number, number]> = {
      ArrowLeft: [-KEYBOARD_STEP, 0],
      ArrowRight: [KEYBOARD_STEP, 0],
      ArrowUp: [0, -KEYBOARD_STEP],
      ArrowDown: [0, KEYBOARD_STEP],
    };
    const delta = deltas[event.key];
    if (!delta) return;
    event.preventDefault();
    const [dx, dy] = delta;
    const rendered = renderedPosition(store, item);
    if (event.shiftKey) {
      place(
        item,
        {
          w: Math.max(MIN_W, (item.nativeAnchor ? rendered.w : (item.w ?? panel.offsetWidth)) + dx),
          h: Math.max(
            MIN_H,
            (item.nativeAnchor ? rendered.h : (item.h ?? panel.offsetHeight)) + dy,
          ),
        },
        'sessionIllustration.resize',
        'keyboard',
      );
      return;
    }
    place(
      item,
      {
        x: (item.nativeAnchor ? rendered.x : (item.x ?? panel.offsetLeft)) + dx,
        y: (item.nativeAnchor ? rendered.y : (item.y ?? panel.offsetTop)) + dy,
      },
      'sessionIllustration.move',
      'keyboard',
    );
  };

  const render = (): void => {
    const state = store.getState();
    const focusedId = focusedPanelId(root, '.fc-illustration', 'illustrationId');
    root.replaceChildren();
    const items = state.illustrations.illustrations.filter(
      (item) => item.sheet === state.data.sheetIndex,
    );
    if (selectedId !== null && !items.some((item) => item.id === selectedId)) selectedId = null;
    // A paste selects the object it adds; selecting another object leaves this one.
    const wanted = state.ui.selectedObject?.id;
    const changed = wanted !== undefined && wanted !== selectedId;
    const adopted = changed && items.some((item) => item.id === wanted) ? wanted : null;
    // Another object took the selection, and the focus with it.
    const left = changed && adopted === null && selectedId !== null;
    if (changed) selectedId = adopted;
    for (const item of items) {
      const panel = document.createElement('section');
      panel.className = 'fc-illustration';
      panel.dataset.illustrationId = item.id;
      panel.tabIndex = 0;
      panel.setAttribute('role', 'group');
      panel.setAttribute(
        'aria-roledescription',
        item.kind === 'image' ? labels.picture : labels.shape,
      );
      panel.setAttribute(
        'aria-label',
        item.kind === 'image'
          ? item.alt || objectName(state, item, labels)
          : objectName(state, item, labels),
      );
      const editable = isEditable(item);
      projectDisabledState(panel, !editable, null);
      panel.style.cssText =
        'position:absolute;pointer-events:auto;background:transparent;box-sizing:border-box;';
      const placement = renderedPosition(store, item);
      panel.style.left = `${placement.x}px`;
      panel.style.top = `${placement.y}px`;
      panel.style.width = `${placement.w}px`;
      panel.style.height = `${placement.h}px`;
      if (item.nativeAnchor) {
        panel.style.clipPath = drawingClipPath(store, item.nativeAnchor, placement);
      }
      panel.addEventListener('focus', () => {
        selectedId = item.id;
        refreshSelection();
      });
      panel.addEventListener('pointerdown', (event) => {
        select(item.id, panel);
        applyDrag(event, item, panel);
      });
      panel.addEventListener('keydown', (event) => applyKeyboard(event, item, panel));
      // The object's own menu, not the cell menu under it.
      panel.addEventListener('contextmenu', (event) => {
        event.preventDefault();
        event.stopPropagation();
        select(item.id, panel);
        openMenu(item, panel, event.clientX, event.clientY);
      });

      if (item.kind === 'image') {
        appendImage(panel, item);
      } else {
        const svg = document.createElementNS(SVG_NS, 'svg');
        svg.setAttribute('aria-hidden', 'true');
        // A frame being dragged stretches the drawing until it is redrawn.
        svg.setAttribute('preserveAspectRatio', 'none');
        svg.style.cssText = 'display:block;width:100%;height:100%;overflow:visible;';
        const scale = item.nativeAnchor ? viewportZoom(store) : 1;
        renderShape(svg, item, placement.w / scale, placement.h / scale);
        panel.appendChild(svg);
      }

      if (editable) {
        appendDrawingHandles(panel, (event, handle) => {
          select(item.id, panel);
          applyDrag(event, item, panel, handle);
        });
      }
      root.appendChild(panel);
    }
    refreshSelection();
    // Rebuilding the panels drops focus; keep it on the object, or on the sheet.
    const refocus = adopted ?? (left ? undefined : focusedId);
    if (refocus !== undefined) {
      focusPanel(root, '.fc-illustration', 'illustrationId', refocus, host);
    }
  };

  const unsubscribe = store.subscribe(() => {
    if (!selection.publishing) render();
  });
  render();

  return {
    refresh: render,
    setLabels(next) {
      labels = { ...labels, ...next };
      render();
    },
    detach() {
      unsubscribe();
      document.removeEventListener('pointerdown', onDocumentPointerDown, true);
      root.removeEventListener('focusout', onFocusOut);
      root.removeEventListener('keydown', keepKeyFromGrid);
      menu.detach();
      selectedId = null;
      syncSelection();
      root.remove();
    },
  };
}
