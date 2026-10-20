import type { History } from '../commands/history.js';
import type { IllustrationMutationAuthorization } from '../commands/session-illustration.js';
import {
  clearSessionIllustration,
  updateSessionIllustration,
} from '../commands/session-illustration.js';
import { drawingClipPath, sessionPosition } from '../engine/drawing-geometry.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SessionIllustration, SpreadsheetStore } from '../store/store.js';
import type { SessionShapeKind, State } from '../store/types.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import {
  appendDrawingHandles,
  createObjectSelectionPublisher,
  type DrawingFrame,
  type DrawingHandle,
  previewFrame,
  trackFrameDrag,
} from './drawing-frame.js';

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
        x: item.x ?? 300,
        y: item.y ?? 88,
        w: item.w ?? 160,
        h: item.h ?? 96,
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

const renderShape = (svg: SVGSVGElement, item: SessionIllustration): void => {
  svg.replaceChildren();
  const color = item.color ?? '#0f6cbd';
  const opacity = item.opacity ?? 0.16;
  const lineWidth = item.lineWidth ?? 3;
  if (!item.shape) return;
  if (item.shape === 'line' || item.shape === 'arrow') {
    appendSvg(svg, 'line', {
      x1: 12,
      y1: 68,
      x2: 148,
      y2: 20,
      stroke: color,
      'stroke-width': Math.max(1, item.lineWidth ?? 5),
      'stroke-linecap': 'round',
    });
    if (item.shape === 'arrow') {
      appendSvg(svg, 'path', {
        d: 'M 148 20 L 132 18 L 142 34 Z',
        fill: color,
      });
    }
    return;
  }
  if (item.shape === 'oval') {
    appendSvg(svg, 'ellipse', {
      cx: 80,
      cy: 48,
      rx: 68,
      ry: 36,
      fill: color,
      'fill-opacity': opacity,
      stroke: color,
      'stroke-width': lineWidth,
    });
    return;
  }
  if (item.shape === 'triangle') {
    appendSvg(svg, 'polygon', {
      points: '80,12 148,84 12,84',
      fill: color,
      'fill-opacity': opacity,
      stroke: color,
      'stroke-width': lineWidth,
      'stroke-linejoin': 'round',
    });
    return;
  }
  if (item.shape === 'diamond') {
    appendSvg(svg, 'polygon', {
      points: '80,10 150,48 80,86 10,48',
      fill: color,
      'fill-opacity': opacity,
      stroke: color,
      'stroke-width': lineWidth,
      'stroke-linejoin': 'round',
    });
    return;
  }
  appendSvg(svg, 'rect', {
    x: 12,
    y: 12,
    width: 136,
    height: 72,
    rx: item.radius ?? (item.shape === 'rounded-rectangle' ? 12 : 2),
    fill: color,
    'fill-opacity': opacity,
    stroke: color,
    'stroke-width': lineWidth,
  });
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
  host.appendChild(root);
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

  // Like Excel, a press anywhere else on the sheet leaves the object.
  const onDocumentPointerDown = (e: PointerEvent): void => {
    if (!(e.target instanceof Node && root.contains(e.target))) deselect();
  };
  document.addEventListener('pointerdown', onDocumentPointerDown, true);
  const onFocusOut = (e: FocusEvent): void => {
    const next = e.relatedTarget;
    if (next !== null && !(next instanceof Node && root.contains(next))) deselect();
  };
  root.addEventListener('focusout', onFocusOut);

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

  const applyKeyboard = (event: KeyboardEvent, item: SessionIllustration, panel: HTMLElement) => {
    if (event.key === 'Escape') {
      event.preventDefault();
      deselect();
      host.focus({ preventScroll: true });
      return;
    }
    if (!isEditable(item)) return;
    if (event.key === 'Delete' || event.key === 'Backspace') {
      event.preventDefault();
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
    root.replaceChildren();
    const items = state.illustrations.illustrations.filter(
      (item) => item.sheet === state.data.sheetIndex,
    );
    if (selectedId !== null && !items.some((item) => item.id === selectedId)) selectedId = null;
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

      if (item.kind === 'image') {
        appendImage(panel, item);
      } else {
        const svg = document.createElementNS(SVG_NS, 'svg');
        svg.setAttribute('viewBox', '0 0 160 96');
        svg.setAttribute('aria-hidden', 'true');
        svg.style.cssText = 'display:block;width:100%;height:100%;';
        renderShape(svg, item);
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
      selectedId = null;
      syncSelection();
      root.remove();
    },
  };
}
