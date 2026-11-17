import type { History } from '../commands/history.js';
import type {
  IllustrationMutationAuthorization,
  IllustrationNameTemplates,
} from '../commands/session-illustration.js';
import {
  arrangeSessionIllustration,
  arrangeStackPosition,
  canDeleteIllustration,
  clearSessionIllustration,
  drawingPanelZIndex,
  duplicateSessionIllustration,
  illustrationCopyRefusal,
  illustrationName,
  sheetDrawingStack,
  updateSessionIllustration,
} from '../commands/session-illustration.js';
import {
  shapeCornerRadius,
  shapeDashStroke,
  shapeFillColor,
  shapeOutlineColor,
  shapeOutlineDash,
  shapeOutlineWeight,
  shapeStyleEditable,
} from '../commands/shape-format.js';
import {
  displayRotation,
  drawingClipPath,
  nearestGridline,
  sessionPosition,
  viewportZoom,
} from '../engine/drawing-geometry.js';
import { fontPointsToCssPx } from '../engine/font-units.js';
import { lineEndsOf } from '../engine/line-style.js';
import { DEFAULT_ILLUSTRATION_FRAME } from '../engine/shapes-sync.js';
import { argbFromCssColor, cssColorFromArgb } from '../engine/theme-color.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { minorThemeFontFace } from '../engine/workbook-handle-styles.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import type { SessionIllustration, SessionLineEnd, SpreadsheetStore } from '../store/store.js';
import type { State } from '../store/types.js';
import { SVG_NS } from '../svg-namespace.js';
import { showCommandRefusal } from '../toolbar/dialogs/command-refusal.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import { focusSheet } from './dialog-focus.js';
import {
  appendDrawingHandles,
  appendRotationHandle,
  bindObjectDeselect,
  createObjectSelectionPublisher,
  type DrawingFrame,
  type DrawingHandle,
  focusedPanelId,
  focusPanel,
  isDuplicateShortcut,
  previewFrame,
  subscribeObjectOverlay,
  trackFrameDrag,
} from './drawing-frame.js';
import { createObjectMenu } from './object-menu.js';

export interface SessionIllustrationsHandle {
  refresh(): void;
  setLabels(next: Partial<SessionIllustrationLabels>): void;
  detach(): void;
}

export interface SessionIllustrationLabels extends IllustrationNameTemplates {
  /** How assistive technology names a picture and a shape. */
  picture: string;
  shape: string;
  /** The object's right-click menu. */
  bringToFront: string;
  sendToBack: string;
}

const DEFAULT_LABELS: SessionIllustrationLabels = defaultStrings.sessionIllustrations;

/** A drawing's name in the name box: the one it carries, else numbered among its kind. */
const objectName = (
  state: State,
  item: SessionIllustration,
  labels: SessionIllustrationLabels,
): string => illustrationName(state.illustrations.illustrations, item, labels);

const MIN_W = 48;
const MIN_H = 32;
const KEYBOARD_STEP = 8;

const illustrationAuthorization = (
  commandId: string,
  origin: IllustrationMutationAuthorization['origin'] = 'contextMenu',
): IllustrationMutationAuthorization => ({ origin, commandId });

/** Where `item` is drawn, in grid pixels under the current view. */
export const renderedPosition = (
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

/** Width and length of a line end, in multiples of the line width (at least 2 px). */
const END_SIZE_FACTOR = { small: 2, medium: 3, large: 5 } as const;

/** Draw one decorated end at `tip`, where the line runs towards it along the unit vector `dir`.
 *  Returns how far the line is cut short so it does not show through the decoration. */
const drawLineEnd = (
  svg: SVGElement,
  end: SessionLineEnd,
  tip: [number, number],
  dir: [number, number],
  lineWidth: number,
  outline: string,
): number => {
  const unit = Math.max(lineWidth, 2);
  const length = END_SIZE_FACTOR[end.length ?? 'medium'] * unit;
  const span = END_SIZE_FACTOR[end.width ?? 'medium'] * unit;
  const [dx, dy] = dir;
  const at = (back: number, side: number): string =>
    `${tip[0] - dx * back - dy * side},${tip[1] - dy * back + dx * side}`;
  const filled = { fill: outline, stroke: 'none' };
  switch (end.type) {
    case 'triangle':
      appendSvg(svg, 'polygon', {
        points: `${at(0, 0)} ${at(length, span / 2)} ${at(length, -span / 2)}`,
        ...filled,
      });
      return length;
    case 'stealth':
      appendSvg(svg, 'polygon', {
        points: `${at(0, 0)} ${at(length, span / 2)} ${at(length * 0.65, 0)} ${at(length, -span / 2)}`,
        ...filled,
      });
      return length * 0.65;
    case 'diamond':
      appendSvg(svg, 'polygon', {
        points: `${at(-length / 2, 0)} ${at(0, span / 2)} ${at(length / 2, 0)} ${at(0, -span / 2)}`,
        ...filled,
      });
      return length / 2;
    case 'oval':
      appendSvg(svg, 'ellipse', {
        cx: tip[0],
        cy: tip[1],
        rx: length / 2,
        ry: span / 2,
        transform: `rotate(${(Math.atan2(dy, dx) * 180) / Math.PI} ${tip[0]} ${tip[1]})`,
        ...filled,
      });
      return length / 2;
    case 'arrow':
      appendSvg(svg, 'polyline', {
        points: `${at(length, span / 2)} ${at(0, 0)} ${at(length, -span / 2)}`,
        fill: 'none',
        stroke: outline,
        'stroke-width': lineWidth,
        'stroke-linejoin': 'miter',
      });
      return 0;
  }
};

/** A line from the top-left corner of its box to the bottom-right, as the file stores it,
 *  with the decorations at its ends. */
const drawLine = (
  svg: SVGElement,
  item: SessionIllustration,
  w: number,
  h: number,
  width: number,
  outline: string | null,
  stroke: Record<string, string | number>,
): void => {
  const length = Math.hypot(w, h) || 1;
  const dir: [number, number] = [w / length, h / length];
  const ends = lineEndsOf(item);
  let [x1, y1, x2, y2] = [0, 0, w, h];
  if (outline && ends.head) {
    const cut = drawLineEnd(svg, ends.head, [0, 0], [-dir[0], -dir[1]], width, outline);
    [x1, y1] = [dir[0] * cut, dir[1] * cut];
  }
  if (outline && ends.tail) {
    const cut = drawLineEnd(svg, ends.tail, [w, h], dir, width, outline);
    [x2, y2] = [w - dir[0] * cut, h - dir[1] * cut];
  }
  appendSvg(svg, 'line', { x1, y1, x2, y2, ...stroke });
};

/** Draw the shape filling a `w` x `h` box, its outline centred on the edge the
 *  way the spreadsheet draws it. */
const drawShape = (svg: SVGSVGElement, item: SessionIllustration, w: number, h: number): void => {
  svg.replaceChildren();
  svg.setAttribute('viewBox', `0 0 ${w} ${h}`);
  if (!item.shape) return;
  // The geometry mirrors inside its own group; the text is drawn outside it.
  const mirror = [
    ...(item.flipH ? [`translate(${w} 0) scale(-1 1)`] : []),
    ...(item.flipV ? [`translate(0 ${h}) scale(1 -1)`] : []),
  ].join(' ');
  drawGeometry(mirror ? appendSvg(svg, 'g', { transform: mirror }) : svg, item, w, h);
};

const drawGeometry = (svg: SVGElement, item: SessionIllustration, w: number, h: number): void => {
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
    drawLine(svg, item, w, h, width, outline, stroke);
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

/** A new shape's text is 11 pt, inset from the top-left corner, as the spreadsheet sets it. */
const SHAPE_TEXT_SIZE = fontPointsToCssPx(11);
const SHAPE_TEXT_INSET_X = 6;
const SHAPE_TEXT_INSET_Y = 4;
const SHAPE_TEXT_LINE_HEIGHT = 1.2;

/** The text face and colour a shape takes from the workbook theme: its minor
 *  font, and the light-1 colour (dark-1 in a text box). Unset where the theme is not reachable. */
interface ShapeTextTheme {
  family?: string;
  color?: string;
}

const shapeTextTheme = (wb: WorkbookHandle | null, item: SessionIllustration): ShapeTextTheme => {
  const theme = wb?.getWorkbookTheme();
  const color = theme?.colors[item.textBox ? 0 : 1];
  const family = wb ? minorThemeFontFace(wb) : undefined;
  return {
    ...(family ? { family } : {}),
    ...(color !== undefined ? { color: cssColorFromArgb(color) } : {}),
  };
};

/** Top-left the shape's text lines in its `w` x `h` box, clipped to it. */
const drawShapeText = (
  svg: SVGSVGElement,
  item: SessionIllustration,
  w: number,
  h: number,
  theme: ShapeTextTheme,
) => {
  if (!item.text || item.shape === 'line' || item.shape === 'arrow') return;
  const lines = item.text.split('\n');
  const step = SHAPE_TEXT_SIZE * SHAPE_TEXT_LINE_HEIGHT;
  const clip = appendSvg(svg, 'svg', { x: 0, y: 0, width: w, height: h });
  const text = appendSvg(clip, 'text', {
    x: SHAPE_TEXT_INSET_X,
    y: SHAPE_TEXT_INSET_Y,
    'dominant-baseline': 'hanging',
    'font-size': SHAPE_TEXT_SIZE,
    ...(theme.family ? { 'font-family': theme.family } : {}),
    fill: theme.color ?? textColorOn(shapeFillColor(item)),
  });
  for (const [index, line] of lines.entries()) {
    const span = appendSvg(text, 'tspan', {
      x: SHAPE_TEXT_INSET_X,
      dy: index === 0 ? 0 : step,
    });
    span.textContent = line;
  }
};

const renderShape = (
  svg: SVGSVGElement,
  item: SessionIllustration,
  w: number,
  h: number,
  theme: ShapeTextTheme,
): void => {
  drawShape(svg, item, w, h);
  drawShapeText(svg, item, w, h, theme);
};

const appendImage = (panel: HTMLElement, item: SessionIllustration): void => {
  const img = document.createElement('img');
  img.src = item.src ?? '';
  img.alt = item.alt ?? '';
  img.draggable = false;
  img.style.cssText = 'display:block;width:100%;height:100%;object-fit:contain;';
  if (item.flipH || item.flipV)
    img.style.transform = `scale(${item.flipH ? -1 : 1}, ${item.flipV ? -1 : 1})`;
  panel.appendChild(img);
};

const drawingRootUsers = new WeakMap<HTMLElement, number>();

/** The one `.fc-drawings` root of `parent`, a stacking context at the grid tier
 *  in which every chart, picture and shape panel interleaves. `release` removes
 *  it once the last overlay using it leaves. */
export function acquireDrawingRoot(parent: HTMLElement): {
  root: HTMLElement;
  release: () => void;
} {
  let found = parent.querySelector<HTMLElement>(':scope > .fc-drawings');
  if (!found) {
    found = document.createElement('div');
    found.className = 'fc-drawings';
    parent.appendChild(found);
  }
  const root = found;
  drawingRootUsers.set(root, (drawingRootUsers.get(root) ?? 0) + 1);
  let released = false;
  return {
    root,
    release: () => {
      if (released) return;
      released = true;
      const users = (drawingRootUsers.get(root) ?? 1) - 1;
      drawingRootUsers.set(root, users);
      if (users <= 0) root.remove();
    },
  };
}

export function attachSessionIllustrations(deps: {
  host: HTMLElement;
  /** The element the pictures and shapes are drawn over, whose top-left corner is the
   *  grid's; defaults to `host`. */
  layer?: HTMLElement;
  store: SpreadsheetStore;
  history?: History | null;
  labels?: Partial<SessionIllustrationLabels>;
  getWb?: () => WorkbookHandle | null;
  /** The strings for the notice a refused duplicate shows. */
  getStrings?: () => Strings;
}): SessionIllustrationsHandle {
  const {
    host,
    store,
    history = null,
    getWb = (): WorkbookHandle | null => null,
    getStrings = () => defaultStrings,
  } = deps;
  let labels: SessionIllustrationLabels = { ...DEFAULT_LABELS, ...deps.labels };
  const root = document.createElement('div');
  root.className = 'fc-illustrations';
  root.style.cssText = 'position:absolute;inset:0;pointer-events:none;';
  const drawings = acquireDrawingRoot(deps.layer ?? host);
  drawings.root.appendChild(root);
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
    const state = store.getState();
    const zIndex = drawingPanelZIndex(state, state.data.sheetIndex, selectedId);
    for (const panel of root.querySelectorAll<HTMLElement>('.fc-illustration')) {
      const selected = panel.dataset.illustrationId === selectedId;
      panel.classList.toggle('fc-illustration--selected', selected);
      panel.setAttribute('aria-selected', selected ? 'true' : 'false');
      panel.style.zIndex = zIndex(panel.dataset.illustrationId ?? '');
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
  const unbindDeselect = bindObjectDeselect(root, menu, deselect);

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

  /** Open the shape's text for editing in place, the caret at the top-left corner. */
  const editText = (item: SessionIllustration, panel: HTMLElement): void => {
    if (!shapeStyleEditable(item) || item.shape === 'line' || item.shape === 'arrow') return;
    if (panel.querySelector('textarea')) return;
    const editor = document.createElement('textarea');
    editor.className = 'fc-illustration__text';
    editor.setAttribute('aria-label', objectName(store.getState(), item, labels));
    editor.value = item.text ?? '';
    const textTheme = shapeTextTheme(getWb(), item);
    editor.style.cssText =
      'position:absolute;inset:0;box-sizing:border-box;margin:0;border:0;' +
      'outline:0;resize:none;overflow:hidden;background:transparent;font:inherit;' +
      `padding:${SHAPE_TEXT_INSET_Y}px ${SHAPE_TEXT_INSET_X}px;` +
      `font-size:${SHAPE_TEXT_SIZE}px;line-height:${SHAPE_TEXT_LINE_HEIGHT};text-align:left;` +
      (textTheme.family ? `font-family:${textTheme.family};` : '') +
      `color:${textTheme.color ?? textColorOn(shapeFillColor(item))};`;
    panel.querySelector('svg text')?.remove();
    let done = false;
    const finish = (): void => {
      if (done) return;
      done = true;
      const text = editor.value;
      editor.remove();
      if (text !== (item.text ?? '')) {
        updateSessionIllustration(
          store,
          item.id,
          { text },
          history,
          getWb() ?? undefined,
          illustrationAuthorization('sessionIllustration.text', 'editor'),
        );
      }
      render();
    };
    editor.addEventListener('blur', finish);
    editor.addEventListener('pointerdown', (event) => event.stopPropagation());
    editor.addEventListener('keydown', (event) => {
      event.stopPropagation();
      if (event.key === 'Escape') {
        event.preventDefault();
        finish();
      }
    });
    panel.appendChild(editor);
    editor.focus({ preventScroll: true });
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
      rotationDeg: displayRotation(item.rotationDeg ?? 0),
      // Excel locks a picture's aspect ratio; a shape keeps it with Shift.
      keepRatio: (move) => item.kind === 'image' || move.shiftKey,
      preview: (patch) => previewFrame(panel, patch),
      commit: (patch) => {
        place(
          item,
          patch,
          handle ? 'sessionIllustration.resize' : 'sessionIllustration.move',
          'editor',
        );
        // A refused change leaves the store as it was; put the frame back.
        render();
      },
    });
  };

  /** The object's right-click menu, in the spreadsheet's order. */
  const openMenu = (item: SessionIllustration, panel: HTMLElement, x: number, y: number): void => {
    const state = store.getState();
    const wb = getWb();
    const { atFront, atBack } = arrangeStackPosition(sheetDrawingStack(state, item.sheet), item.id);
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
    const orderable =
      editable && (item.nativeObjectId === undefined || wb?.capabilities.drawings === true);
    menu.open(
      x,
      y,
      [
        {
          id: 'bringToFront',
          label: labels.bringToFront,
          disabled: !orderable || atFront,
          run: () => arrange('bring-front'),
        },
        {
          id: 'sendToBack',
          label: labels.sendToBack,
          disabled: !orderable || atBack,
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
      focusSheet(host, { preventScroll: true });
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
      focusSheet(host, { preventScroll: true });
      return;
    }
    if (isDuplicateShortcut(event)) {
      event.preventDefault();
      const refusal = illustrationCopyRefusal(item);
      if (refusal) {
        const strings = getStrings();
        const title = strings.workbookObjects.objectDuplicate;
        void showCommandRefusal({ host, strings, title, reason: refusal });
        return;
      }
      if (!isEditable(item)) return;
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
    if (!isEditable(item)) return;
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
      // Turning the panel turns its shape, text, outline and handles, and its hit area, together.
      const angle = displayRotation(item.rotationDeg ?? 0);
      if (angle !== 0) {
        panel.style.transform = `rotate(${angle}deg)`;
        panel.style.transformOrigin = '50% 50%';
      }
      if (item.nativeAnchor) {
        panel.style.clipPath = drawingClipPath(store, item.nativeAnchor, placement, angle);
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
      panel.addEventListener('dblclick', () => editText(item, panel));
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
        renderShape(
          svg,
          item,
          placement.w / scale,
          placement.h / scale,
          shapeTextTheme(getWb(), item),
        );
        panel.appendChild(svg);
      }

      if (editable) {
        appendDrawingHandles(panel, (event, handle) => {
          select(item.id, panel);
          applyDrag(event, item, panel, handle);
        });
        appendRotationHandle(panel);
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

  const unsubscribe = subscribeObjectOverlay(store, selection, render);
  render();

  return {
    refresh: render,
    setLabels(next) {
      labels = { ...labels, ...next };
      render();
    },
    detach() {
      unsubscribe();
      unbindDeselect();
      menu.detach();
      selectedId = null;
      syncSelection();
      root.remove();
      drawings.release();
    },
  };
}
