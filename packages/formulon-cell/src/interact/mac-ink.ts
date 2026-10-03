import type { History } from '../commands/history.js';
import type { OperationIntent } from '../commands/interaction-policy.js';
import { clearSessionIllustration, createSessionImage } from '../commands/session-illustration.js';
import { recordIllustrationsChange } from '../commands/slice-history.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import {
  bodyBandOrigin,
  cellRectUnclamped,
  gridOriginX,
  gridOriginY,
  hitTest,
  layoutForView,
  type ViewLayout,
} from '../render/geometry.js';
import { mutators, type SessionIllustration, type SpreadsheetStore } from '../store/store.js';

/** The drawing tools exposed by the Mac Draw tab. */
export type MacInkTool =
  | 'pen-black'
  | 'pen-red'
  | 'pencil'
  | 'highlighter'
  | 'eraser'
  | 'lasso'
  | 'add-pen'
  | 'trackpad';

export type MacInkDrawingTool = Exclude<MacInkTool, 'eraser' | 'lasso' | 'add-pen' | 'trackpad'>;

export interface MacInkInstance {
  host: HTMLElement;
  store: SpreadsheetStore;
  history?: History | null;
  /** Optional mount command policy. Restricted embeds reject object mutation. */
  commands?: { canExecute(intent: OperationIntent): { allowed: boolean } };
  /** Source of the accessible labels; the default dictionary is used when absent. */
  i18n?: { readonly strings: Strings };
}

export interface MacInkController {
  /** Enter Draw mode with a tool. Returns false for tools the web surface does not support. */
  activate(tool?: MacInkTool): boolean;
  /** Toggle Draw mode, preserving the selected pen when no tool is supplied. */
  toggle(tool?: MacInkTool): boolean;
  /** Leave Draw mode and restore normal grid input. */
  deactivate(): void;
  /** Select a tool and enter Draw mode when it is supported. */
  setTool(tool: MacInkTool): boolean;
  /** Compatibility hook used by ribbon hosts. */
  setInkMode(mode: 'pen' | 'erase'): boolean;
  setTrackpadMode(enabled: boolean): boolean;
  isActive(): boolean;
  getTool(): MacInkTool | null;
  getTrackpadMode(): boolean;
  supports(tool: MacInkTool): boolean;
  subscribe(listener: () => void): () => void;
  dispose(): void;
}

interface InkPoint {
  row: number;
  col: number;
  /** Relative position inside the cell. This keeps a stroke attached when a
   * column width, row height, zoom, or viewport changes. */
  u: number;
  v: number;
}

interface MacInkMetadata {
  version: 1;
  tool: MacInkDrawingTool;
  points: readonly InkPoint[];
}

type MacInkIllustration = SessionIllustration & { macInk?: MacInkMetadata };

interface ToolStyle {
  color: string;
  lineWidth: number;
  opacity: number;
}

const SVG_NS = 'http://www.w3.org/2000/svg';
const MAC_INK_PREFIX = 'mac-ink-';
const HOST_CLASS = 'fc-host--mac-ink';
const INK_SELECTOR = '[data-fc-mac-ink-id]';
const MAX_POINTS = 2_000;
const DEFAULT_TOOL: MacInkDrawingTool = 'pen-black';
let clipSequence = 0;

const nextClipId = (): string => {
  clipSequence += 1;
  return `fc-mac-ink-clip-${clipSequence}`;
};
const SUPPORTED_TOOLS = new Set<MacInkTool>([
  'pen-black',
  'pen-red',
  'pencil',
  'highlighter',
  'eraser',
  'trackpad',
]);

const styleFor = (tool: MacInkDrawingTool): ToolStyle => {
  switch (tool) {
    case 'pen-red':
      return { color: '#d13438', lineWidth: 2.25, opacity: 0.98 };
    case 'pencil':
      return { color: '#5f6368', lineWidth: 1.75, opacity: 0.84 };
    case 'highlighter':
      return { color: '#ffd75e', lineWidth: 9, opacity: 0.42 };
    default:
      return { color: '#1f1f1f', lineWidth: 2.25, opacity: 0.98 };
  }
};

const finite = (value: number, fallback = 0): number => (Number.isFinite(value) ? value : fallback);

const clamp01 = (value: number): number => Math.max(0, Math.min(1, finite(value)));

const escapeXml = (value: string): string =>
  value.replaceAll('&', '&amp;').replaceAll('"', '&quot;').replaceAll('<', '&lt;');

const pathFor = (points: readonly { x: number; y: number }[], offsetX = 0, offsetY = 0): string => {
  if (points.length === 0) return '';
  const first = points[0];
  if (!first) return '';
  const start = `M ${finite(first.x - offsetX)} ${finite(first.y - offsetY)}`;
  if (points.length === 1) return `${start} l 0.01 0.01`;
  return `${start} ${points
    .slice(1)
    .map((point) => `L ${finite(point.x - offsetX)} ${finite(point.y - offsetY)}`)
    .join(' ')}`;
};

const svgImageFor = (
  points: readonly { x: number; y: number }[],
  style: ToolStyle,
): { src: string; x: number; y: number; w: number; h: number } => {
  const xs = points.map((point) => point.x);
  const ys = points.map((point) => point.y);
  const minX = Math.min(...xs);
  const minY = Math.min(...ys);
  const maxX = Math.max(...xs);
  const maxY = Math.max(...ys);
  const padding = Math.max(4, style.lineWidth * 1.5);
  const x = minX - padding;
  const y = minY - padding;
  const w = Math.max(2, maxX - minX + padding * 2);
  const h = Math.max(2, maxY - minY + padding * 2);
  const d = pathFor(points, x, y);
  const svg = `<svg xmlns="${SVG_NS}" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}"><path d="${escapeXml(d)}" fill="none" stroke="${style.color}" stroke-width="${style.lineWidth}" stroke-linecap="round" stroke-linejoin="round" opacity="${style.opacity}"/></svg>`;
  return { src: `data:image/svg+xml,${encodeURIComponent(svg)}`, x, y, w, h };
};

const metadataFor = (item: SessionIllustration): MacInkMetadata | null => {
  if (!item.id.startsWith(MAC_INK_PREFIX)) return null;
  const metadata = (item as MacInkIllustration).macInk;
  if (metadata?.version !== 1) return null;
  if (!Array.isArray(metadata.points) || metadata.points.length === 0) return null;
  return metadata;
};

const nextId = (store: SpreadsheetStore, sheet: number): string => {
  const ids = new Set(store.getState().illustrations.illustrations.map((item) => item.id));
  let index = 1;
  let id = `${MAC_INK_PREFIX}${sheet}-${index}`;
  while (ids.has(id)) {
    index += 1;
    id = `${MAC_INK_PREFIX}${sheet}-${index}`;
  }
  return id;
};

const localSize = (grid: HTMLElement, state: ReturnType<SpreadsheetStore['getState']>) => {
  const rect = grid.getBoundingClientRect();
  return {
    width: Math.max(1, finite(rect.width, state.viewport.widthPx || grid.clientWidth || 1)),
    height: Math.max(1, finite(rect.height, grid.clientHeight || 1)),
  };
};

const pointAt = (
  state: ReturnType<SpreadsheetStore['getState']>,
  x: number,
  y: number,
): InkPoint | null => {
  const layout = layoutForView(state);
  const cell = hitTest(layout, state.viewport, x, y);
  if (!cell) return null;
  const rect = cellRectUnclamped(layout, state.viewport, cell.row, cell.col);
  if (rect.w <= 0 || rect.h <= 0) return null;
  return {
    row: cell.row,
    col: cell.col,
    u: clamp01((x - rect.x) / rect.w),
    v: clamp01((y - rect.y) / rect.h),
  };
};

const pointToLocal = (
  layout: ViewLayout,
  viewport: ReturnType<SpreadsheetStore['getState']>['viewport'],
  point: InkPoint,
): { x: number; y: number } | null => {
  const rect = cellRectUnclamped(layout, viewport, point.row, point.col);
  if (rect.w <= 0 || rect.h <= 0) return null;
  return { x: rect.x + point.u * rect.w, y: rect.y + point.v * rect.h };
};

const rangeForMetadata = (sheet: number, metadata: MacInkMetadata) => ({
  sheet,
  r0: Math.min(...metadata.points.map((point) => point.row)),
  c0: Math.min(...metadata.points.map((point) => point.col)),
  r1: Math.max(...metadata.points.map((point) => point.row)),
  c1: Math.max(...metadata.points.map((point) => point.col)),
});

const localPoints = (
  state: ReturnType<SpreadsheetStore['getState']>,
  metadata: MacInkMetadata,
): { x: number; y: number }[] => {
  const layout = layoutForView(state);
  return metadata.points.flatMap((point) => {
    const local = pointToLocal(layout, state.viewport, point);
    return local ? [local] : [];
  });
};

const createSvg = <K extends keyof SVGElementTagNameMap>(
  tag: K,
  attrs: Record<string, string | number>,
): SVGElementTagNameMap[K] => {
  const element = document.createElementNS(SVG_NS, tag);
  for (const [key, value] of Object.entries(attrs)) element.setAttribute(key, String(value));
  return element;
};

interface InkClipBounds {
  x: number;
  y: number;
  width: number;
  height: number;
}

type InkPane = 'body' | 'frozen-rows' | 'frozen-cols' | 'corner';

type InkClipLayers = Record<InkPane, SVGGElement>;

const INK_PANES: readonly InkPane[] = ['body', 'frozen-rows', 'frozen-cols', 'corner'];

const clipBoundsFor = (
  state: ReturnType<SpreadsheetStore['getState']>,
  size: { width: number; height: number },
): Record<InkPane, InkClipBounds> => {
  const layout = layoutForView(state);
  const originX = gridOriginX(layout);
  const originY = gridOriginY(layout);
  const dataStartX = layout.rtl ? 0 : Math.min(size.width, Math.max(0, originX));
  const dataEndX = layout.rtl ? Math.max(0, size.width - Math.max(0, originX)) : size.width;
  const dataStartY = Math.min(size.height, Math.max(0, originY));
  const dataEndY = size.height;
  const body = bodyBandOrigin(layout, state.viewport);
  const bodyX = Math.max(dataStartX, Math.min(dataEndX, body.x));
  const bodyY = Math.max(dataStartY, Math.min(dataEndY, body.y));
  const rect = (x0: number, y0: number, x1: number, y1: number): InkClipBounds => ({
    x: Math.min(x0, x1),
    y: Math.min(y0, y1),
    width: Math.max(0, Math.abs(x1 - x0)),
    height: Math.max(0, Math.abs(y1 - y0)),
  });
  return {
    body: layout.rtl
      ? rect(dataStartX, bodyY, bodyX, dataEndY)
      : rect(bodyX, bodyY, dataEndX, dataEndY),
    'frozen-rows': layout.rtl
      ? rect(dataStartX, dataStartY, bodyX, bodyY)
      : rect(bodyX, dataStartY, dataEndX, bodyY),
    'frozen-cols': layout.rtl
      ? rect(bodyX, bodyY, dataEndX, dataEndY)
      : rect(dataStartX, bodyY, bodyX, dataEndY),
    corner: layout.rtl
      ? rect(bodyX, dataStartY, dataEndX, bodyY)
      : rect(dataStartX, dataStartY, bodyX, bodyY),
  };
};

const paneForAnchor = (layout: ViewLayout, anchor: InkPoint): InkPane => {
  const frozenRow = anchor.row < layout.freezeRows;
  const frozenCol = anchor.col < layout.freezeCols;
  if (frozenRow && frozenCol) return 'corner';
  if (frozenRow) return 'frozen-rows';
  if (frozenCol) return 'frozen-cols';
  return 'body';
};

const appendClipLayers = (
  svg: SVGSVGElement,
  clipId: string,
  bounds: Record<InkPane, InkClipBounds>,
): InkClipLayers => {
  const defs = createSvg('defs', {});
  const layers = {} as InkClipLayers;
  for (const pane of INK_PANES) {
    const clip = createSvg('clipPath', {
      id: `${clipId}-${pane}`,
      clipPathUnits: 'userSpaceOnUse',
    });
    const paneBounds = bounds[pane];
    clip.appendChild(
      createSvg('rect', {
        x: paneBounds.x,
        y: paneBounds.y,
        width: paneBounds.width,
        height: paneBounds.height,
      }),
    );
    defs.appendChild(clip);
    layers[pane] = createSvg('g', { 'clip-path': `url(#${clipId}-${pane})` });
  }
  svg.append(defs, ...INK_PANES.map((pane) => layers[pane]));
  return layers;
};

const addMetadata = (
  store: SpreadsheetStore,
  item: SessionIllustration,
  metadata: MacInkMetadata,
): void => {
  mutators.updateIllustration(store, item.id, { macInk: metadata } as unknown as Partial<
    Omit<SessionIllustration, 'id'>
  >);
};

/**
 * Attach Mac Draw ink to a mounted spreadsheet. The state item deliberately
 * uses the existing session-image channel, so Cmd/Ctrl+Z and the workbook
 * objects surface see a normal illustration. The rendered path is projected
 * from cell-relative points on every store update, which keeps it attached to
 * the sheet while the viewport scrolls or zooms.
 */
export function ensureMacInk(instance: MacInkInstance): MacInkController | undefined {
  if (typeof document === 'undefined') return undefined;
  if (instance.host.dataset.fcPlatform !== 'mac') return undefined;

  const existing = CONTROLLERS.get(instance);
  if (existing) return existing;

  const grid = instance.host.querySelector<HTMLElement>('.fc-host__grid');
  if (!grid) return undefined;

  let active = false;
  let tool: MacInkTool = DEFAULT_TOOL;
  let selectedDrawingTool: MacInkDrawingTool = DEFAULT_TOOL;
  let trackpadMode = false;
  let input: SVGSVGElement | null = null;
  let drawingPointerId: number | null = null;
  let inProgress: InkPoint[] = [];
  let platformIsMac = instance.host.dataset.fcPlatform === 'mac';
  const listeners = new Set<() => void>();

  const inkStrings = (): Strings['macInk'] => (instance.i18n?.strings ?? defaultStrings).macInk;

  const canMutate = (range: ReturnType<typeof rangeForMetadata>): boolean => {
    const commands = instance.commands;
    if (!commands) return true;
    return commands.canExecute({
      operation: 'object',
      origin: 'ribbon',
      effects: [{ kind: 'range', range }],
    }).allowed;
  };

  const display = createSvg('svg', {
    class: 'fc-mac-ink__display',
    'aria-hidden': 'true',
    'data-fc-mac-ink-display': 'true',
  });
  display.style.cssText =
    'position:absolute;inset:0;width:100%;height:100%;pointer-events:none;overflow:visible;z-index:1002;';
  grid.appendChild(display);
  const displayClipId = nextClipId();
  const inputClipId = nextClipId();

  const setDisplayVisible = (visible: boolean): void => {
    if (visible) display.removeAttribute('hidden');
    else display.setAttribute('hidden', '');
  };

  const notify = (): void => {
    for (const listener of listeners) listener();
  };

  const setHostState = (): void => {
    instance.host.classList.toggle(HOST_CLASS, active);
    if (active) instance.host.dataset.fcMacInkTool = tool;
    else delete instance.host.dataset.fcMacInkTool;
    if (active && trackpadMode) instance.host.dataset.fcMacInkTrackpad = 'true';
    else delete instance.host.dataset.fcMacInkTrackpad;
  };

  const effectiveTool = (): MacInkDrawingTool | 'eraser' | null => {
    if (!active) return null;
    if (tool === 'trackpad') return selectedDrawingTool;
    if (tool === 'eraser') return 'eraser';
    if (SUPPORTED_TOOLS.has(tool)) return tool as MacInkDrawingTool;
    return null;
  };

  const gridPoint = (event: PointerEvent): InkPoint | null => {
    const rect = grid.getBoundingClientRect();
    return pointAt(instance.store.getState(), event.clientX - rect.left, event.clientY - rect.top);
  };

  const setInputPointerMode = (): void => {
    if (!input) return;
    const mode = effectiveTool();
    input.style.pointerEvents = mode === 'eraser' ? 'none' : mode ? 'auto' : 'none';
    input.dataset.fcMacInkTool = tool;
  };

  const renderInProgress = (): void => {
    if (!input) return;
    const state = instance.store.getState();
    const size = localSize(grid, state);
    input.setAttribute('viewBox', `0 0 ${size.width} ${size.height}`);
    input.replaceChildren();
    const layers = appendClipLayers(input, inputClipId, clipBoundsFor(state, size));
    const mode = effectiveTool();
    if (!mode || mode === 'eraser' || inProgress.length === 0) return;
    const points = inProgress.flatMap((point) => {
      const local = pointToLocal(layoutForView(state), state.viewport, point);
      return local ? [local] : [];
    });
    if (points.length === 0) return;
    const style = styleFor(mode);
    const path = createSvg('path', {
      d: pathFor(points),
      fill: 'none',
      stroke: style.color,
      'stroke-width': style.lineWidth * (state.viewport.zoom || 1),
      opacity: style.opacity,
      'stroke-linecap': 'round',
      'stroke-linejoin': 'round',
      'vector-effect': 'non-scaling-stroke',
    });
    const anchor = inProgress[0];
    if (!anchor) return;
    layers[paneForAnchor(layoutForView(state), anchor)].appendChild(path);
  };

  const renderDisplay = (): void => {
    const state = instance.store.getState();
    const size = localSize(grid, state);
    display.setAttribute('viewBox', `0 0 ${size.width} ${size.height}`);
    display.replaceChildren();
    setDisplayVisible(true);
    const layers = appendClipLayers(display, displayClipId, clipBoundsFor(state, size));
    const activeSheet = state.data.sheetIndex;
    for (const item of state.illustrations.illustrations) {
      if (item.sheet !== activeSheet) continue;
      const metadata = metadataFor(item);
      if (!metadata) continue;
      const points = localPoints(state, metadata);
      if (points.length === 0) continue;
      const style = styleFor(metadata.tool);
      const path = createSvg('path', {
        d: pathFor(points),
        fill: 'none',
        stroke: style.color,
        'stroke-width': style.lineWidth * (state.viewport.zoom || 1),
        opacity: style.opacity,
        'stroke-linecap': 'round',
        'stroke-linejoin': 'round',
        'vector-effect': 'non-scaling-stroke',
        'pointer-events': active && effectiveTool() === 'eraser' ? 'stroke' : 'none',
      });
      path.dataset.fcMacInkId = item.id;
      path.setAttribute('aria-label', inkStrings().strokeLabel);
      const anchor = metadata.points[0];
      if (!anchor) continue;
      layers[paneForAnchor(layoutForView(state), anchor)].appendChild(path);
    }
  };

  const render = (): void => {
    platformIsMac = instance.host.dataset.fcPlatform === 'mac';
    setDisplayVisible(true);
    renderDisplay();
    renderInProgress();
  };

  const removeInput = (): void => {
    input?.remove();
    input = null;
    drawingPointerId = null;
    inProgress = [];
  };

  const finishStroke = (): void => {
    if (drawingPointerId === null || inProgress.length === 0) return;
    const points = inProgress.slice();
    drawingPointerId = null;
    inProgress = [];
    const mode = effectiveTool();
    if (!mode || mode === 'eraser' || points.length === 0) {
      renderInProgress();
      return;
    }
    const state = instance.store.getState();
    const local = points.flatMap((point) => {
      const resolved = pointToLocal(layoutForView(state), state.viewport, point);
      return resolved ? [resolved] : [];
    });
    if (local.length === 0) {
      renderInProgress();
      return;
    }
    const style = styleFor(mode);
    const image = svgImageFor(local, style);
    const sheet = state.data.sheetIndex;
    const id = nextId(instance.store, sheet);
    const metadata: MacInkMetadata = { version: 1, tool: mode, points };
    const hostRect = instance.host.getBoundingClientRect();
    const gridRect = grid.getBoundingClientRect();
    const range = {
      sheet,
      r0: Math.min(...points.map((point) => point.row)),
      c0: Math.min(...points.map((point) => point.col)),
      r1: Math.max(...points.map((point) => point.row)),
      c1: Math.max(...points.map((point) => point.col)),
    };
    if (!canMutate(range)) {
      renderInProgress();
      return;
    }
    recordIllustrationsChange(instance.history ?? null, instance.store, () => {
      const created = createSessionImage(
        instance.store,
        range,
        {
          id,
          src: image.src,
          alt: inkStrings().strokeLabel,
          x: gridRect.left - hostRect.left + image.x,
          y: gridRect.top - hostRect.top + image.y,
          w: image.w,
          h: image.h,
        },
        null,
      );
      if (created) addMetadata(instance.store, created, metadata);
    });
    render();
  };

  const cancelStroke = (): void => {
    drawingPointerId = null;
    inProgress = [];
    renderInProgress();
  };

  const onPointerDown = (event: PointerEvent): void => {
    if (
      !platformIsMac ||
      instance.host.dataset.fcPlatform !== 'mac' ||
      !active ||
      event.button !== 0
    )
      return;
    const mode = effectiveTool();
    if (mode === 'eraser') {
      event.preventDefault();
      event.stopPropagation();
      return;
    }
    if (!mode || !input || event.target !== input) return;
    const point = gridPoint(event);
    if (!point) return;
    event.preventDefault();
    event.stopPropagation();
    drawingPointerId = event.pointerId;
    inProgress = [point];
    input.setPointerCapture?.(event.pointerId);
    renderInProgress();
  };

  const onPointerMove = (event: PointerEvent): void => {
    if (
      !platformIsMac ||
      instance.host.dataset.fcPlatform !== 'mac' ||
      !active ||
      drawingPointerId !== event.pointerId ||
      !input
    )
      return;
    const point = gridPoint(event);
    if (!point) return;
    event.preventDefault();
    event.stopPropagation();
    if (inProgress.length < MAX_POINTS) inProgress.push(point);
    renderInProgress();
  };

  const onPointerUp = (event: PointerEvent): void => {
    if (drawingPointerId !== event.pointerId) return;
    event.preventDefault();
    event.stopPropagation();
    input?.releasePointerCapture?.(event.pointerId);
    finishStroke();
  };

  const onPointerCancel = (event: PointerEvent): void => {
    if (drawingPointerId !== event.pointerId) return;
    event.preventDefault();
    event.stopPropagation();
    input?.releasePointerCapture?.(event.pointerId);
    cancelStroke();
  };

  const onKey = (event: KeyboardEvent): void => {
    if (!active) return;
    if (event.key === 'Escape') {
      event.preventDefault();
      cancelStroke();
      deactivate();
    }
  };

  const createInput = (): void => {
    if (input) return;
    input = createSvg('svg', {
      class: 'fc-mac-ink__input',
      'aria-label': inkStrings().drawLabel,
      role: 'application',
    });
    input.style.cssText =
      'position:absolute;inset:0;width:100%;height:100%;z-index:1003;touch-action:none;overflow:visible;';
    input.addEventListener('pointerdown', onPointerDown);
    input.addEventListener('pointermove', onPointerMove);
    input.addEventListener('pointerup', onPointerUp);
    input.addEventListener('pointercancel', onPointerCancel);
    grid.appendChild(input);
    setInputPointerMode();
  };

  const onGridPointerDown = (event: PointerEvent): void => {
    if (
      !platformIsMac ||
      instance.host.dataset.fcPlatform !== 'mac' ||
      !active ||
      effectiveTool() !== 'eraser'
    )
      return;
    const target = event.target instanceof Element ? event.target.closest(INK_SELECTOR) : null;
    event.preventDefault();
    event.stopPropagation();
    if (target instanceof SVGElement) {
      const id = target.dataset.fcMacInkId;
      const item = id
        ? instance.store
            .getState()
            .illustrations.illustrations.find((candidate) => candidate.id === id)
        : undefined;
      const metadata = item ? metadataFor(item) : null;
      if (id && item && metadata && canMutate(rangeForMetadata(item.sheet, metadata)))
        clearSessionIllustration(instance.store, id, instance.history ?? null);
    }
  };

  const activate = (next?: MacInkTool): boolean => {
    if (instance.host.dataset.fcPlatform !== 'mac') {
      syncPlatform();
      return false;
    }
    if (next !== undefined) {
      if (!SUPPORTED_TOOLS.has(next)) return false;
      tool = next;
      if (next === 'pen-black' || next === 'pen-red' || next === 'pencil' || next === 'highlighter')
        selectedDrawingTool = next;
      if (next === 'trackpad') trackpadMode = true;
    }
    active = true;
    createInput();
    setInputPointerMode();
    setHostState();
    render();
    notify();
    return true;
  };

  const deactivate = (): void => {
    if (!active && !input) return;
    active = false;
    cancelStroke();
    removeInput();
    setHostState();
    renderDisplay();
    notify();
  };

  const syncPlatform = (): void => {
    const nextIsMac = instance.host.dataset.fcPlatform === 'mac';
    if (nextIsMac === platformIsMac) return;
    platformIsMac = nextIsMac;
    setDisplayVisible(nextIsMac);
    if (!nextIsMac) {
      if (active || input) {
        deactivate();
      } else {
        cancelStroke();
        setHostState();
        renderDisplay();
        notify();
      }
      return;
    }
    // Returning to Mac restores the custom projection for the current sheet,
    // while Draw remains off until the ribbon explicitly re-enters the mode.
    render();
    notify();
  };

  const unsubscribe = instance.store.subscribe(() => {
    syncPlatform();
    render();
  });
  grid.addEventListener('pointerdown', onGridPointerDown, true);
  document.addEventListener('keydown', onKey, true);

  const platformObserver =
    typeof MutationObserver !== 'undefined' ? new MutationObserver(() => syncPlatform()) : null;
  platformObserver?.observe(instance.host, {
    attributes: true,
    attributeFilter: ['data-fc-platform'],
  });

  const resizeObserver =
    typeof ResizeObserver !== 'undefined' ? new ResizeObserver(() => render()) : null;
  resizeObserver?.observe(grid);

  const controller: MacInkController = {
    activate,
    toggle(next) {
      if (active && next === undefined) {
        deactivate();
        return false;
      }
      if (active && next !== undefined && tool === next) {
        deactivate();
        return false;
      }
      return activate(next);
    },
    deactivate,
    setTool(next) {
      return activate(next);
    },
    setInkMode(mode) {
      return activate(mode === 'erase' ? 'eraser' : selectedDrawingTool);
    },
    setTrackpadMode(enabled) {
      if (instance.host.dataset.fcPlatform !== 'mac') {
        syncPlatform();
        return false;
      }
      trackpadMode = enabled;
      if (enabled) {
        tool = 'trackpad';
        active = true;
        createInput();
      } else if (tool === 'trackpad') {
        tool = selectedDrawingTool;
      }
      setInputPointerMode();
      setHostState();
      render();
      notify();
      return true;
    },
    isActive: () => active,
    getTool: () => (active ? tool : null),
    getTrackpadMode: () => trackpadMode,
    supports: (candidate) => SUPPORTED_TOOLS.has(candidate),
    subscribe(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    dispose() {
      deactivate();
      unsubscribe();
      platformObserver?.disconnect();
      resizeObserver?.disconnect();
      grid.removeEventListener('pointerdown', onGridPointerDown, true);
      document.removeEventListener('keydown', onKey, true);
      display.remove();
      listeners.clear();
      CONTROLLERS.delete(instance);
    },
  };

  CONTROLLERS.set(instance, controller);
  render();
  return controller;
}

/** Return the already-attached Mac Draw controller without creating one. */
export function getMacInk(instance: MacInkInstance): MacInkController | undefined {
  return CONTROLLERS.get(instance);
}

/** Detach the per-instance Mac Draw controller, if one was attached. */
export function disposeMacInk(instance: MacInkInstance): void {
  CONTROLLERS.get(instance)?.dispose();
}

/** Cancel an active stroke without creating a controller for non-Mac mounts. */
export function deactivateMacInk(instance: MacInkInstance): void {
  CONTROLLERS.get(instance)?.deactivate();
}

const CONTROLLERS = new WeakMap<object, MacInkController>();
