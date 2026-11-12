export interface ViewportPanelPositionOptions {
  pad?: number;
  fallbackWidth?: number;
  fallbackHeight?: number;
}

export interface ViewportPanelPosition {
  x: number;
  y: number;
}

interface FixedContainingBlockGeometry {
  element: HTMLElement;
  rect: DOMRect;
  scaleX: number;
  scaleY: number;
}

export const viewportSize = (ownerDocument?: Document): { width: number; height: number } => {
  const doc = ownerDocument ?? document;
  const view = doc.defaultView;
  return {
    width: view?.innerWidth || doc.documentElement.clientWidth || 1024,
    height: view?.innerHeight || doc.documentElement.clientHeight || 768,
  };
};

/** `value` within `[min, max]`; `min` wins when `max < min`, so a panel wider than the viewport
 *  stays at its leading edge. */
export const clamp = (value: number, min: number, max: number): number =>
  Math.min(Math.max(value, min), Math.max(min, max));

export const panelSize = (
  panel: HTMLElement,
  fallbackWidth = 0,
  fallbackHeight = 0,
): { width: number; height: number } => {
  const rect = panel.getBoundingClientRect();
  return {
    width: Math.ceil(rect.width || panel.offsetWidth || fallbackWidth),
    height: Math.ceil(rect.height || panel.offsetHeight || fallbackHeight),
  };
};

const computedStyleFor = (element: HTMLElement): CSSStyleDeclaration | null => {
  const view = element.ownerDocument.defaultView;
  return view?.getComputedStyle(element) ?? null;
};

const propertyOrInline = (
  style: CSSStyleDeclaration | null,
  element: HTMLElement,
  property: keyof CSSStyleDeclaration,
): string => {
  const computed = style?.[property];
  if (typeof computed === 'string' && computed.length > 0) return computed;
  return element.style[property] as string;
};

const establishesFixedContainingBlock = (element: HTMLElement): boolean => {
  const style = computedStyleFor(element);
  const transform = propertyOrInline(style, element, 'transform');
  const filter = propertyOrInline(style, element, 'filter');
  const perspective = propertyOrInline(style, element, 'perspective');
  const contain = propertyOrInline(style, element, 'contain');
  const willChange = propertyOrInline(style, element, 'willChange');
  if (
    (transform && transform !== 'none') ||
    (filter && filter !== 'none') ||
    (perspective && perspective !== 'none')
  ) {
    return true;
  }
  if (contain && contain !== 'none') {
    return contain
      .split(/\s+/)
      .some((token) => ['layout', 'paint', 'strict', 'content'].includes(token));
  }
  return willChange
    .split(',')
    .map((value) => value.trim())
    .some((value) => ['transform', 'filter', 'perspective'].includes(value));
};

const scaleFor = (rectSize: number, layoutSize: number): number => {
  if (layoutSize > 0 && Number.isFinite(rectSize / layoutSize)) {
    const scale = rectSize / layoutSize;
    if (scale > 0) return scale;
  }
  return 1;
};

/** Find the nearest ancestor that changes `position: fixed` from viewport to
 * a local containing block. Rotated/skewed transforms are intentionally left
 * to the host because scalar width/height conversion cannot preserve them. */
const fixedContainingBlockFor = (panel: HTMLElement): FixedContainingBlockGeometry | null => {
  for (let ancestor = panel.parentElement; ancestor; ancestor = ancestor.parentElement) {
    if (!establishesFixedContainingBlock(ancestor)) continue;
    const rect = ancestor.getBoundingClientRect();
    return {
      element: ancestor,
      rect,
      scaleX: scaleFor(rect.width, ancestor.offsetWidth),
      scaleY: scaleFor(rect.height, ancestor.offsetHeight),
    };
  }
  return null;
};

const toContainingBlockPosition = (
  value: number,
  edge: number,
  clientEdge: number,
  scroll: number,
  scale: number,
): number => (value - edge - clientEdge * scale + scroll * scale) / scale;

/** Default gap, in CSS px, kept between a clamped panel and the viewport edge. */
export const VIEWPORT_PANEL_PAD = 4;

export const clampPanelToViewport = (
  panel: HTMLElement,
  x: number,
  y: number,
  options: ViewportPanelPositionOptions = {},
): ViewportPanelPosition => {
  const pad = options.pad ?? VIEWPORT_PANEL_PAD;
  const { width, height } = panelSize(panel, options.fallbackWidth, options.fallbackHeight);
  const viewport = viewportSize(panel.ownerDocument);
  const clientPosition = {
    x: clamp(x, pad, viewport.width - width - pad),
    y: clamp(y, pad, viewport.height - height - pad),
  };
  const containingBlock = fixedContainingBlockFor(panel);
  if (!containingBlock) return clientPosition;

  const { element, rect, scaleX, scaleY } = containingBlock;
  return {
    x: toContainingBlockPosition(
      clientPosition.x,
      rect.left,
      element.clientLeft,
      element.scrollLeft,
      scaleX,
    ),
    y: toContainingBlockPosition(
      clientPosition.y,
      rect.top,
      element.clientTop,
      element.scrollTop,
      scaleY,
    ),
  };
};

/** Moves `panel` to `(x, y)` clamped into the viewport, on whole pixels so its text stays
 *  crisp, and returns the position written to its `left`/`top`. */
export const placePanel = (
  panel: HTMLElement,
  x: number,
  y: number,
  options: ViewportPanelPositionOptions = {},
): ViewportPanelPosition => {
  const position = clampPanelToViewport(panel, x, y, options);
  const placed = { x: Math.round(position.x), y: Math.round(position.y) };
  panel.style.left = `${placed.x}px`;
  panel.style.top = `${placed.y}px`;
  return placed;
};
