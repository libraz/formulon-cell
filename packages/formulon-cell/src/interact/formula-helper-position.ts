import { clampPanelToViewport, panelSize, viewportSize } from './overlay-position.js';

export type FormulaHelperPlacement = 'above' | 'below';

export interface FormulaCaretPositioner {
  /** Return the client rectangle at the given character offset. */
  caretRect(offset?: number): DOMRect;
  /** Whether the input anchor still intersects the viewport. */
  isAnchorVisible(): boolean;
  /** Reposition a panel around the current caret. */
  position(
    panel: HTMLElement,
    placement?: FormulaHelperPlacement,
    avoid?: HTMLElement | null,
  ): void;
  /** Listen for viewport, scroll, and selection changes that move the caret. */
  subscribe(listener: () => void): () => void;
  /** Remove the hidden measurement mirror and all shared listeners. */
  dispose(): void;
}

const MIRRORED_STYLE_PROPERTIES = [
  'fontFamily',
  'fontSize',
  'fontStyle',
  'fontWeight',
  'fontVariant',
  'fontStretch',
  'lineHeight',
  'letterSpacing',
  'wordSpacing',
  'textAlign',
  'textIndent',
  'textTransform',
  'textRendering',
  'paddingTop',
  'paddingRight',
  'paddingBottom',
  'paddingLeft',
  'borderTopWidth',
  'borderRightWidth',
  'borderBottomWidth',
  'borderLeftWidth',
  'borderTopStyle',
  'borderRightStyle',
  'borderBottomStyle',
  'borderLeftStyle',
  'overflowX',
  'overflowY',
  'boxSizing',
  'direction',
  'whiteSpace',
  'overflowWrap',
  'wordBreak',
  'tabSize',
] as const;

const rect = (left: number, top: number, width = 0, height = 0): DOMRect => {
  const right = left + width;
  const bottom = top + height;
  return {
    x: left,
    y: top,
    left,
    top,
    right,
    bottom,
    width,
    height,
    toJSON: () => ({ left, top, right, bottom, width, height }),
  } as DOMRect;
};

const overlaps = (a: DOMRect, b: DOMRect): boolean =>
  a.left < b.right && a.right > b.left && a.top < b.bottom && a.bottom > b.top;

/** Position a helper at the caret, flip when its preferred side has no room,
 * and clamp it to the viewport (including transformed overlay portals). */
export function positionFormulaHelper(
  panel: HTMLElement,
  caret: DOMRect,
  options: {
    placement?: FormulaHelperPlacement;
    avoid?: HTMLElement | null;
    align?: 'start' | 'end';
  } = {},
): void {
  const pad = 6;
  const gap = 4;
  const viewport = viewportSize(panel.ownerDocument);
  const initialRect = panel.getBoundingClientRect();
  const scaleX =
    panel.offsetWidth > 0 && initialRect.width > 0 ? initialRect.width / panel.offsetWidth : 1;
  const scaleY =
    panel.offsetHeight > 0 && initialRect.height > 0 ? initialRect.height / panel.offsetHeight : 1;
  panel.style.position = 'fixed';
  panel.style.maxWidth = `${Math.max(0, (viewport.width - pad * 2) / scaleX)}px`;
  panel.style.maxHeight = `${Math.max(0, (viewport.height - pad * 2) / scaleY)}px`;

  const size = panelSize(panel, 260, 48);
  const startX = options.align === 'end' ? caret.right - size.width : caret.left;
  const preferred = options.placement ?? 'below';
  const belowY = caret.bottom + gap;
  const aboveY = caret.top - size.height - gap;
  const fitsBelow = belowY + size.height <= viewport.height - pad;
  const fitsAbove = aboveY >= pad;
  let y = preferred === 'below' ? belowY : aboveY;
  if (preferred === 'below' && !fitsBelow && fitsAbove) y = aboveY;
  if (preferred === 'above' && !fitsAbove && fitsBelow) y = belowY;

  const clampCandidate = (x: number, candidateY: number) => {
    const position = clampPanelToViewport(panel, x, candidateY, {
      pad,
      fallbackWidth: 260,
      fallbackHeight: 48,
    });
    panel.style.left = `${position.x}px`;
    panel.style.top = `${position.y}px`;
    const bounds = panel.getBoundingClientRect();
    return {
      position,
      bounds:
        bounds.width > 0 && bounds.height > 0
          ? bounds
          : rect(position.x, position.y, size.width, size.height),
    };
  };

  let placed = clampCandidate(startX, y);
  const avoidRect = options.avoid?.isConnected ? options.avoid.getBoundingClientRect() : null;
  if (avoidRect && overlaps(placed.bounds, avoidRect)) {
    const alternate = y === belowY ? aboveY : belowY;
    const alternateFits = y === belowY ? fitsAbove : fitsBelow;
    const sideY = Math.min(Math.max(y, pad), Math.max(pad, viewport.height - size.height - pad));
    const alternatives = [
      ...(alternateFits ? [{ x: startX, y: alternate }] : []),
      { x: avoidRect.right + gap, y: sideY },
      { x: avoidRect.left - size.width - gap, y: sideY },
    ];
    for (const alternative of alternatives) {
      const candidate = clampCandidate(alternative.x, alternative.y);
      if (overlaps(candidate.bounds, avoidRect)) continue;
      placed = candidate;
      break;
    }
  }

  panel.style.left = `${placed.position.x}px`;
  panel.style.top = `${placed.position.y}px`;
}

/** Create a hidden text mirror so textarea wrapping, typography, direction,
 * padding, borders, and scrolling determine the caret's client coordinates. */
export function createFormulaCaretPositioner(
  input: HTMLInputElement | HTMLTextAreaElement,
): FormulaCaretPositioner {
  const doc = input.ownerDocument;
  const view = doc.defaultView;
  const mirror = doc.createElement('div');
  const marker = doc.createElement('span');
  mirror.className = 'fc-formula-helper-caret-mirror';
  mirror.setAttribute('aria-hidden', 'true');
  mirror.style.cssText = [
    'position: fixed',
    'left: -100000px',
    'top: 0',
    'visibility: hidden',
    'pointer-events: none',
    'overflow: hidden',
    'margin: 0',
  ].join(';');
  marker.textContent = '\u200b';
  marker.style.cssText = 'display: inline-block; width: 0; padding: 0; border: 0;';
  mirror.appendChild(marker);
  (doc.body ?? doc.documentElement).appendChild(mirror);

  const listeners = new Set<() => void>();
  let disposed = false;
  const notify = (): void => {
    if (disposed) return;
    for (const listener of listeners) listener();
  };
  const onWindowResize = (): void => notify();
  const onDocumentChange = (): void => notify();
  const onWindowScroll = (): void => notify();
  const onDocumentScroll = (event: Event): void => {
    const target = event.target;
    if (
      target === doc ||
      target === view ||
      target === doc.documentElement ||
      target === doc.body
    ) {
      notify();
      return;
    }
    const targetElement = target as Element | null;
    if (targetElement?.nodeType === 1 && targetElement.contains(input)) notify();
  };
  view?.addEventListener('resize', onWindowResize);
  view?.addEventListener('scroll', onWindowScroll);
  doc.addEventListener('selectionchange', onDocumentChange);
  doc.addEventListener('scroll', onDocumentScroll, true);

  const caretRect = (requestedOffset?: number): DOMRect => {
    const style = view?.getComputedStyle(input);
    const styleValue = (property: (typeof MIRRORED_STYLE_PROPERTIES)[number]): string => {
      const computed = style?.[property];
      return typeof computed === 'string' && computed.length > 0 ? computed : input.style[property];
    };
    const pixelValue = (property: (typeof MIRRORED_STYLE_PROPERTIES)[number]): number =>
      Number.parseFloat(styleValue(property)) || 0;
    if (style) {
      for (const property of MIRRORED_STYLE_PROPERTIES) {
        mirror.style[property] = styleValue(property);
      }
    }
    const inputRect = input.getBoundingClientRect();
    const layoutWidth = input.offsetWidth || inputRect.width;
    const layoutHeight = input.offsetHeight || inputRect.height;
    const boxSizing = styleValue('boxSizing') || 'content-box';
    const horizontalInset =
      pixelValue('paddingLeft') +
      pixelValue('paddingRight') +
      pixelValue('borderLeftWidth') +
      pixelValue('borderRightWidth');
    const verticalInset =
      pixelValue('paddingTop') +
      pixelValue('paddingBottom') +
      pixelValue('borderTopWidth') +
      pixelValue('borderBottomWidth');
    mirror.style.width = `${Math.max(0, boxSizing === 'border-box' ? layoutWidth : layoutWidth - horizontalInset)}px`;
    mirror.style.height = `${Math.max(0, boxSizing === 'border-box' ? layoutHeight : layoutHeight - verticalInset)}px`;
    const isTextarea = input.tagName === 'TEXTAREA';
    mirror.style.whiteSpace = isTextarea ? style?.whiteSpace || 'pre-wrap' : 'pre';
    mirror.style.overflowWrap = isTextarea ? style?.overflowWrap || 'break-word' : 'normal';
    mirror.style.wordBreak = isTextarea ? style?.wordBreak || 'break-word' : 'normal';

    const text = input.value;
    const offset = Math.max(
      0,
      Math.min(requestedOffset ?? input.selectionStart ?? text.length, text.length),
    );
    const fullText = doc.createTextNode(text);
    mirror.replaceChildren(fullText);
    let caretBounds: DOMRect | null = null;
    try {
      const range = doc.createRange();
      range.setStart(fullText, offset);
      range.collapse(true);
      const bounds = range.getBoundingClientRect();
      if (bounds.width > 0 || bounds.height > 0) caretBounds = bounds;
      if (!caretBounds) {
        const clientRect = [...range.getClientRects()].find(
          (candidate) => candidate.width > 0 || candidate.height > 0,
        );
        if (clientRect) caretBounds = clientRect;
      }
    } catch {
      // Some non-layout DOM implementations cannot measure a collapsed range.
    }
    if (!caretBounds) {
      // A zero-width marker covers a terminal caret or caret after a newline
      // when the platform reports no collapsed Range rectangle.
      const prefix = doc.createTextNode(text.slice(0, offset));
      const suffix = doc.createTextNode(text.slice(offset));
      mirror.replaceChildren(prefix, marker, suffix);
      caretBounds = marker.getBoundingClientRect();
    }
    const mirrorRect = mirror.getBoundingClientRect();
    const scaleX = layoutWidth > 0 ? inputRect.width / layoutWidth : 1;
    const scaleY = layoutHeight > 0 ? inputRect.height / layoutHeight : 1;
    const left = inputRect.left + (caretBounds.left - mirrorRect.left - input.scrollLeft) * scaleX;
    const top = inputRect.top + (caretBounds.top - mirrorRect.top - input.scrollTop) * scaleY;
    const height =
      Number.parseFloat(style?.lineHeight ?? '') || Number.parseFloat(style?.fontSize ?? '') || 16;
    return rect(left, top, 0, height * scaleY);
  };

  const isAnchorVisible = (): boolean => {
    const bounds = input.getBoundingClientRect();
    // DOM test environments often report a zero rectangle for visible inputs.
    if (bounds.width <= 0 || bounds.height <= 0) return true;
    const viewport = viewportSize(doc);
    return (
      bounds.right > 0 &&
      bounds.bottom > 0 &&
      bounds.left < viewport.width &&
      bounds.top < viewport.height
    );
  };

  const position = (
    panel: HTMLElement,
    placement: FormulaHelperPlacement = 'below',
    avoid: HTMLElement | null = null,
  ): void => {
    const rtl = view?.getComputedStyle(input).direction === 'rtl';
    positionFormulaHelper(panel, caretRect(), { placement, avoid, align: rtl ? 'end' : 'start' });
  };

  return {
    caretRect,
    isAnchorVisible,
    position,
    subscribe(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    dispose() {
      if (disposed) return;
      disposed = true;
      view?.removeEventListener('resize', onWindowResize);
      view?.removeEventListener('scroll', onWindowScroll);
      doc.removeEventListener('selectionchange', onDocumentChange);
      doc.removeEventListener('scroll', onDocumentScroll, true);
      listeners.clear();
      mirror.remove();
    },
  };
}
