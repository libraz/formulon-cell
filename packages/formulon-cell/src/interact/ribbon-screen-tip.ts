import { overlayPortalFor } from './overlay-portal.js';
import { clampPanelToViewport, panelSize, viewportSize } from './overlay-position.js';

/** The two shortcut conventions used by the ribbon. `default` covers the
 * Windows/Linux surface as well as an explicitly unspecified platform. */
export type RibbonScreenTipPlatform = 'mac' | 'default';

export interface RibbonScreenTipsOptions {
  /** Root containing the rendered ribbon. The root may be replaced by the
   * toolbar renderer; `refresh()` hides a tip that belonged to the old tree. */
  host: HTMLElement;
  /** Resolve the sheet host that owns the overlay portal. This is useful for
   * a toolbar mounted outside the sheet's `.fc-host` element. */
  getSheetHost?: () => HTMLElement | null | undefined;
  /** Resolve the current shortcut convention without coupling this module to
   * the mount layer's platform option type. */
  getPlatform?: () => RibbonScreenTipPlatform;
  /** Delay for pointer hover. Keyboard focus remains immediate. */
  delayMs?: number;
}

export interface RibbonScreenTipsHandle {
  /** Hide an open tip and cancel a pending pointer timer. */
  hide(): void;
  /** Hide before a toolbar rerender or a locale/platform change. */
  refresh(): void;
  /** Remove listeners, observers and the overlay. Idempotent. */
  detach(): void;
}

const SCREEN_TIP_TARGETS =
  '[data-ribbon-command], [data-ribbon-tab], [data-ribbon-toggle], [data-ribbon-display-option], [title], button[aria-label], [role="button"][aria-label]';
const DEFAULT_DELAY_MS = 500;
const SCREEN_TIP_PAD = 8;
const SCREEN_TIP_GAP = 8;
const SCREEN_TIP_FALLBACK_WIDTH = 320;
const SCREEN_TIP_FALLBACK_HEIGHT = 40;
const SCREEN_TIP_DISMISS_DELAY_MS = 180;
let nextTipId = 0;

const isElement = (value: EventTarget | null): value is Element =>
  value !== null && typeof value === 'object' && 'nodeType' in value && value.nodeType === 1;

const containsTarget = (root: HTMLElement, node: EventTarget | null): node is Element =>
  isElement(node) && (node === root || root.contains(node));

const resolvedPlatform = (
  ownerDocument: Document,
  getPlatform?: () => RibbonScreenTipPlatform,
): RibbonScreenTipPlatform => {
  if (getPlatform) return getPlatform();
  const platform = ownerDocument.defaultView?.navigator.platform?.toLowerCase() ?? '';
  const userAgent = ownerDocument.defaultView?.navigator.userAgent?.toLowerCase() ?? '';
  return platform.includes('mac') || userAgent.includes('macintosh') ? 'mac' : 'default';
};

const preferredShortcut = (value: string, platform: RibbonScreenTipPlatform): string | null => {
  const alternatives = value.split(/\s+/).filter(Boolean);
  if (alternatives.length === 0) return null;
  const modifier = platform === 'mac' ? 'Meta' : 'Control';
  return (
    alternatives.find((shortcut) => shortcut.split('+').includes(modifier)) ??
    alternatives[0] ??
    null
  );
};

/** Render one declared `aria-keyshortcuts` alternative using the active
 * platform's familiar modifier glyphs. No shortcut is invented here: callers
 * only receive shortcuts already declared by the ribbon command. */
export const formatRibbonShortcut = (
  value: string | null | undefined,
  platform: RibbonScreenTipPlatform,
): string | null => {
  if (!value?.trim()) return null;
  const shortcut = preferredShortcut(value, platform);
  if (!shortcut) return null;
  const parts = shortcut.split('+').filter(Boolean);
  const key = parts.pop();
  if (!key) return null;
  if (platform === 'mac') {
    const modifierGlyphs: Record<string, string> = {
      Alt: '⌥',
      Control: '⌃',
      Meta: '⌘',
      Shift: '⇧',
    };
    return `${parts.map((part) => modifierGlyphs[part] ?? part).join('')}${key}`;
  }
  const modifierNames: Record<string, string> = {
    Alt: 'Alt',
    Control: 'Ctrl',
    Meta: 'Win',
    Shift: 'Shift',
  };
  return [...parts.map((part) => modifierNames[part] ?? part), key].join('+');
};

/** Build the visible screen-tip text from the current DOM metadata. Keeping
 * this derived from `title` and `aria-keyshortcuts` means locale rerenders do
 * not need a second translation table in the tooltip layer. */
export const ribbonScreenTipText = (
  target: HTMLElement,
  platform: RibbonScreenTipPlatform,
): string | null => {
  const title =
    target.getAttribute('title')?.trim() || target.getAttribute('aria-label')?.trim() || '';
  return screenTipText(title, target.getAttribute('aria-keyshortcuts'), platform);
};

const screenTipText = (
  title: string,
  declaredShortcuts: string | null,
  platform: RibbonScreenTipPlatform,
): string | null => {
  if (!title) return null;
  const shortcut = formatRibbonShortcut(declaredShortcuts, platform);
  if (!shortcut) return title;
  const rawShortcut = declaredShortcuts ? preferredShortcut(declaredShortcuts, platform) : null;
  const combinedShortcut = rawShortcut?.replace(/\b(?:Control|Meta)\+/g, 'Ctrl/Cmd+');
  if (combinedShortcut && combinedShortcut !== rawShortcut && title.includes(combinedShortcut)) {
    return title.replaceAll(combinedShortcut, shortcut);
  }
  const shortcutForms = [
    shortcut,
    rawShortcut,
    rawShortcut?.replace(/\bControl\+/g, 'Ctrl+').replace(/\bMeta\+/g, 'Cmd+'),
  ].filter((value): value is string => !!value);
  if (shortcutForms.some((value) => title.includes(value))) return title;
  return `${title}\n${shortcut}`;
};

const candidateFor = (host: HTMLElement, node: EventTarget | null): HTMLElement | null => {
  if (!containsTarget(host, node) || !isElement(node)) return null;
  let current: Element | null = node;
  while (current && current !== host) {
    if (current.matches(SCREEN_TIP_TARGETS)) {
      const element = current as HTMLElement;
      if (element.getAttribute('title')?.trim() || element.getAttribute('aria-label')?.trim()) {
        return element;
      }
    }
    current = current.parentElement;
  }
  return null;
};

const sameOrContains = (parent: Element | null, child: EventTarget | null): boolean =>
  !!parent && isElement(child) && (parent === child || parent.contains(child));

export function attachRibbonScreenTips(options: RibbonScreenTipsOptions): RibbonScreenTipsHandle {
  const { host } = options;
  const ownerDocument = host.ownerDocument;
  const ownerWindow = ownerDocument.defaultView;
  const tip = ownerDocument.createElement('div');
  tip.className = 'fc-tb__screen-tip';
  tip.setAttribute('role', 'tooltip');
  tip.setAttribute('aria-hidden', 'true');
  tip.hidden = true;
  tip.id = `fc-ribbon-screen-tip-${++nextTipId}`;
  if (host.lang) tip.lang = host.lang;

  const delay = Math.max(0, options.delayMs ?? DEFAULT_DELAY_MS);
  let disposed = false;
  let timer: ReturnType<typeof setTimeout> | null = null;
  let dismissTimer: ReturnType<typeof setTimeout> | null = null;
  let pendingTarget: HTMLElement | null = null;
  let openTarget: HTMLElement | null = null;
  let originalTitle: string | null = null;
  let hadTitle = false;
  let originalDescribedBy: string | null = null;
  let hadDescribedBy = false;

  const clearTimer = (): void => {
    if (timer !== null) clearTimeout(timer);
    timer = null;
    pendingTarget = null;
  };

  const clearDismissTimer = (): void => {
    if (dismissTimer === null) return;
    clearTimeout(dismissTimer);
    dismissTimer = null;
  };

  const restoreTarget = (): void => {
    const target = openTarget;
    if (!target) return;
    if (target.isConnected) {
      // A consumer may replace the title while the custom tip is visible (for
      // example when a disabled reason or locale changes). Restore only while
      // the attribute is still suppressed by this handle.
      const currentTitle = target.getAttribute('title');
      if (currentTitle === null) {
        if (hadTitle && originalTitle !== null) target.setAttribute('title', originalTitle);
      } else if (!hadTitle && currentTitle === '') {
        target.removeAttribute('title');
      }

      const currentDescribedBy = target.getAttribute('aria-describedby');
      if (currentDescribedBy?.split(/\s+/).includes(tip.id)) {
        const remaining = currentDescribedBy
          .split(/\s+/)
          .filter((value) => value && value !== tip.id)
          .join(' ');
        if (remaining) target.setAttribute('aria-describedby', remaining);
        else if (hadDescribedBy && originalDescribedBy !== null) {
          target.setAttribute('aria-describedby', originalDescribedBy);
        } else {
          target.removeAttribute('aria-describedby');
        }
      }
    }
  };

  const hide = (): void => {
    clearTimer();
    clearDismissTimer();
    restoreTarget();
    openTarget = null;
    originalTitle = null;
    hadTitle = false;
    originalDescribedBy = null;
    hadDescribedBy = false;
    tip.hidden = true;
    tip.setAttribute('aria-hidden', 'true');
  };

  const isTipTarget = (node: EventTarget | null): boolean =>
    isElement(node) && (node === tip || tip.contains(node));

  const scheduleDismiss = (): void => {
    clearDismissTimer();
    dismissTimer = setTimeout(() => {
      dismissTimer = null;
      hide();
    }, SCREEN_TIP_DISMISS_DELAY_MS);
  };

  const position = (target: HTMLElement): void => {
    const rect = target.getBoundingClientRect();
    const viewport = viewportSize(ownerDocument);
    const size = panelSize(tip, SCREEN_TIP_FALLBACK_WIDTH, SCREEN_TIP_FALLBACK_HEIGHT);
    const desiredX =
      rect.left + size.width + SCREEN_TIP_PAD <= viewport.width
        ? rect.left
        : rect.right - size.width;
    const belowY = rect.bottom + SCREEN_TIP_GAP;
    const aboveY = rect.top - size.height - SCREEN_TIP_GAP;
    const desiredY =
      belowY + size.height + SCREEN_TIP_PAD <= viewport.height || aboveY < SCREEN_TIP_PAD
        ? belowY
        : aboveY;
    const clamped = clampPanelToViewport(tip, desiredX, desiredY, { pad: SCREEN_TIP_PAD });
    tip.style.left = `${Math.round(clamped.x)}px`;
    tip.style.top = `${Math.round(clamped.y)}px`;
  };

  const show = (target: HTMLElement): void => {
    if (disposed || !target.isConnected || !host.contains(target)) return;
    const platform = resolvedPlatform(ownerDocument, options.getPlatform);
    const text = ribbonScreenTipText(target, platform);
    if (!text) return;
    const portalHost = options.getSheetHost?.() ?? host;
    const portal = overlayPortalFor(portalHost);
    if (tip.parentElement !== portal) portal.appendChild(tip);
    if (host.lang) tip.lang = host.lang;
    else tip.removeAttribute('lang');
    const fontFamily = ownerWindow?.getComputedStyle(target).fontFamily;
    if (fontFamily) tip.style.setProperty('--fc-tb-font', fontFamily);
    else tip.style.removeProperty('--fc-tb-font');
    originalTitle = target.getAttribute('title');
    hadTitle = target.hasAttribute('title');
    originalDescribedBy = target.getAttribute('aria-describedby');
    hadDescribedBy = target.hasAttribute('aria-describedby');
    openTarget = target;
    tip.textContent = text;
    tip.hidden = false;
    tip.setAttribute('aria-hidden', 'false');
    target.removeAttribute('title');
    const describedBy = [originalDescribedBy, tip.id].filter(Boolean).join(' ');
    target.setAttribute('aria-describedby', describedBy);
    position(target);
  };

  const schedule = (target: HTMLElement, immediate: boolean): void => {
    clearTimer();
    clearDismissTimer();
    if (openTarget === target) return;
    hide();
    const run = (): void => {
      timer = null;
      pendingTarget = null;
      if (disposed || !target.isConnected || !host.contains(target)) return;
      show(target);
    };
    if (immediate || delay === 0) run();
    else {
      pendingTarget = target;
      timer = setTimeout(run, delay);
    }
  };

  const onPointerOver = (event: PointerEvent): void => {
    clearDismissTimer();
    const target =
      (openTarget && containsTarget(openTarget, event.target) ? openTarget : null) ??
      candidateFor(host, event.target);
    if (!target) return;
    const related = event.relatedTarget;
    if (sameOrContains(target, related)) return;
    schedule(target, false);
  };

  const onPointerOut = (event: PointerEvent): void => {
    const target =
      (openTarget && containsTarget(openTarget, event.target) ? openTarget : null) ??
      candidateFor(host, event.target);
    if (!target || sameOrContains(target, event.relatedTarget)) return;
    if (isTipTarget(event.relatedTarget)) {
      scheduleDismiss();
      return;
    }
    if (
      pendingTarget === target ||
      openTarget === target ||
      !event.relatedTarget ||
      !containsTarget(host, event.relatedTarget)
    ) {
      if (openTarget === target) scheduleDismiss();
      else hide();
    }
  };

  const onFocusIn = (event: FocusEvent): void => {
    const target = candidateFor(host, event.target);
    if (target) schedule(target, true);
  };

  const onFocusOut = (event: FocusEvent): void => {
    const target = candidateFor(host, event.target);
    if (target && sameOrContains(target, event.relatedTarget)) return;
    hide();
  };

  const onKeyDown = (event: KeyboardEvent): void => {
    if (event.key === 'Escape') hide();
  };
  const onPointerDown = (): void => hide();

  const onHostLeave = (event: PointerEvent): void => {
    if (openTarget || isTipTarget(event.relatedTarget)) scheduleDismiss();
    else hide();
  };
  const onTipEnter = (): void => clearDismissTimer();
  const onTipLeave = (): void => scheduleDismiss();
  const onScroll = (event: Event): void => {
    if (!isTipTarget(event.target)) hide();
  };
  const onBlur = (): void => hide();
  const MutationObserverCtor = ownerWindow?.MutationObserver ?? globalThis.MutationObserver;
  const observer = new MutationObserverCtor((records) => {
    if (!openTarget) return;
    if (!openTarget.isConnected || !host.contains(openTarget)) {
      hide();
      return;
    }
    const title = openTarget.getAttribute('title');
    const metadataChanged = records.some(
      (record) => record.type === 'attributes' && record.target === openTarget,
    );
    // Ignore our own title suppression. A new title supplied by the active
    // state projection, however, replaces both the visible tip and the value
    // restored when it closes.
    const accessibleMetadataChanged = records.some(
      (record) =>
        record.target === openTarget &&
        (record.attributeName === 'aria-label' || record.attributeName === 'aria-keyshortcuts'),
    );
    if (metadataChanged && (title !== null || accessibleMetadataChanged)) {
      if (title !== null) {
        originalTitle = title;
        hadTitle = true;
      }
      const text = screenTipText(
        title?.trim() ||
          originalTitle?.trim() ||
          openTarget.getAttribute('aria-label')?.trim() ||
          '',
        openTarget.getAttribute('aria-keyshortcuts'),
        resolvedPlatform(ownerDocument, options.getPlatform),
      );
      if (!text) {
        hide();
        return;
      }
      tip.textContent = text;
      openTarget.removeAttribute('title');
      position(openTarget);
    }
  });

  host.addEventListener('pointerover', onPointerOver);
  host.addEventListener('pointerout', onPointerOut);
  host.addEventListener('focusin', onFocusIn);
  host.addEventListener('focusout', onFocusOut);
  ownerDocument.addEventListener('keydown', onKeyDown, true);
  host.addEventListener('pointerleave', onHostLeave);
  host.addEventListener('pointerdown', onPointerDown);
  tip.addEventListener('pointerenter', onTipEnter);
  tip.addEventListener('pointerleave', onTipLeave);
  ownerDocument.addEventListener('scroll', onScroll, true);
  ownerDocument.addEventListener('visibilitychange', hide);
  ownerWindow?.addEventListener('resize', onScroll);
  ownerWindow?.addEventListener('blur', onBlur);
  observer.observe(host, {
    childList: true,
    subtree: true,
    attributes: true,
    attributeFilter: ['title', 'aria-label', 'aria-keyshortcuts'],
  });

  const detach = (): void => {
    if (disposed) return;
    disposed = true;
    hide();
    host.removeEventListener('pointerover', onPointerOver);
    host.removeEventListener('pointerout', onPointerOut);
    host.removeEventListener('focusin', onFocusIn);
    host.removeEventListener('focusout', onFocusOut);
    ownerDocument.removeEventListener('keydown', onKeyDown, true);
    host.removeEventListener('pointerleave', onHostLeave);
    host.removeEventListener('pointerdown', onPointerDown);
    tip.removeEventListener('pointerenter', onTipEnter);
    tip.removeEventListener('pointerleave', onTipLeave);
    ownerDocument.removeEventListener('scroll', onScroll, true);
    ownerDocument.removeEventListener('visibilitychange', hide);
    ownerWindow?.removeEventListener('resize', onScroll);
    ownerWindow?.removeEventListener('blur', onBlur);
    observer.disconnect();
    tip.remove();
  };

  const refresh = (): void => hide();
  return { hide, refresh, detach };
}
