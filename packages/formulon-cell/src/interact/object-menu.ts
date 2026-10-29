import { focusMenuItem, handleMenuKeydown, projectDisabledState } from '../toolbar/menu-a11y.js';
import { createInteractionButton } from './chip-button.js';
import { overlayPortalFor } from './overlay-portal.js';
import { clampPanelToViewport } from './overlay-position.js';

/** One row of a chart's or picture's right-click menu. */
export type ObjectMenuEntry =
  | { id: string; label: string; run: () => void; disabled?: boolean }
  | { separator: true };

export interface ObjectMenuHandle {
  /** Show `entries` at the client point (x, y); focus returns to `restoreFocusTo`. */
  open(
    x: number,
    y: number,
    entries: readonly ObjectMenuEntry[],
    options: { label: string; restoreFocusTo: HTMLElement },
  ): void;
  close(restoreFocus?: boolean): void;
  isOpen(): boolean;
  /** Whether `target` is inside the menu, so its object stays selected. */
  contains(target: EventTarget | null): boolean;
  detach(): void;
}

const VIEWPORT_PAD = 4;

/**
 * The right-click menu of a chart, picture or shape. It takes the look and
 * keyboard handling of the cell context menu, but lists the object's own
 * commands, as the spreadsheet does for a selected object.
 */
export function createObjectMenu(host: HTMLElement): ObjectMenuHandle {
  // The menu exists only while it is open, so a closed one leaves nothing in the portal.
  const root = document.createElement('div');
  root.className = 'fc-ctxmenu fc-ctxmenu--object';
  root.setAttribute('role', 'menu');
  root.tabIndex = -1;
  let restoreFocusTo: HTMLElement | null = null;

  const onDocumentPointerDown = (event: PointerEvent): void => {
    if (!(event.target instanceof Node && root.contains(event.target))) close();
  };
  const onDocumentFocusIn = (event: FocusEvent): void => {
    if (!(event.target instanceof Node && root.contains(event.target))) close();
  };
  const onWindowBlur = (): void => close();

  const close = (restoreFocus = false): void => {
    if (!root.isConnected) return;
    document.removeEventListener('pointerdown', onDocumentPointerDown, true);
    document.removeEventListener('focusin', onDocumentFocusIn, true);
    window.removeEventListener('blur', onWindowBlur);
    root.remove();
    root.replaceChildren();
    const target = restoreFocusTo;
    restoreFocusTo = null;
    if (restoreFocus && target?.isConnected) target.focus({ preventScroll: true });
  };

  root.addEventListener('keydown', (event) => {
    if (event.key === 'Tab') {
      event.preventDefault();
      close(true);
      return;
    }
    handleMenuKeydown(event, root, { close, restoreFocusTo });
  });

  return {
    open(x, y, entries, options) {
      root.replaceChildren();
      restoreFocusTo = options.restoreFocusTo;
      root.setAttribute('aria-label', options.label);
      for (const entry of entries) {
        if ('separator' in entry) {
          const sep = document.createElement('hr');
          sep.className = 'fc-ctxmenu__sep';
          root.appendChild(sep);
          continue;
        }
        const button = createInteractionButton({
          className: 'fc-ctxmenu__item',
          dataset: { fcAction: entry.id },
          role: 'menuitem',
          tabIndex: -1,
        });
        const label = document.createElement('span');
        label.className = 'fc-ctxmenu__label';
        label.textContent = entry.label;
        button.appendChild(label);
        if (entry.disabled) projectDisabledState(button, true, null);
        button.addEventListener('click', () => {
          if (button.disabled) return;
          close(true);
          entry.run();
        });
        root.appendChild(button);
      }
      if (!root.isConnected) {
        overlayPortalFor(host).appendChild(root);
        document.addEventListener('pointerdown', onDocumentPointerDown, true);
        document.addEventListener('focusin', onDocumentFocusIn, true);
        window.addEventListener('blur', onWindowBlur);
      }
      root.style.display = 'block';
      root.style.left = '-9999px';
      root.style.top = '-9999px';
      const placed = clampPanelToViewport(root, x, y, { pad: VIEWPORT_PAD });
      root.style.left = `${placed.x}px`;
      root.style.top = `${placed.y}px`;
      focusMenuItem(root, 'first');
    },
    close,
    isOpen: () => root.isConnected,
    contains: (target) => target instanceof Node && root.contains(target),
    detach() {
      close();
    },
  };
}
