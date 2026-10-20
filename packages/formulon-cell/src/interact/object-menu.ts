import { focusMenuItem, handleMenuKeydown } from '../toolbar/menu-a11y.js';
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
  const root = document.createElement('div');
  root.className = 'fc-ctxmenu fc-ctxmenu--object';
  root.setAttribute('role', 'menu');
  root.style.display = 'none';
  root.tabIndex = -1;
  overlayPortalFor(host).appendChild(root);
  let restoreFocusTo: HTMLElement | null = null;

  const close = (restoreFocus = false): void => {
    if (root.style.display === 'none') return;
    root.style.display = 'none';
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
  const onDocumentPointerDown = (event: PointerEvent): void => {
    if (!(event.target instanceof Node && root.contains(event.target))) close();
  };
  document.addEventListener('pointerdown', onDocumentPointerDown, true);
  const onWindowBlur = (): void => close();
  window.addEventListener('blur', onWindowBlur);

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
        if (entry.disabled) {
          button.disabled = true;
          button.setAttribute('aria-disabled', 'true');
        }
        button.addEventListener('click', () => {
          if (button.disabled) return;
          close(true);
          entry.run();
        });
        root.appendChild(button);
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
    isOpen: () => root.style.display !== 'none',
    contains: (target) => target instanceof Node && root.contains(target),
    detach() {
      document.removeEventListener('pointerdown', onDocumentPointerDown, true);
      window.removeEventListener('blur', onWindowBlur);
      root.remove();
    },
  };
}
