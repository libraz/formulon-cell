// Swatch palette flyouts for the Format Cells font and border tabs.

import { clampPanelToViewport } from './overlay-position.js';

/** A swatch palette that hangs off a color control instead of sitting inline. */
export interface PaletteFlyout {
  readonly toggle: HTMLButtonElement;
  setOpen(open: boolean): void;
  isOpen(): boolean;
  /** True when `node` is inside the flyout or its trigger. */
  owns(node: Node): boolean;
}

/** Wire a chevron trigger to a palette flyout.
 *
 * The font and border tabs have no room left for a palette inline, so theirs
 * hang off the color control. The flyout is `position: fixed`, escaping the
 * panel's clip rect, and is placed against the viewport like the grid's own
 * menus. */
export function createPaletteFlyout(
  toggle: HTMLButtonElement,
  flyout: HTMLElement,
  palette: { focus(): void },
): PaletteFlyout {
  return {
    toggle,
    setOpen(open: boolean): void {
      flyout.hidden = !open;
      toggle.setAttribute('aria-expanded', open ? 'true' : 'false');
      if (!open) return;
      flyout.style.left = '-9999px';
      flyout.style.top = '-9999px';
      const anchor = toggle.getBoundingClientRect();
      const { x, y } = clampPanelToViewport(flyout, anchor.left, anchor.bottom + 4, { pad: 8 });
      flyout.style.left = `${x}px`;
      flyout.style.top = `${y}px`;
      palette.focus();
    },
    isOpen: () => !flyout.hidden,
    owns: (node: Node) => flyout.contains(node) || toggle.contains(node),
  };
}
