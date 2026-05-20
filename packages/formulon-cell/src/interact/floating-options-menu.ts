import { createInteractionButton } from './chip-button.js';

export interface FloatingOptionsButtonOptions {
  className: string;
  hasPopup?: 'menu' | 'true';
  hidden?: boolean;
}

export interface FloatingOptionsMenuItemOptions {
  className: string;
  mode: string;
  role?: 'menuitem' | 'menuitemradio';
}

export const createFloatingOptionsButton = (
  opts: FloatingOptionsButtonOptions,
): HTMLButtonElement => {
  const button = createInteractionButton({ className: opts.className });
  button.setAttribute('aria-haspopup', opts.hasPopup ?? 'menu');
  if (opts.hidden ?? true) button.style.display = 'none';
  return button;
};

export const createFloatingOptionsMenuItem = (
  opts: FloatingOptionsMenuItemOptions,
): HTMLButtonElement => {
  return createInteractionButton({
    className: opts.className,
    dataset: { fcMode: opts.mode },
    role: opts.role ?? 'menuitem',
  });
};
