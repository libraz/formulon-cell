export interface RibbonButtonOptions {
  className: string;
  id?: string;
  text?: string;
  title?: string;
  ariaLabel?: string;
  ariaHaspopup?: string;
  ariaExpanded?: boolean;
  ariaSelected?: boolean;
  ariaChecked?: boolean;
  ariaKeyshortcuts?: string;
  role?: string;
  tabIndex?: number;
  dataset?: Record<string, string>;
}

export const createRibbonButton = (opts: RibbonButtonOptions): HTMLButtonElement => {
  const button = document.createElement('button');
  button.type = 'button';
  button.className = opts.className;
  if (opts.id) button.id = opts.id;
  if (opts.text !== undefined) button.textContent = opts.text;
  if (opts.title) button.title = opts.title;
  if (opts.ariaLabel) button.setAttribute('aria-label', opts.ariaLabel);
  if (opts.ariaHaspopup) button.setAttribute('aria-haspopup', opts.ariaHaspopup);
  if (opts.ariaExpanded !== undefined) {
    button.setAttribute('aria-expanded', String(opts.ariaExpanded));
  }
  if (opts.ariaSelected !== undefined) {
    button.setAttribute('aria-selected', String(opts.ariaSelected));
  }
  if (opts.ariaChecked !== undefined) {
    button.setAttribute('aria-checked', String(opts.ariaChecked));
  }
  if (opts.ariaKeyshortcuts) button.setAttribute('aria-keyshortcuts', opts.ariaKeyshortcuts);
  if (opts.role) button.setAttribute('role', opts.role);
  if (opts.tabIndex !== undefined) button.tabIndex = opts.tabIndex;
  if (opts.dataset) {
    for (const [key, value] of Object.entries(opts.dataset)) button.dataset[key] = value;
  }
  return button;
};
