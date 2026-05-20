export interface InteractionButtonOptions {
  className: string;
  text?: string;
  dataset?: Record<string, string>;
  role?: string;
  ariaLabel?: string;
  pressed?: boolean;
  selected?: boolean;
  tabIndex?: number;
}

export interface InteractionChipButtonOptions {
  className: string;
  label: string;
  dataset?: Record<string, string>;
  role?: string;
  pressed?: boolean;
  selected?: boolean;
  tabIndex?: number;
}

export const createInteractionButton = (opts: InteractionButtonOptions): HTMLButtonElement => {
  const button = document.createElement('button');
  button.type = 'button';
  button.className = opts.className;
  if (opts.text !== undefined) button.textContent = opts.text;
  if (opts.role) button.setAttribute('role', opts.role);
  if (opts.ariaLabel) button.setAttribute('aria-label', opts.ariaLabel);
  if (opts.pressed !== undefined) {
    button.setAttribute('aria-pressed', String(opts.pressed));
  }
  if (opts.selected !== undefined) {
    button.setAttribute('aria-selected', String(opts.selected));
  }
  if (opts.tabIndex !== undefined) button.tabIndex = opts.tabIndex;
  if (opts.dataset) {
    for (const [key, value] of Object.entries(opts.dataset)) button.dataset[key] = value;
  }
  return button;
};

export const createInteractionChipButton = (
  opts: InteractionChipButtonOptions,
): HTMLButtonElement =>
  createInteractionButton({
    ...opts,
    text: opts.label,
  });
