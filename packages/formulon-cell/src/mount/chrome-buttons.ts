export interface HostButtonOptions {
  className: string;
  ariaLabel?: string;
  ariaExpanded?: boolean;
  ariaSelected?: boolean;
  ariaChecked?: boolean;
  iconPaths?: readonly string[];
  role?: string;
  dataset?: Record<string, string>;
  tabIndex?: number;
  title?: string;
  text?: string;
}

export function createHostButton(opts: HostButtonOptions): HTMLButtonElement {
  const button = document.createElement('button');
  button.type = 'button';
  button.className = opts.className;
  if (opts.text !== undefined) button.textContent = opts.text;
  if (opts.tabIndex !== undefined) button.tabIndex = opts.tabIndex;
  if (opts.ariaLabel) button.setAttribute('aria-label', opts.ariaLabel);
  if (opts.role) button.setAttribute('role', opts.role);
  if (opts.ariaExpanded !== undefined) {
    button.setAttribute('aria-expanded', String(opts.ariaExpanded));
  }
  if (opts.ariaSelected !== undefined) {
    button.setAttribute('aria-selected', String(opts.ariaSelected));
  }
  if (opts.ariaChecked !== undefined) {
    button.setAttribute('aria-checked', String(opts.ariaChecked));
  }
  if (opts.title) button.title = opts.title;
  if (opts.dataset) {
    for (const [key, value] of Object.entries(opts.dataset)) button.dataset[key] = value;
  }
  if (opts.iconPaths) appendHostIcon(button, opts.iconPaths);
  return button;
}

export function appendHostIcon(
  button: HTMLButtonElement,
  paths: readonly string[],
  viewBox = '0 0 20 20',
): void {
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.setAttribute('class', 'fc-host__icon');
  svg.setAttribute('viewBox', viewBox);
  svg.setAttribute('fill', 'none');
  svg.setAttribute('stroke', 'currentColor');
  svg.setAttribute('stroke-width', '1.5');
  svg.setAttribute('stroke-linecap', 'round');
  svg.setAttribute('stroke-linejoin', 'round');
  svg.setAttribute('aria-hidden', 'true');
  for (const d of paths) {
    const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    path.setAttribute('d', d);
    svg.appendChild(path);
  }
  button.replaceChildren(svg);
}
