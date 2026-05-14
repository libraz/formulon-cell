export const formatSheetLabel = (template: string, name: string): string =>
  template.replace('{name}', name);

export function createSheetMenuButton(
  label: string,
  onClick: () => void,
  closeMenu: () => void,
  disabled = false,
): HTMLButtonElement {
  const button = document.createElement('button');
  button.type = 'button';
  button.className = 'fc-sheetmenu__item';
  button.textContent = label;
  button.disabled = disabled;
  button.setAttribute('role', 'menuitem');
  button.addEventListener('click', () => {
    if (disabled) return;
    closeMenu();
    onClick();
  });
  return button;
}

export function createSheetMenuSeparator(): HTMLDivElement {
  const sep = document.createElement('div');
  sep.className = 'fc-sheetmenu__sep';
  sep.setAttribute('role', 'separator');
  return sep;
}

export function positionSheetMenu(menu: HTMLElement, x: number, y: number): void {
  menu.hidden = false;
  menu.style.left = '0px';
  menu.style.top = '0px';
  const rect = menu.getBoundingClientRect();
  const width = Math.ceil(rect.width || menu.offsetWidth || 180);
  const height = Math.ceil(rect.height || menu.offsetHeight || 160);
  const pad = 8;
  const maxX = Math.max(pad, window.innerWidth - width - pad);
  const maxY = Math.max(pad, window.innerHeight - height - pad);
  const left = Math.min(Math.max(pad, x), maxX);
  const top = Math.min(Math.max(pad, y), maxY);
  menu.style.left = `${left}px`;
  menu.style.top = `${top}px`;
}
