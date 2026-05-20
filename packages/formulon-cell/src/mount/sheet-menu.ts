import { clampPanelToViewport } from '../interact/overlay-position.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import { createHostButton } from './chrome-buttons.js';

export const formatSheetLabel = (template: string, name: string): string =>
  template.replace('{name}', name);

export function createSheetMenuButton(
  label: string,
  onClick: () => void,
  closeMenu: () => void,
  disabled = false,
  disabledReason: string | null = null,
): HTMLButtonElement {
  const button = createHostButton({
    className: 'fc-sheetmenu__item',
    role: 'menuitem',
    text: label,
  });
  projectDisabledState(button, disabled, disabledReason, {
    datasetKey: 'disabledReason',
    titlePrefix: label,
  });
  button.addEventListener('click', () => {
    if (button.disabled) return;
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

export function createSheetTabButton(input: {
  index: number;
  label: string;
  selected: boolean;
  tabColor?: string | null;
}): HTMLButtonElement {
  const button = createHostButton({
    className: 'fc-host__sheetbar-tab',
    role: 'tab',
    dataset: { fcSheetIndex: String(input.index) },
    ariaSelected: input.selected,
    tabIndex: input.selected ? 0 : -1,
    text: input.label,
  });
  if (input.tabColor) {
    button.dataset.fcSheetTabColor = 'true';
    button.style.setProperty('--fc-sheet-tab-color', input.tabColor);
  }
  return button;
}

export function createSheetMenuColorButton(
  label: string,
  color: string | null,
  selected: boolean,
  onClick: () => void,
): HTMLButtonElement {
  const ariaLabel = color ? `${label} ${color}` : label;
  const button = createHostButton({
    className: color
      ? 'fc-sheetmenu__swatch'
      : 'fc-sheetmenu__swatch fc-sheetmenu__swatch--none',
    role: 'menuitemradio',
    ariaLabel,
    ariaChecked: selected,
    title: ariaLabel,
  });
  if (color) button.style.setProperty('--fc-sheet-tab-color', color);
  button.addEventListener('click', () => {
    onClick();
  });
  return button;
}

export function positionSheetMenu(menu: HTMLElement, x: number, y: number): void {
  menu.hidden = false;
  menu.style.left = '0px';
  menu.style.top = '0px';
  const { x: left, y: top } = clampPanelToViewport(menu, x, y, {
    pad: 8,
    fallbackWidth: 180,
    fallbackHeight: 160,
  });
  menu.style.left = `${left}px`;
  menu.style.top = `${top}px`;
}
