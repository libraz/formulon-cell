import {
  isWorkbookStructureProtected,
  moveSheet,
  mutators,
  removeSheet,
  renameSheet,
  type SpreadsheetInstance,
  setSheetHidden,
} from '@libraz/formulon-cell';
import { showConfirm, showPrompt } from './dialogs.js';
import { focusMenuItem, handleMenuKeydown, prepareMenu } from './menu-a11y.js';

const SHEET_TAB_COLORS = [
  '#c00000',
  '#ed7d31',
  '#ffc000',
  '#70ad47',
  '#00b0f0',
  '#4472c4',
  '#7030a0',
  '#a5a5a5',
] as const;

export function openSheetTabMenu(input: {
  closeTabMenu: () => void;
  idx: number;
  inst: SpreadsheetInstance;
  renderSheetTabs: () => void;
  setTabMenuEl: (el: HTMLDivElement) => void;
  x: number;
  y: number;
}): void {
  const { closeTabMenu, idx, inst, renderSheetTabs, setTabMenuEl, x, y } = input;
  closeTabMenu();
  const wb = inst.workbook;
  const store = inst.store;
  const n = wb.sheetCount;
  const structureProtected = isWorkbookStructureProtected(store.getState());
  const opener = document.activeElement instanceof HTMLElement ? document.activeElement : null;

  const menu = document.createElement('div');
  menu.className = 'app__menu';
  prepareMenu(menu, 'Sheet tab');
  menu.style.position = 'fixed';
  menu.style.left = `${x}px`;
  menu.style.top = `${y}px`;
  menu.style.zIndex = '90';
  let cleanupMenuListeners = (): void => {};

  const addItem = (text: string, disabled: boolean, onClick: () => void): void => {
    const it = document.createElement('button');
    it.type = 'button';
    it.className = 'app__menu-item';
    it.setAttribute('role', 'menuitem');
    it.tabIndex = -1;
    it.textContent = text;
    it.disabled = disabled;
    it.style.opacity = disabled ? '0.45' : '1';
    it.style.cursor = disabled ? 'not-allowed' : 'pointer';
    it.addEventListener('click', () => {
      if (disabled) return;
      closeTabMenu();
      cleanupMenuListeners();
      onClick();
    });
    menu.appendChild(it);
  };

  const addColorPalette = (): void => {
    const selectedColor = store.getState().layout.sheetTabColors.get(idx);
    const group = document.createElement('div');
    group.className = 'app__sheet-tab-colors';
    group.setAttribute('role', 'group');
    group.setAttribute('aria-label', 'Tab Color');

    const label = document.createElement('div');
    label.className = 'app__sheet-tab-colors-label';
    label.textContent = 'Tab Color';
    const swatches = document.createElement('div');
    swatches.className = 'app__sheet-tab-swatches';

    const addSwatch = (text: string, color: string | null, selected: boolean): void => {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = color
        ? 'app__sheet-tab-swatch'
        : 'app__sheet-tab-swatch app__sheet-tab-swatch--none';
      button.setAttribute('role', 'menuitemradio');
      button.setAttribute('aria-label', color ? `${text} ${color}` : text);
      button.setAttribute('aria-checked', selected ? 'true' : 'false');
      button.tabIndex = -1;
      button.title = color ? `${text} ${color}` : text;
      if (color) button.style.setProperty('--app-sheet-tab-color', color);
      button.addEventListener('click', () => {
        closeTabMenu();
        cleanupMenuListeners();
        mutators.setSheetTabColor(store, idx, color);
        renderSheetTabs();
      });
      swatches.appendChild(button);
    };

    addSwatch('No Color', null, selectedColor === undefined);
    for (const color of SHEET_TAB_COLORS) {
      addSwatch('Tab Color', color, selectedColor?.toLowerCase() === color);
    }
    group.append(label, swatches);
    menu.appendChild(group);
  };

  addItem('Rename…', structureProtected, () => {
    const cur = wb.sheetName(idx);
    void showPrompt({
      title: 'Rename sheet',
      label: 'Sheet name',
      initial: cur,
      placeholder: 'Sheet name',
      okLabel: 'Rename',
      validate: (v) => (v.trim().length === 0 ? 'Enter a sheet name.' : null),
    }).then((next) => {
      const trimmed = next?.trim();
      if (!trimmed || trimmed === cur) return;
      if (renameSheet(wb, idx, trimmed, store)) renderSheetTabs();
    });
  });
  addItem('Delete', structureProtected || n <= 1, () => {
    const name = wb.sheetName(idx);
    void showConfirm({
      title: 'Delete sheet',
      message: `Delete "${name}"? This action can't be undone.`,
      okLabel: 'Delete',
      destructive: true,
    }).then((ok) => {
      if (!ok) return;
      if (removeSheet(store, wb, idx)) {
        const newActive = store.getState().data.sheetIndex;
        mutators.replaceCells(store, wb.cells(newActive));
        renderSheetTabs();
      }
    });
  });
  const visibleCount = n - store.getState().layout.hiddenSheets.size;
  addItem(
    'Hide tab',
    structureProtected || !wb.capabilities.sheetTabHidden || visibleCount <= 1,
    () => {
      if (setSheetHidden(store, wb, inst.history, idx, true)) {
        const newActive = store.getState().data.sheetIndex;
        mutators.replaceCells(store, wb.cells(newActive));
        renderSheetTabs();
      }
    },
  );
  const sep = document.createElement('div');
  sep.className = 'app__menu-sep';
  menu.appendChild(sep);
  addColorPalette();
  const colorSep = document.createElement('div');
  colorSep.className = 'app__menu-sep';
  menu.appendChild(colorSep);
  addItem('Move left', structureProtected || idx === 0, () => {
    if (moveSheet(store, wb, idx, idx - 1)) renderSheetTabs();
  });
  addItem('Move right', structureProtected || idx >= n - 1, () => {
    if (moveSheet(store, wb, idx, idx + 1)) renderSheetTabs();
  });

  document.body.appendChild(menu);
  setTabMenuEl(menu);
  focusMenuItem(menu);

  const rect = menu.getBoundingClientRect();
  if (rect.right > window.innerWidth)
    menu.style.left = `${Math.max(0, window.innerWidth - rect.width - 4)}px`;
  if (rect.bottom > window.innerHeight)
    menu.style.top = `${Math.max(0, window.innerHeight - rect.height - 4)}px`;

  const onDocClick = (ev: MouseEvent): void => {
    if (ev.target instanceof Node && menu.contains(ev.target)) return;
    closeTabMenu();
    cleanupMenuListeners();
  };
  const onDocKey = (ev: KeyboardEvent): void => {
    handleMenuKeydown(ev, menu, {
      close: (restoreFocus) => {
        closeTabMenu();
        cleanupMenuListeners();
        if (restoreFocus) opener?.focus();
      },
    });
  };
  cleanupMenuListeners = () => {
    document.removeEventListener('mousedown', onDocClick, true);
    document.removeEventListener('keydown', onDocKey, true);
  };
  document.addEventListener('mousedown', onDocClick, true);
  document.addEventListener('keydown', onDocKey, true);
}
