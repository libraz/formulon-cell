import {
  moveSheet,
  mutators,
  removeSheet,
  renameSheet,
  type SpreadsheetInstance,
  setSheetHidden,
} from '@libraz/formulon-cell';
import { showConfirm, showPrompt } from './dialogs.js';
import { focusMenuItem, handleMenuKeydown, prepareMenu } from './menu-a11y.js';

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

  addItem('Rename…', false, () => {
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
      if (renameSheet(wb, idx, trimmed)) renderSheetTabs();
    });
  });
  addItem('Delete', n <= 1, () => {
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
  addItem('Hide tab', !wb.capabilities.sheetTabHidden || visibleCount <= 1, () => {
    if (setSheetHidden(store, wb, inst.history, idx, true)) {
      const newActive = store.getState().data.sheetIndex;
      mutators.replaceCells(store, wb.cells(newActive));
      renderSheetTabs();
    }
  });
  const sep = document.createElement('div');
  sep.className = 'app__menu-sep';
  menu.appendChild(sep);
  addItem('Move left', idx === 0, () => {
    if (moveSheet(store, wb, idx, idx - 1)) renderSheetTabs();
  });
  addItem('Move right', idx >= n - 1, () => {
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
