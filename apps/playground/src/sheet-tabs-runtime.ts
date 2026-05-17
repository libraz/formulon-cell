import {
  addSheet,
  focusMenuItem,
  handleMenuKeydown,
  isWorkbookStructureProtected,
  mutators,
  prepareMenu,
  type SpreadsheetInstance,
  setSheetHidden,
} from '@libraz/formulon-cell';
import { openSheetTabMenu } from './sheet-tab-menu.js';

export interface SheetTabsCtx {
  getInst: () => SpreadsheetInstance | null;
  focusSheet: () => void;
  statusMetric: HTMLElement | null;
  workbookStructureProtectedBlockedText: string;
}

export interface SheetTabsApi {
  renderSheetTabs: () => void;
  switchSheet: (idx: number) => void;
  openTabMenu: (idx: number, x: number, y: number) => void;
  openUnhideMenu: (x: number, y: number) => void;
  closeTabMenu: () => void;
}

export const createSheetTabs = (ctx: SheetTabsCtx): SheetTabsApi => {
  const tabsList = document.getElementById('sheet-tabs');
  const tabAddBtn = document.getElementById('btn-sheet-add');
  const tabPrevBtn = document.getElementById('btn-sheet-prev');
  const tabNextBtn = document.getElementById('btn-sheet-next');

  let tabMenuEl: HTMLDivElement | null = null;

  const closeTabMenu = (): void => {
    if (!tabMenuEl) return;
    tabMenuEl.remove();
    tabMenuEl = null;
  };

  const renderSheetTabs = (): void => {
    const inst = ctx.getInst();
    if (!inst || !tabsList) return;
    const wb = inst.workbook;
    const state = inst.store.getState();
    const activeIdx = state.data.sheetIndex;
    const hidden = state.layout.hiddenSheets;
    const n = wb.sheetCount;
    tabsList.replaceChildren();
    for (let i = 0; i < n; i += 1) {
      if (hidden.has(i)) continue;
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'app__tab';
      if (i === activeIdx) btn.classList.add('app__tab--active');
      btn.setAttribute('role', 'tab');
      btn.setAttribute('aria-selected', i === activeIdx ? 'true' : 'false');
      const tabColor = state.layout.sheetTabColors.get(i);
      if (tabColor) {
        btn.dataset.sheetTabColor = 'true';
        btn.style.setProperty('--app-sheet-tab-color', tabColor);
      }
      const label = document.createElement('span');
      label.className = 'app__tab-label';
      label.textContent = wb.sheetName(i);
      btn.appendChild(label);
      btn.addEventListener('click', () => switchSheet(i));
      btn.addEventListener('contextmenu', (e) => {
        e.preventDefault();
        openTabMenu(i, e.clientX, e.clientY);
      });
      tabsList.appendChild(btn);
    }
    // "Unhide…" affordance — surfaced as an extra tab pill when at least one
    // sheet is hidden. Click opens a list of hidden sheets to restore.
    if (hidden.size > 0) {
      const unhide = document.createElement('button');
      unhide.type = 'button';
      unhide.className = 'app__tab app__tab--unhide';
      unhide.textContent = `Unhide… (${hidden.size})`;
      unhide.addEventListener('click', (e) => {
        const r = (e.currentTarget as HTMLElement).getBoundingClientRect();
        openUnhideMenu(r.left, r.bottom);
      });
      tabsList.appendChild(unhide);
    }
  };

  const openUnhideMenu = (x: number, y: number): void => {
    const inst = ctx.getInst();
    if (!inst) return;
    closeTabMenu();
    const wb = inst.workbook;
    const store = inst.store;
    const hidden = store.getState().layout.hiddenSheets;
    if (hidden.size === 0) return;

    const menu = document.createElement('div');
    menu.className = 'app__menu';
    prepareMenu(menu, 'Unhide sheet');
    menu.style.position = 'fixed';
    menu.style.left = `${x}px`;
    menu.style.top = `${y}px`;
    menu.style.zIndex = '90';
    let cleanupMenuListeners = (): void => {};

    for (const i of Array.from(hidden).sort((a, b) => a - b)) {
      const it = document.createElement('button');
      it.type = 'button';
      it.className = 'app__menu-item';
      it.setAttribute('role', 'menuitem');
      it.tabIndex = -1;
      it.textContent = wb.sheetName(i);
      it.addEventListener('click', () => {
        closeTabMenu();
        cleanupMenuListeners();
        if (setSheetHidden(store, wb, inst?.history ?? null, i, false)) {
          renderSheetTabs();
        }
      });
      menu.appendChild(it);
    }

    document.body.appendChild(menu);
    tabMenuEl = menu;
    focusMenuItem(menu);

    const rect = menu.getBoundingClientRect();
    if (rect.right > window.innerWidth) {
      menu.style.left = `${Math.max(0, window.innerWidth - rect.width - 4)}px`;
    }
    if (rect.bottom > window.innerHeight) {
      menu.style.top = `${Math.max(0, window.innerHeight - rect.height - 4)}px`;
    }

    const onDocDown = (ev: MouseEvent): void => {
      if (!tabMenuEl) return;
      if (ev.target instanceof Node && tabMenuEl.contains(ev.target)) return;
      closeTabMenu();
      cleanupMenuListeners();
    };
    const onDocKey = (ev: KeyboardEvent): void => {
      handleMenuKeydown(ev, menu, {
        close: (restoreFocus) => {
          closeTabMenu();
          cleanupMenuListeners();
          if (restoreFocus) {
            document.querySelector<HTMLButtonElement>('.app__tab--unhide')?.focus();
          }
        },
      });
    };
    cleanupMenuListeners = () => {
      document.removeEventListener('mousedown', onDocDown, true);
      document.removeEventListener('keydown', onDocKey, true);
    };
    document.addEventListener('mousedown', onDocDown, true);
    document.addEventListener('keydown', onDocKey, true);
  };

  const openTabMenu = (idx: number, x: number, y: number): void => {
    const inst = ctx.getInst();
    if (!inst) return;
    openSheetTabMenu({
      closeTabMenu,
      idx,
      inst,
      renderSheetTabs,
      setTabMenuEl: (el) => {
        tabMenuEl = el;
      },
      x,
      y,
    });
  };

  const switchSheet = (idx: number): void => {
    const inst = ctx.getInst();
    if (!inst) return;
    const n = inst.workbook.sheetCount;
    if (idx < 0 || idx >= n) return;
    if (inst.store.getState().data.sheetIndex === idx) return;
    mutators.setSheetIndex(inst.store, idx);
    mutators.replaceCells(inst.store, inst.workbook.cells(idx));
    renderSheetTabs();
    ctx.focusSheet();
  };

  tabAddBtn?.addEventListener('click', () => {
    const inst = ctx.getInst();
    if (!inst) return;
    const idx = addSheet(inst.store, inst.workbook);
    if (idx < 0) {
      if (ctx.statusMetric && isWorkbookStructureProtected(inst.store.getState())) {
        ctx.statusMetric.textContent = ctx.workbookStructureProtectedBlockedText;
      }
      return;
    }
    // The wb.subscribe handler in mount.ts will pick up sheet-add as a no-op for cells,
    // but we re-render tabs and switch to the new sheet here.
    renderSheetTabs();
    switchSheet(idx);
  });

  tabPrevBtn?.addEventListener('click', () => {
    const inst = ctx.getInst();
    if (!inst) return;
    switchSheet(inst.store.getState().data.sheetIndex - 1);
  });
  tabNextBtn?.addEventListener('click', () => {
    const inst = ctx.getInst();
    if (!inst) return;
    switchSheet(inst.store.getState().data.sheetIndex + 1);
  });

  return {
    renderSheetTabs,
    switchSheet,
    openTabMenu,
    openUnhideMenu,
    closeTabMenu,
  };
};
