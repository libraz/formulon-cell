import {
  colLetter,
  focusMenuItem,
  handleMenuKeydown,
  type SpreadsheetInstance,
} from '@libraz/formulon-cell';

export interface HomeMenuWiringsCtx {
  getInst: () => SpreadsheetInstance | null;
  // Sibling menu close callbacks owned by other modules / main.ts.
  closeBorderMenu: (restoreFocus?: boolean) => void;
  closeFreezeMenu: (restoreFocus?: boolean) => void;
  closeFindSelectMenu: (restoreFocus?: boolean) => void;
  closeSortFilterHomeMenu: (restoreFocus?: boolean) => void;
  // Action handlers consumed by these menus.
  pasteClipboardIntoSelection: () => Promise<void>;
  applyRibbonPasteAction: (action: string) => Promise<void>;
  applyConditionalMenuAction: (action: string, panel?: string) => Promise<void>;
  applyFillSeries: (modeOverride?: 'days' | 'weekdays' | 'months' | 'years') => Promise<void>;
  applyFillDirection: (direction: 'down' | 'right' | 'up' | 'left') => void;
  applyClearAction: (action: string) => void;
  applyTextOrientationAction: (action: string) => void;
  applyCellInsertAction: (action: string) => Promise<void>;
  applyCellDeleteAction: (action: string) => Promise<void>;
  applyCellFormatAction: (action: string) => Promise<void>;
  applyPrintAreaAction: (action: 'set' | 'clear') => void;
  insertSymbolIntoActiveCell: (symbol: string) => void;
  insertCustomSymbolIntoActiveCell: () => Promise<void>;
}

export interface HomeMenuWiringsApi {
  closePasteMenu: (restoreFocus?: boolean) => void;
  closeConditionalMenu: (restoreFocus?: boolean) => void;
  closeFillMenu: (restoreFocus?: boolean) => void;
  closeClearMenu: (restoreFocus?: boolean) => void;
  closePrintAreaMenu: (restoreFocus?: boolean) => void;
  closeSymbolMenu: (restoreFocus?: boolean) => void;
  closeTextOrientationMenu: (restoreFocus?: boolean) => void;
  closeCellsMenus: (restoreFocusTo?: HTMLElement | null) => void;
  openPrintAreaMenu: (printAreaBtn?: HTMLButtonElement | null) => void;
  openSymbolMenu: (symbolBtn?: HTMLButtonElement | null) => void;
  getPrintAreaMenu: () => HTMLDivElement | null;
  getSymbolMenu: () => HTMLDivElement | null;
  selectionToA1Range: () => string | null;
}

export const createHomeMenuWirings = (ctx: HomeMenuWiringsCtx): HomeMenuWiringsApi => {
  const {
    getInst,
    closeBorderMenu,
    closeFreezeMenu,
    closeFindSelectMenu,
    closeSortFilterHomeMenu,
    pasteClipboardIntoSelection,
    applyRibbonPasteAction,
    applyConditionalMenuAction,
    applyFillSeries,
    applyFillDirection,
    applyClearAction,
    applyTextOrientationAction,
    applyCellInsertAction,
    applyCellDeleteAction,
    applyCellFormatAction,
    applyPrintAreaAction,
    insertSymbolIntoActiveCell,
    insertCustomSymbolIntoActiveCell,
  } = ctx;

  const pasteBtn = document.querySelector<HTMLButtonElement>('button[data-ribbon-command="paste"]');
  const pasteMenu = document.getElementById('menu-paste') as HTMLDivElement | null;
  pasteBtn?.setAttribute('aria-haspopup', 'menu');
  pasteBtn?.setAttribute('aria-expanded', 'false');
  const closePasteMenu = (restoreFocus = false): void => {
    if (!pasteMenu) return;
    pasteMenu.hidden = true;
    pasteBtn?.setAttribute('aria-expanded', 'false');
    if (restoreFocus) pasteBtn?.focus();
  };
  const openPasteMenu = (): void => {
    if (!pasteMenu || !pasteBtn) return;
    closeBorderMenu();
    closeFreezeMenu();
    closeConditionalMenu();
    closeTextOrientationMenu();
    closeFillMenu();
    closeClearMenu();
    closeCellsMenus();
    closeSortFilterHomeMenu();
    closeFindSelectMenu();
    pasteMenu.hidden = false;
    pasteBtn.setAttribute('aria-haspopup', 'menu');
    pasteBtn.setAttribute('aria-expanded', 'true');
    focusMenuItem(pasteMenu, 'first');
  };
  pasteBtn?.addEventListener('click', (event) => {
    event.preventDefault();
    event.stopPropagation();
    const target = event.target as Element | null;
    if (event.altKey || event.shiftKey || target?.closest('.demo__rb-split-chevron')) {
      if (!pasteMenu) return;
      if (pasteMenu.hidden) openPasteMenu();
      else closePasteMenu(true);
      return;
    }
    closePasteMenu();
    void pasteClipboardIntoSelection();
  });
  pasteBtn?.addEventListener('keydown', (event) => {
    if (event.key !== 'ArrowDown') return;
    event.preventDefault();
    event.stopPropagation();
    if (pasteMenu?.hidden) openPasteMenu();
  });
  pasteMenu?.addEventListener('click', (event) => {
    const item = (event.target as Element | null)?.closest<HTMLButtonElement>(
      '[data-paste-action]',
    );
    const action = item?.dataset.pasteAction;
    if (!action) return;
    event.preventDefault();
    event.stopPropagation();
    closePasteMenu();
    void applyRibbonPasteAction(action);
  });
  pasteMenu?.addEventListener('keydown', (event) => {
    handleMenuKeydown(event, pasteMenu, { close: closePasteMenu, restoreFocusTo: pasteBtn });
  });
  document.addEventListener('mousedown', (event) => {
    const target = event.target as Node | null;
    if (
      pasteMenu?.hidden === false &&
      target &&
      !pasteMenu.contains(target) &&
      !pasteBtn?.contains(target)
    ) {
      closePasteMenu();
    }
  });
  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && pasteMenu?.hidden === false) closePasteMenu(true);
  });

  const conditionalBtn = document.querySelector<HTMLButtonElement>(
    'button[data-ribbon-command="conditional"]',
  );
  const conditionalMenu = document.getElementById('menu-conditional') as HTMLDivElement | null;
  let conditionalSubmenuCloseTimer: number | null = null;
  const cancelConditionalSubmenuClose = (): void => {
    if (conditionalSubmenuCloseTimer === null) return;
    window.clearTimeout(conditionalSubmenuCloseTimer);
    conditionalSubmenuCloseTimer = null;
  };
  const scheduleConditionalSubmenuClose = (): void => {
    cancelConditionalSubmenuClose();
    conditionalSubmenuCloseTimer = window.setTimeout(() => {
      closeConditionalSubmenus();
    }, 180);
  };
  const closeConditionalSubmenus = (): void => {
    cancelConditionalSubmenuClose();
    conditionalMenu?.querySelectorAll<HTMLElement>('[data-cf-panel]').forEach((panel) => {
      panel.hidden = true;
    });
    conditionalMenu?.querySelectorAll<HTMLElement>('[data-cf-submenu]').forEach((trigger) => {
      trigger.classList.remove('app__menu-item--active');
    });
  };
  const closeConditionalMenu = (restoreFocus = false): void => {
    if (!conditionalMenu) return;
    conditionalMenu.hidden = true;
    closeConditionalSubmenus();
    conditionalBtn?.setAttribute('aria-expanded', 'false');
    if (restoreFocus) conditionalBtn?.focus();
  };
  const openConditionalMenu = (): void => {
    if (!conditionalMenu || !conditionalBtn) return;
    closeBorderMenu();
    closeFreezeMenu();
    closeTextOrientationMenu();
    closeFillMenu();
    closeClearMenu();
    conditionalMenu.hidden = false;
    conditionalBtn.setAttribute('aria-haspopup', 'menu');
    conditionalBtn.setAttribute('aria-expanded', 'true');
    focusMenuItem(conditionalMenu, 'first');
  };
  const openConditionalSubmenu = (key: string, trigger: HTMLElement): void => {
    if (!conditionalMenu) return;
    cancelConditionalSubmenuClose();
    closeConditionalSubmenus();
    const panel = conditionalMenu.querySelector<HTMLElement>(`[data-cf-panel="${key}"]`);
    if (!panel) return;
    const menuRect = conditionalMenu.getBoundingClientRect();
    const triggerRect = trigger.getBoundingClientRect();
    panel.style.top = `${Math.max(0, triggerRect.top - menuRect.top - 4)}px`;
    panel.hidden = false;
    trigger.classList.add('app__menu-item--active');
  };

  conditionalBtn?.addEventListener('click', (event) => {
    event.preventDefault();
    event.stopPropagation();
    if (!conditionalMenu) return;
    if (conditionalMenu.hidden) openConditionalMenu();
    else closeConditionalMenu(true);
  });
  conditionalMenu?.addEventListener('click', (event) => {
    const target = event.target as Element | null;
    const submenu = target?.closest<HTMLElement>('[data-cf-submenu]');
    if (submenu) {
      event.preventDefault();
      event.stopPropagation();
      openConditionalSubmenu(submenu.dataset.cfSubmenu ?? '', submenu);
      return;
    }
    const item = target?.closest<HTMLButtonElement>('[data-cf-action]');
    const action = item?.dataset.cfAction;
    if (!item || !action || action.startsWith('submenu-')) return;
    event.preventDefault();
    event.stopPropagation();
    const panel = item.closest<HTMLElement>('[data-cf-panel]')?.dataset.cfPanel;
    closeConditionalMenu();
    void applyConditionalMenuAction(action, panel);
  });
  conditionalMenu?.querySelectorAll<HTMLElement>('[data-cf-submenu]').forEach((trigger) => {
    trigger.addEventListener('mouseenter', () =>
      openConditionalSubmenu(trigger.dataset.cfSubmenu ?? '', trigger),
    );
  });
  conditionalMenu
    ?.querySelectorAll<HTMLElement>('.app__menu-item:not([data-cf-submenu])')
    .forEach((item) => {
      item.addEventListener('mouseenter', scheduleConditionalSubmenuClose);
    });
  conditionalMenu?.querySelectorAll<HTMLElement>('[data-cf-panel]').forEach((panel) => {
    panel.addEventListener('mouseenter', cancelConditionalSubmenuClose);
    panel.addEventListener('mouseleave', scheduleConditionalSubmenuClose);
  });
  document.addEventListener('click', (event) => {
    if (conditionalMenu?.hidden !== false) return;
    const target = event.target as Node | null;
    if (target && (conditionalMenu.contains(target) || conditionalBtn?.contains(target))) return;
    closeConditionalMenu();
  });
  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && conditionalMenu?.hidden === false) closeConditionalMenu(true);
  });

  const fillBtn = document.querySelector<HTMLButtonElement>(
    'button[data-ribbon-command="fillHome"]',
  );
  const fillMenu = document.getElementById('menu-fill') as HTMLDivElement | null;
  const closeFillMenu = (restoreFocus = false): void => {
    if (!fillMenu) return;
    fillMenu.hidden = true;
    fillBtn?.setAttribute('aria-expanded', 'false');
    if (restoreFocus) fillBtn?.focus();
  };
  const openFillMenu = (): void => {
    if (!fillMenu || !fillBtn) return;
    closeBorderMenu();
    closeFreezeMenu();
    closeConditionalMenu();
    closeTextOrientationMenu();
    closeClearMenu();
    fillMenu.hidden = false;
    fillBtn.setAttribute('aria-haspopup', 'menu');
    fillBtn.setAttribute('aria-expanded', 'true');
    focusMenuItem(fillMenu, 'first');
  };
  fillBtn?.addEventListener('click', (event) => {
    event.preventDefault();
    event.stopPropagation();
    if (!fillMenu) return;
    if (fillMenu.hidden) openFillMenu();
    else closeFillMenu(true);
  });
  fillMenu?.addEventListener('click', (event) => {
    const item = (event.target as Element | null)?.closest<HTMLButtonElement>('[data-fill]');
    const action = item?.dataset.fill;
    if (!action) return;
    event.preventDefault();
    event.stopPropagation();
    closeFillMenu();
    if (action === 'series') {
      void applyFillSeries();
      return;
    }
    if (action === 'days' || action === 'weekdays' || action === 'months' || action === 'years') {
      void applyFillSeries(action);
      return;
    }
    applyFillDirection(action as 'down' | 'right' | 'up' | 'left');
  });
  fillMenu?.addEventListener('keydown', (event) => {
    handleMenuKeydown(event, fillMenu, { close: closeFillMenu, restoreFocusTo: fillBtn });
  });

  const clearBtn = document.querySelector<HTMLButtonElement>(
    '.demo__ribbon-group--editing button[data-ribbon-command="clearFormat"]',
  );
  const clearMenu = document.getElementById('menu-clear') as HTMLDivElement | null;
  const closeClearMenu = (restoreFocus = false): void => {
    if (!clearMenu) return;
    clearMenu.hidden = true;
    clearBtn?.setAttribute('aria-expanded', 'false');
    if (restoreFocus) clearBtn?.focus();
  };
  const openClearMenu = (): void => {
    if (!clearMenu || !clearBtn) return;
    closeBorderMenu();
    closeFreezeMenu();
    closeConditionalMenu();
    closeTextOrientationMenu();
    closeFillMenu();
    clearMenu.hidden = false;
    clearBtn.setAttribute('aria-haspopup', 'menu');
    clearBtn.setAttribute('aria-expanded', 'true');
    focusMenuItem(clearMenu, 'first');
  };
  clearBtn?.addEventListener('click', (event) => {
    event.preventDefault();
    event.stopPropagation();
    if (!clearMenu) return;
    if (clearMenu.hidden) openClearMenu();
    else closeClearMenu(true);
  });
  clearMenu?.addEventListener('click', (event) => {
    const item = (event.target as Element | null)?.closest<HTMLButtonElement>('[data-clear]');
    const action = item?.dataset.clear;
    if (!action) return;
    event.preventDefault();
    event.stopPropagation();
    closeClearMenu();
    applyClearAction(action);
  });
  clearMenu?.addEventListener('keydown', (event) => {
    handleMenuKeydown(event, clearMenu, { close: closeClearMenu, restoreFocusTo: clearBtn });
  });

  const getPrintAreaBtn = (): HTMLButtonElement | null =>
    document.querySelector<HTMLButtonElement>('button[data-ribbon-command="printArea"]');
  const getPrintAreaMenu = (): HTMLDivElement | null =>
    document.getElementById('menu-print-area') as HTMLDivElement | null;
  const getSymbolBtn = (): HTMLButtonElement | null =>
    document.querySelector<HTMLButtonElement>('button[data-ribbon-command="symbolInsert"]');
  const getSymbolMenu = (): HTMLDivElement | null =>
    document.getElementById('menu-symbol') as HTMLDivElement | null;
  const closePrintAreaMenu = (restoreFocus = false): void => {
    const printAreaMenu = getPrintAreaMenu();
    const printAreaBtn = getPrintAreaBtn();
    if (!printAreaMenu) return;
    printAreaMenu.hidden = true;
    printAreaBtn?.setAttribute('aria-expanded', 'false');
    if (restoreFocus) printAreaBtn?.focus();
  };
  const closeSymbolMenu = (restoreFocus = false): void => {
    const symbolMenu = getSymbolMenu();
    const symbolBtn = getSymbolBtn();
    if (!symbolMenu) return;
    symbolMenu.hidden = true;
    symbolBtn?.setAttribute('aria-expanded', 'false');
    if (restoreFocus) symbolBtn?.focus();
  };
  const openPrintAreaMenu = (printAreaBtn: HTMLButtonElement | null = getPrintAreaBtn()): void => {
    const printAreaMenu = getPrintAreaMenu();
    if (!printAreaMenu || !printAreaBtn) return;
    closeBorderMenu();
    closeFreezeMenu();
    closeConditionalMenu();
    closeTextOrientationMenu();
    closeFillMenu();
    closeClearMenu();
    printAreaMenu.hidden = false;
    printAreaBtn.setAttribute('aria-haspopup', 'menu');
    printAreaBtn.setAttribute('aria-expanded', 'true');
    focusMenuItem(printAreaMenu, 'first');
  };
  const openSymbolMenu = (symbolBtn: HTMLButtonElement | null = getSymbolBtn()): void => {
    const symbolMenu = getSymbolMenu();
    if (!symbolMenu || !symbolBtn) return;
    closeBorderMenu();
    closeFreezeMenu();
    closeConditionalMenu();
    closeTextOrientationMenu();
    closeFillMenu();
    closeClearMenu();
    closePrintAreaMenu();
    symbolMenu.hidden = false;
    symbolBtn.setAttribute('aria-haspopup', 'menu');
    symbolBtn.setAttribute('aria-expanded', 'true');
    focusMenuItem(symbolMenu, 'first');
  };
  document.addEventListener('click', (event) => {
    const item = (event.target as Element | null)?.closest<HTMLButtonElement>(
      '[data-print-area-action]',
    );
    const action = item?.dataset.printAreaAction;
    if (!action) return;
    event.preventDefault();
    event.stopPropagation();
    closePrintAreaMenu();
    applyPrintAreaAction(action as 'set' | 'clear');
  });
  document.addEventListener('click', (event) => {
    const actionItem = (event.target as Element | null)?.closest<HTMLButtonElement>(
      '[data-symbol-action]',
    );
    if (actionItem?.dataset.symbolAction === 'more') {
      event.preventDefault();
      event.stopPropagation();
      closeSymbolMenu();
      void insertCustomSymbolIntoActiveCell();
      return;
    }
    const item = (event.target as Element | null)?.closest<HTMLButtonElement>('[data-symbol]');
    const symbol = item?.dataset.symbol;
    if (!symbol) return;
    event.preventDefault();
    event.stopPropagation();
    closeSymbolMenu();
    insertSymbolIntoActiveCell(symbol);
  });
  document.addEventListener('keydown', (event) => {
    const printAreaMenu = getPrintAreaMenu();
    if (!printAreaMenu || printAreaMenu.hidden) return;
    handleMenuKeydown(event, printAreaMenu, {
      close: closePrintAreaMenu,
      restoreFocusTo: getPrintAreaBtn(),
    });
  });
  document.addEventListener('keydown', (event) => {
    const symbolMenu = getSymbolMenu();
    if (!symbolMenu || symbolMenu.hidden) return;
    handleMenuKeydown(event, symbolMenu, {
      close: closeSymbolMenu,
      restoreFocusTo: getSymbolBtn(),
    });
  });
  document.addEventListener('mousedown', (event) => {
    const target = event.target as Node | null;
    const printAreaMenu = getPrintAreaMenu();
    const printAreaBtn = getPrintAreaBtn();
    const symbolMenu = getSymbolMenu();
    const symbolBtn = getSymbolBtn();
    if (
      fillMenu?.hidden === false &&
      target &&
      !fillMenu.contains(target) &&
      !fillBtn?.contains(target)
    ) {
      closeFillMenu();
    }
    if (
      clearMenu?.hidden === false &&
      target &&
      !clearMenu.contains(target) &&
      !clearBtn?.contains(target)
    ) {
      closeClearMenu();
    }
    if (
      printAreaMenu?.hidden === false &&
      target &&
      !printAreaMenu.contains(target) &&
      !printAreaBtn?.contains(target)
    ) {
      closePrintAreaMenu();
    }
    if (
      symbolMenu?.hidden === false &&
      target &&
      !symbolMenu.contains(target) &&
      !symbolBtn?.contains(target)
    ) {
      closeSymbolMenu();
    }
  });
  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && fillMenu?.hidden === false) closeFillMenu(true);
    if (event.key === 'Escape' && clearMenu?.hidden === false) closeClearMenu(true);
    if (event.key === 'Escape' && getPrintAreaMenu()?.hidden === false) closePrintAreaMenu(true);
    if (event.key === 'Escape' && getSymbolMenu()?.hidden === false) closeSymbolMenu(true);
  });

  const textOrientationBtn = document.querySelector<HTMLButtonElement>(
    'button[data-ribbon-command="textOrientation"]',
  );
  const textOrientationMenu = document.getElementById(
    'menu-text-orientation',
  ) as HTMLDivElement | null;
  const closeTextOrientationMenu = (restoreFocus = false): void => {
    if (!textOrientationMenu) return;
    textOrientationMenu.hidden = true;
    textOrientationBtn?.setAttribute('aria-expanded', 'false');
    if (restoreFocus) textOrientationBtn?.focus();
  };
  const openTextOrientationMenu = (): void => {
    if (!textOrientationMenu || !textOrientationBtn) return;
    closeBorderMenu();
    closeFreezeMenu();
    closeConditionalMenu();
    closeFillMenu();
    closeClearMenu();
    closeTextOrientationMenu();
    closeCellsMenus();
    textOrientationMenu.hidden = false;
    textOrientationBtn.setAttribute('aria-haspopup', 'menu');
    textOrientationBtn.setAttribute('aria-expanded', 'true');
    focusMenuItem(textOrientationMenu, 'first');
  };
  textOrientationBtn?.addEventListener('click', (event) => {
    event.preventDefault();
    event.stopPropagation();
    if (!textOrientationMenu) return;
    if (textOrientationMenu.hidden) openTextOrientationMenu();
    else closeTextOrientationMenu(true);
  });
  textOrientationMenu?.addEventListener('click', (event) => {
    const item = (event.target as Element | null)?.closest<HTMLButtonElement>(
      '[data-text-orientation]',
    );
    const action = item?.dataset.textOrientation;
    if (!action) return;
    event.preventDefault();
    event.stopPropagation();
    closeTextOrientationMenu();
    applyTextOrientationAction(action);
  });
  textOrientationMenu?.addEventListener('keydown', (event) => {
    handleMenuKeydown(event, textOrientationMenu, {
      close: closeTextOrientationMenu,
      restoreFocusTo: textOrientationBtn,
    });
  });
  document.addEventListener('mousedown', (event) => {
    const target = event.target as Node | null;
    if (
      textOrientationMenu?.hidden === false &&
      target &&
      !textOrientationMenu.contains(target) &&
      !textOrientationBtn?.contains(target)
    ) {
      closeTextOrientationMenu();
    }
  });
  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && textOrientationMenu?.hidden === false)
      closeTextOrientationMenu(true);
  });

  const insertCellsBtn = document.querySelector<HTMLButtonElement>(
    'button[data-ribbon-command="insertRows"]',
  );
  const deleteCellsBtn = document.querySelector<HTMLButtonElement>(
    'button[data-ribbon-command="deleteRows"]',
  );
  const formatCellsBtn = document.querySelector<HTMLButtonElement>(
    'button[data-ribbon-command="formatCellsHome"]',
  );
  const insertCellsMenu = document.getElementById('menu-insert-cells') as HTMLDivElement | null;
  const deleteCellsMenu = document.getElementById('menu-delete-cells') as HTMLDivElement | null;
  const formatCellsMenu = document.getElementById('menu-format-cells') as HTMLDivElement | null;

  const closeCellsMenus = (restoreFocusTo?: HTMLElement | null): void => {
    for (const [menu, btn] of [
      [insertCellsMenu, insertCellsBtn],
      [deleteCellsMenu, deleteCellsBtn],
      [formatCellsMenu, formatCellsBtn],
    ] as const) {
      if (!menu) continue;
      menu.hidden = true;
      btn?.setAttribute('aria-expanded', 'false');
    }
    restoreFocusTo?.focus();
  };

  const openCellsMenu = (menu: HTMLDivElement | null, button: HTMLButtonElement | null): void => {
    if (!menu || !button) return;
    closeBorderMenu();
    closeFreezeMenu();
    closeConditionalMenu();
    closeFillMenu();
    closeClearMenu();
    closeCellsMenus();
    menu.hidden = false;
    button.setAttribute('aria-haspopup', 'menu');
    button.setAttribute('aria-expanded', 'true');
    focusMenuItem(menu, 'first');
  };

  insertCellsBtn?.addEventListener('click', (event) => {
    event.preventDefault();
    event.stopPropagation();
    if (!insertCellsMenu) return;
    if (insertCellsMenu.hidden) openCellsMenu(insertCellsMenu, insertCellsBtn);
    else closeCellsMenus(insertCellsBtn);
  });
  deleteCellsBtn?.addEventListener('click', (event) => {
    event.preventDefault();
    event.stopPropagation();
    if (!deleteCellsMenu) return;
    if (deleteCellsMenu.hidden) openCellsMenu(deleteCellsMenu, deleteCellsBtn);
    else closeCellsMenus(deleteCellsBtn);
  });
  formatCellsBtn?.addEventListener('click', (event) => {
    event.preventDefault();
    event.stopPropagation();
    if (!formatCellsMenu) return;
    if (formatCellsMenu.hidden) openCellsMenu(formatCellsMenu, formatCellsBtn);
    else closeCellsMenus(formatCellsBtn);
  });

  insertCellsMenu?.addEventListener('click', (event) => {
    const item = (event.target as Element | null)?.closest<HTMLButtonElement>('[data-cell-insert]');
    const action = item?.dataset.cellInsert;
    if (!action) return;
    event.preventDefault();
    event.stopPropagation();
    closeCellsMenus();
    void applyCellInsertAction(action);
  });
  deleteCellsMenu?.addEventListener('click', (event) => {
    const item = (event.target as Element | null)?.closest<HTMLButtonElement>('[data-cell-delete]');
    const action = item?.dataset.cellDelete;
    if (!action) return;
    event.preventDefault();
    event.stopPropagation();
    closeCellsMenus();
    void applyCellDeleteAction(action);
  });
  formatCellsMenu?.addEventListener('click', (event) => {
    const item = (event.target as Element | null)?.closest<HTMLButtonElement>('[data-cell-format]');
    const action = item?.dataset.cellFormat;
    if (!action) return;
    event.preventDefault();
    event.stopPropagation();
    closeCellsMenus();
    void applyCellFormatAction(action);
  });
  insertCellsMenu?.addEventListener('keydown', (event) => {
    handleMenuKeydown(event, insertCellsMenu, {
      close: () => closeCellsMenus(insertCellsBtn),
      restoreFocusTo: insertCellsBtn,
    });
  });
  deleteCellsMenu?.addEventListener('keydown', (event) => {
    handleMenuKeydown(event, deleteCellsMenu, {
      close: () => closeCellsMenus(deleteCellsBtn),
      restoreFocusTo: deleteCellsBtn,
    });
  });
  formatCellsMenu?.addEventListener('keydown', (event) => {
    handleMenuKeydown(event, formatCellsMenu, {
      close: () => closeCellsMenus(formatCellsBtn),
      restoreFocusTo: formatCellsBtn,
    });
  });
  document.addEventListener('mousedown', (event) => {
    const target = event.target as Node | null;
    if (!target) return;
    const inside =
      insertCellsMenu?.contains(target) ||
      deleteCellsMenu?.contains(target) ||
      formatCellsMenu?.contains(target) ||
      insertCellsBtn?.contains(target) ||
      deleteCellsBtn?.contains(target) ||
      formatCellsBtn?.contains(target);
    if (!inside) closeCellsMenus();
  });
  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape') closeCellsMenus();
  });

  const selectionToA1Range = (): string | null => {
    const i = getInst();
    if (!i) return null;
    const r = i.store.getState().selection.range;
    const start = `${colLetter(r.c0)}${r.r0 + 1}`;
    const end = `${colLetter(r.c1)}${r.r1 + 1}`;
    return start === end ? start : `${start}:${end}`;
  };

  return {
    closePasteMenu,
    closeConditionalMenu,
    closeFillMenu,
    closeClearMenu,
    closePrintAreaMenu,
    closeSymbolMenu,
    closeTextOrientationMenu,
    closeCellsMenus,
    openPrintAreaMenu,
    openSymbolMenu,
    getPrintAreaMenu,
    getSymbolMenu,
    selectionToA1Range,
  };
};
