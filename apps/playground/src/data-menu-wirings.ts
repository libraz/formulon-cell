// Data-tab and adjacent ribbon menu wirings extracted from main.ts.
//
// Bundles the dropdown wirings for the Home/Data sort-filter buttons, the
// find-&-select menu, the AutoSum split-button menus (used on both Home and
// Formulas), the calc-options menu, the chart-insert menu, and the freeze
// menu. They were grouped here because they share the same close/open
// boilerplate and because main.ts dropped below the 5k-line threshold once
// these blocks were lifted out.
//
// The factory takes a context with live getters for the spreadsheet instance
// plus cross-tab menu closers and action handlers that still live in main.ts.
// Behavior is preserved verbatim — the function signatures, DOM ids, event
// types, and side-effect ordering all match the original inline code.
//
// External callers of the returned API:
//   - applySortMenuAction / applyCalcOptionAction / applyFreezeMenuAction are
//     dispatched from DYNAMIC_DROPDOWN_HANDLERS.
//   - updateCalcOptionsMenu is invoked when openDynamicRibbonDropdown opens
//     the calc-options menu so it can refresh the active-mode indicator.
//   - closeFreezeMenu / closeSortFilterHomeMenu / closeFindSelectMenu /
//     closeCalcOptionsMenu / closeChartInsertMenu are called by other menu
//     open helpers (border, paste, conditional, fill, clear, cells,
//     text-orientation, etc.) to keep only one dropdown open at a time.
//   - openFreezeMenu / closeFreezeMenu / getFreezeMenu / getFreezeBtn are
//     used by the id='freeze' branch of the generic ribbon click handler.

import type { AutoSumFormulaName } from '@libraz/formulon-cell';
import {
  clearFilter,
  filterBySelectedCellValue,
  inferAutoFilterRange,
  reapplyFilters,
  recordFilterChange,
  type SessionChartKind,
  type SpreadsheetInstance,
  setFreezePanes,
} from '@libraz/formulon-cell';

import { showMessage } from './dialogs.js';
import { focusMenuItem, handleMenuKeydown } from './menu-a11y.js';

export interface DataMenuWiringsCtx {
  getInst: () => SpreadsheetInstance | null;
  ribbonLang: 'ja' | 'en';
  getSheetEl: () => HTMLElement;
  focusSheet: () => void;
  refreshWorkbookCells: () => void;
  // Cross-tab menu closers (defined elsewhere in main.ts).
  closeBorderMenu: () => void;
  closeConditionalMenu: () => void;
  closeFillMenu: () => void;
  closeClearMenu: () => void;
  closeCellsMenus: () => void;
  closeTextOrientationMenu: () => void;
  closePasteMenu: () => void;
  // Action handlers consumed from main.ts.
  sortSelection: (direction: 'asc' | 'desc') => void;
  customSortSelection: () => void | Promise<void>;
  openFilterForSelection: () => void;
  applyAdvancedFilterAction: () => void | Promise<void>;
  removeDuplicateRows: () => void;
  applyFindSelectAction: (action: string) => void;
  applyAutoSumFormula: (fn?: AutoSumFormulaName) => void;
  createChartFromSelection: (kind: SessionChartKind) => void;
  createRecommendedChartFromSelection: () => void | Promise<void>;
  chartKindFromAction: (action: string) => SessionChartKind;
}

export interface DataMenuWiringsApi {
  applySortMenuAction: (action: string) => void;
  applyCalcOptionAction: (action: string) => void;
  applyFreezeMenuAction: (action: string | undefined) => void;
  updateCalcOptionsMenu: (menu?: HTMLElement) => void;
  openSortFilterHomeMenu: () => void;
  closeSortFilterHomeMenu: (restoreFocus?: boolean) => void;
  openDataSortMenu: () => void;
  closeDataSortMenu: (restoreFocus?: boolean) => void;
  openFindSelectMenu: () => void;
  closeFindSelectMenu: (restoreFocus?: boolean) => void;
  openCalcOptionsMenu: () => void;
  closeCalcOptionsMenu: (restoreFocus?: boolean) => void;
  openChartInsertMenu: () => void;
  closeChartInsertMenu: (restoreFocus?: boolean) => void;
  openFreezeMenu: () => void;
  closeFreezeMenu: (restoreFocus?: boolean) => void;
  getFreezeBtn: () => HTMLButtonElement | null;
  getFreezeMenu: () => HTMLDivElement | null;
}

export const createDataMenuWirings = (ctx: DataMenuWiringsCtx): DataMenuWiringsApi => {
  const {
    getInst,
    ribbonLang,
    getSheetEl,
    focusSheet,
    refreshWorkbookCells,
    closeBorderMenu,
    closeConditionalMenu,
    closeFillMenu,
    closeClearMenu,
    closeCellsMenus,
    closeTextOrientationMenu,
    closePasteMenu,
    sortSelection,
    customSortSelection,
    openFilterForSelection,
    applyAdvancedFilterAction,
    removeDuplicateRows,
    applyFindSelectAction,
    applyAutoSumFormula,
    createChartFromSelection,
    createRecommendedChartFromSelection,
    chartKindFromAction,
  } = ctx;

  // ── Sort / Filter (Home tab) ───────────────────────────────────────────
  const sortFilterHomeBtn = document.querySelector<HTMLButtonElement>(
    'button[data-ribbon-command="sortFilterHome"]',
  );
  const sortFilterHomeMenu = document.getElementById('menu-sort-home') as HTMLDivElement | null;
  const closeSortFilterHomeMenu = (restoreFocus = false): void => {
    if (!sortFilterHomeMenu) return;
    sortFilterHomeMenu.hidden = true;
    sortFilterHomeBtn?.setAttribute('aria-expanded', 'false');
    if (restoreFocus) sortFilterHomeBtn?.focus();
  };
  const openSortFilterHomeMenu = (): void => {
    if (!sortFilterHomeMenu || !sortFilterHomeBtn) return;
    closeBorderMenu();
    closeFreezeMenu();
    closeConditionalMenu();
    closeFillMenu();
    closeClearMenu();
    closeCellsMenus();
    closeTextOrientationMenu();
    closeFindSelectMenu();
    sortFilterHomeMenu.hidden = false;
    sortFilterHomeBtn.setAttribute('aria-haspopup', 'menu');
    sortFilterHomeBtn.setAttribute('aria-expanded', 'true');
    focusMenuItem(sortFilterHomeMenu, 'first');
  };
  const applySortMenuAction = (action: string): void => {
    const i = getInst();
    if (!i) return;
    if (action === 'asc' || action === 'desc') sortSelection(action);
    else if (action === 'custom') void customSortSelection();
    else if (action === 'filter') openFilterForSelection();
    else if (action === 'filter-by-value') {
      const state = i.store.getState();
      const range = state.ui.filterRange ?? inferAutoFilterRange(state);
      recordFilterChange(i.history, i.store, () => {
        filterBySelectedCellValue(i.store.getState(), i.store, range);
      });
      focusSheet();
    } else if (action === 'filter-clear') {
      const state = i.store.getState();
      const range = state.ui.filterRange ?? inferAutoFilterRange(state);
      recordFilterChange(i.history, i.store, () => {
        clearFilter(i.store.getState(), i.store, range);
      });
      focusSheet();
    } else if (action === 'filter-reapply') {
      recordFilterChange(i.history, i.store, () => {
        reapplyFilters(i.store.getState(), i.store);
      });
      focusSheet();
    } else if (action === 'filter-advanced') {
      void applyAdvancedFilterAction();
    } else if (action === 'dedupe') removeDuplicateRows();
    else if (action === 'conditional') i.openConditionalDialog();
    else if (action === 'named') i.openNamedRangeDialog();
  };
  sortFilterHomeBtn?.addEventListener('click', (event) => {
    event.preventDefault();
    event.stopPropagation();
    if (!sortFilterHomeMenu) return;
    if (sortFilterHomeMenu.hidden) openSortFilterHomeMenu();
    else closeSortFilterHomeMenu(true);
  });
  sortFilterHomeMenu?.addEventListener('click', (event) => {
    const item = (event.target as Element | null)?.closest<HTMLButtonElement>('[data-sort]');
    const action = item?.dataset.sort;
    if (!action || !getInst()) return;
    event.preventDefault();
    event.stopPropagation();
    closeSortFilterHomeMenu();
    applySortMenuAction(action);
  });
  sortFilterHomeMenu?.addEventListener('keydown', (event) => {
    handleMenuKeydown(event, sortFilterHomeMenu, {
      close: closeSortFilterHomeMenu,
      restoreFocusTo: sortFilterHomeBtn,
    });
  });
  document.addEventListener('mousedown', (event) => {
    const target = event.target as Node | null;
    if (
      sortFilterHomeMenu?.hidden === false &&
      target &&
      !sortFilterHomeMenu.contains(target) &&
      !sortFilterHomeBtn?.contains(target)
    ) {
      closeSortFilterHomeMenu();
    }
  });
  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && sortFilterHomeMenu?.hidden === false)
      closeSortFilterHomeMenu(true);
  });

  // ── Data tab sort/filter button ─────────────────────────────────────────
  const dataSortBtn = document.querySelector<HTMLButtonElement>(
    'button[data-ribbon-command="filter"]',
  );
  const dataSortMenu = document.getElementById('menu-sort') as HTMLDivElement | null;
  const closeDataSortMenu = (restoreFocus = false): void => {
    if (!dataSortMenu) return;
    dataSortMenu.hidden = true;
    dataSortBtn?.setAttribute('aria-expanded', 'false');
    if (restoreFocus) dataSortBtn?.focus();
  };
  const openDataSortMenu = (): void => {
    if (!dataSortMenu || !dataSortBtn) return;
    closeBorderMenu();
    closeFreezeMenu();
    closeConditionalMenu();
    closeFillMenu();
    closeClearMenu();
    closeCellsMenus();
    closeTextOrientationMenu();
    closeSortFilterHomeMenu();
    closeFindSelectMenu();
    closePasteMenu();
    dataSortMenu.hidden = false;
    dataSortBtn.setAttribute('aria-haspopup', 'menu');
    dataSortBtn.setAttribute('aria-expanded', 'true');
    focusMenuItem(dataSortMenu, 'first');
  };
  dataSortBtn?.addEventListener('click', (event) => {
    event.preventDefault();
    event.stopPropagation();
    if (!dataSortMenu) return;
    if (dataSortMenu.hidden) openDataSortMenu();
    else closeDataSortMenu(true);
  });
  dataSortMenu?.addEventListener('click', (event) => {
    const item = (event.target as Element | null)?.closest<HTMLButtonElement>('[data-sort]');
    const action = item?.dataset.sort;
    if (!action) return;
    event.preventDefault();
    event.stopPropagation();
    closeDataSortMenu();
    applySortMenuAction(action);
  });
  dataSortMenu?.addEventListener('keydown', (event) => {
    handleMenuKeydown(event, dataSortMenu, {
      close: closeDataSortMenu,
      restoreFocusTo: dataSortBtn,
    });
  });
  document.addEventListener('mousedown', (event) => {
    const target = event.target as Node | null;
    if (
      dataSortMenu?.hidden === false &&
      target &&
      !dataSortMenu.contains(target) &&
      !dataSortBtn?.contains(target)
    ) {
      closeDataSortMenu();
    }
  });
  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && dataSortMenu?.hidden === false) closeDataSortMenu(true);
  });

  // ── Find & Select ───────────────────────────────────────────────────────
  const findSelectBtn = document.querySelector<HTMLButtonElement>(
    'button[data-ribbon-command="findHome"]',
  );
  const findSelectMenu = document.getElementById('menu-find-select') as HTMLDivElement | null;
  const closeFindSelectMenu = (restoreFocus = false): void => {
    if (!findSelectMenu) return;
    findSelectMenu.hidden = true;
    findSelectBtn?.setAttribute('aria-expanded', 'false');
    if (restoreFocus) findSelectBtn?.focus();
  };
  const openFindSelectMenu = (): void => {
    if (!findSelectMenu || !findSelectBtn) return;
    closeBorderMenu();
    closeFreezeMenu();
    closeConditionalMenu();
    closeFillMenu();
    closeClearMenu();
    closeCellsMenus();
    closeTextOrientationMenu();
    closeSortFilterHomeMenu();
    findSelectMenu.hidden = false;
    findSelectBtn.setAttribute('aria-haspopup', 'menu');
    findSelectBtn.setAttribute('aria-expanded', 'true');
    focusMenuItem(findSelectMenu, 'first');
  };
  findSelectBtn?.addEventListener('click', (event) => {
    event.preventDefault();
    event.stopPropagation();
    if (!findSelectMenu) return;
    if (findSelectMenu.hidden) openFindSelectMenu();
    else closeFindSelectMenu(true);
  });
  findSelectMenu?.addEventListener('click', (event) => {
    const item = (event.target as Element | null)?.closest<HTMLButtonElement>('[data-find-select]');
    const action = item?.dataset.findSelect;
    if (!action) return;
    event.preventDefault();
    event.stopPropagation();
    closeFindSelectMenu();
    applyFindSelectAction(action);
  });
  findSelectMenu?.addEventListener('keydown', (event) => {
    handleMenuKeydown(event, findSelectMenu, {
      close: closeFindSelectMenu,
      restoreFocusTo: findSelectBtn,
    });
  });
  document.addEventListener('mousedown', (event) => {
    const target = event.target as Node | null;
    if (
      findSelectMenu?.hidden === false &&
      target &&
      !findSelectMenu.contains(target) &&
      !findSelectBtn?.contains(target)
    ) {
      closeFindSelectMenu();
    }
  });
  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && findSelectMenu?.hidden === false) closeFindSelectMenu(true);
  });

  // ── AutoSum split buttons (Home + Formulas) ────────────────────────────
  const setupAutoSumMenu = (command: 'autosum' | 'autosumFormula', menuId: string): void => {
    const button = document.querySelector<HTMLButtonElement>(
      `button[data-ribbon-command="${command}"]`,
    );
    const menu = document.getElementById(menuId) as HTMLDivElement | null;
    const close = (restoreFocus = false): void => {
      if (!menu) return;
      menu.hidden = true;
      button?.setAttribute('aria-expanded', 'false');
      if (restoreFocus) button?.focus();
    };
    const open = (): void => {
      if (!button || !menu) return;
      closeBorderMenu();
      closeFreezeMenu();
      closeConditionalMenu();
      closeFillMenu();
      closeClearMenu();
      closeCellsMenus();
      closeTextOrientationMenu();
      closeSortFilterHomeMenu();
      closeFindSelectMenu();
      menu.hidden = false;
      button.setAttribute('aria-haspopup', 'menu');
      button.setAttribute('aria-expanded', 'true');
      focusMenuItem(menu, 'first');
    };
    button?.addEventListener('click', (event) => {
      event.preventDefault();
      event.stopPropagation();
      const target = event.target as Element | null;
      if (!menu) return;
      if (event.altKey || event.shiftKey || target?.closest('.demo__rb-split-chevron')) {
        if (menu.hidden) open();
        else close(true);
        return;
      }
      close();
      applyAutoSumFormula('SUM');
    });
    button?.addEventListener('keydown', (event) => {
      if (event.key !== 'ArrowDown') return;
      event.preventDefault();
      event.stopPropagation();
      if (menu?.hidden) open();
    });
    menu?.addEventListener('click', (event) => {
      const item = (event.target as Element | null)?.closest<HTMLButtonElement>(
        '[data-autosum-fn]',
      );
      const fn = item?.dataset.autosumFn as AutoSumFormulaName | undefined;
      if (!fn) return;
      event.preventDefault();
      event.stopPropagation();
      close();
      applyAutoSumFormula(fn);
    });
    menu?.addEventListener('keydown', (event) => {
      handleMenuKeydown(event, menu, { close, restoreFocusTo: button });
    });
    document.addEventListener('mousedown', (event) => {
      const target = event.target as Node | null;
      if (menu?.hidden === false && target && !menu.contains(target) && !button?.contains(target)) {
        close();
      }
    });
    document.addEventListener('keydown', (event) => {
      if (event.key === 'Escape' && menu?.hidden === false) close(true);
    });
  };

  setupAutoSumMenu('autosum', 'menu-autosum-home');
  setupAutoSumMenu('autosumFormula', 'menu-autosum-formulas');

  // ── Calculation options ─────────────────────────────────────────────────
  const applyCalcOptionAction = (action: string): void => {
    const i = getInst();
    if (!i) return;
    if (action === 'auto' || action === 'manual' || action === 'auto-no-table') {
      const mode = action === 'auto' ? 0 : action === 'manual' ? 1 : 2;
      const ok = i.workbook.setCalcMode(mode as 0 | 1 | 2);
      if (!ok) {
        void showMessage({
          title: ribbonLang === 'ja' ? '計算方法' : 'Calculation Options',
          message:
            ribbonLang === 'ja'
              ? 'このエンジンでは計算モードの保存に対応していません。'
              : 'This engine does not support saving calculation mode.',
        });
        return;
      }
      focusSheet();
      updateCalcOptionsMenu();
      return;
    }
    if (action === 'calculate-now' || action === 'calculate-sheet') {
      i.recalc();
      refreshWorkbookCells();
      focusSheet();
      return;
    }
    if (action === 'iterative') {
      i.openIterativeDialog();
    }
  };

  const updateCalcOptionsMenu = (menu: HTMLElement = document.body): void => {
    const mode = getInst()?.workbook.calcMode();
    const active =
      mode === 0 ? 'auto' : mode === 1 ? 'manual' : mode === 2 ? 'auto-no-table' : null;
    for (const item of menu.querySelectorAll<HTMLElement>('[data-calc-option]')) {
      const isModeItem =
        item.dataset.calcOption === 'auto' ||
        item.dataset.calcOption === 'manual' ||
        item.dataset.calcOption === 'auto-no-table';
      if (!isModeItem) continue;
      const selected = item.dataset.calcOption === active;
      item.setAttribute('aria-checked', selected ? 'true' : 'false');
      item.classList.toggle('app__menu-item--checked', selected);
    }
  };

  const calcOptionsBtn = document.querySelector<HTMLButtonElement>(
    'button[data-ribbon-command="calcOptions"]',
  );
  const calcOptionsMenu = document.getElementById('menu-calc-options') as HTMLDivElement | null;
  const closeCalcOptionsMenu = (restoreFocus = false): void => {
    if (!calcOptionsMenu) return;
    calcOptionsMenu.hidden = true;
    calcOptionsBtn?.setAttribute('aria-expanded', 'false');
    if (restoreFocus) calcOptionsBtn?.focus();
  };
  const openCalcOptionsMenu = (): void => {
    if (!calcOptionsMenu || !calcOptionsBtn) return;
    closeBorderMenu();
    closeFreezeMenu();
    closeConditionalMenu();
    closeFillMenu();
    closeClearMenu();
    closeCellsMenus();
    closeTextOrientationMenu();
    closeSortFilterHomeMenu();
    closeFindSelectMenu();
    closePasteMenu();
    calcOptionsMenu.hidden = false;
    calcOptionsBtn.setAttribute('aria-haspopup', 'menu');
    calcOptionsBtn.setAttribute('aria-expanded', 'true');
    focusMenuItem(calcOptionsMenu, 'first');
  };
  calcOptionsBtn?.addEventListener('click', (event) => {
    event.preventDefault();
    event.stopPropagation();
    if (!calcOptionsMenu) return;
    if (calcOptionsMenu.hidden) openCalcOptionsMenu();
    else closeCalcOptionsMenu(true);
  });
  calcOptionsMenu?.addEventListener('click', (event) => {
    const item = (event.target as Element | null)?.closest<HTMLButtonElement>('[data-calc-option]');
    const action = item?.dataset.calcOption;
    if (!action) return;
    event.preventDefault();
    event.stopPropagation();
    closeCalcOptionsMenu();
    applyCalcOptionAction(action);
  });
  calcOptionsMenu?.addEventListener('keydown', (event) => {
    handleMenuKeydown(event, calcOptionsMenu, {
      close: closeCalcOptionsMenu,
      restoreFocusTo: calcOptionsBtn,
    });
  });
  document.addEventListener('mousedown', (event) => {
    const target = event.target as Node | null;
    if (
      calcOptionsMenu?.hidden === false &&
      target &&
      !calcOptionsMenu.contains(target) &&
      !calcOptionsBtn?.contains(target)
    ) {
      closeCalcOptionsMenu();
    }
  });
  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && calcOptionsMenu?.hidden === false) closeCalcOptionsMenu(true);
  });

  // ── Chart insert ────────────────────────────────────────────────────────
  const chartInsertBtn = document.querySelector<HTMLButtonElement>(
    'button[data-ribbon-command="chartInsert"]',
  );
  const chartInsertMenu = document.getElementById('menu-chart-insert') as HTMLDivElement | null;
  const closeChartInsertMenu = (restoreFocus = false): void => {
    if (!chartInsertMenu) return;
    chartInsertMenu.hidden = true;
    chartInsertBtn?.setAttribute('aria-expanded', 'false');
    if (restoreFocus) chartInsertBtn?.focus();
  };
  const openChartInsertMenu = (): void => {
    if (!chartInsertMenu || !chartInsertBtn) return;
    closeBorderMenu();
    closeFreezeMenu();
    closeConditionalMenu();
    closeFillMenu();
    closeClearMenu();
    closeCellsMenus();
    closeTextOrientationMenu();
    closeSortFilterHomeMenu();
    closeFindSelectMenu();
    closePasteMenu();
    closeCalcOptionsMenu();
    chartInsertMenu.hidden = false;
    chartInsertBtn.setAttribute('aria-haspopup', 'menu');
    chartInsertBtn.setAttribute('aria-expanded', 'true');
    focusMenuItem(chartInsertMenu, 'first');
  };
  chartInsertBtn?.addEventListener('click', (event) => {
    event.preventDefault();
    event.stopPropagation();
    if (!chartInsertMenu) return;
    if (chartInsertMenu.hidden) openChartInsertMenu();
    else closeChartInsertMenu(true);
  });
  chartInsertMenu?.addEventListener('click', (event) => {
    const item = (event.target as Element | null)?.closest<HTMLButtonElement>(
      '[data-chart-insert]',
    );
    const action = item?.dataset.chartInsert;
    if (!action) return;
    event.preventDefault();
    event.stopPropagation();
    closeChartInsertMenu();
    if (action === 'recommended') void createRecommendedChartFromSelection();
    else createChartFromSelection(chartKindFromAction(action));
  });
  chartInsertMenu?.addEventListener('keydown', (event) => {
    handleMenuKeydown(event, chartInsertMenu, {
      close: closeChartInsertMenu,
      restoreFocusTo: chartInsertBtn,
    });
  });
  document.addEventListener('mousedown', (event) => {
    const target = event.target as Node | null;
    if (
      chartInsertMenu?.hidden === false &&
      target &&
      !chartInsertMenu.contains(target) &&
      !chartInsertBtn?.contains(target)
    ) {
      closeChartInsertMenu();
    }
  });
  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && chartInsertMenu?.hidden === false) closeChartInsertMenu(true);
  });

  // ── Freeze panes ────────────────────────────────────────────────────────
  const freezeBtn = document.getElementById('btn-freeze');
  const freezeMenu = document.getElementById('menu-freeze');
  const getFreezeBtn = (): HTMLButtonElement | null =>
    document.getElementById('btn-freeze') as HTMLButtonElement | null;
  const getFreezeMenu = (): HTMLDivElement | null =>
    document.getElementById('menu-freeze') as HTMLDivElement | null;

  function closeFreezeMenu(restoreFocus = false): void {
    const menu = getFreezeMenu();
    const btn = getFreezeBtn();
    if (!menu) return;
    menu.hidden = true;
    btn?.setAttribute('aria-expanded', 'false');
    if (restoreFocus) btn?.focus();
  }
  const openFreezeMenu = (): void => {
    const menu = getFreezeMenu();
    const btn = getFreezeBtn();
    if (!menu) return;
    menu.hidden = false;
    btn?.setAttribute('aria-expanded', 'true');
    focusMenuItem(menu);
  };

  freezeBtn?.addEventListener('click', (e) => {
    e.stopPropagation();
    if (!freezeMenu) return;
    if (freezeMenu.hidden) openFreezeMenu();
    else closeFreezeMenu();
  });

  document.addEventListener('mousedown', (e) => {
    const menu = getFreezeMenu();
    const btn = getFreezeBtn();
    if (!menu || menu.hidden) return;
    if (menu.contains(e.target as Node)) return;
    if (btn?.contains(e.target as Node)) return;
    closeFreezeMenu();
  });

  document.addEventListener('keydown', (e) => {
    const menu = getFreezeMenu();
    if (e.key === 'Escape' && !menu?.hidden) closeFreezeMenu(true);
  });

  freezeMenu?.addEventListener('keydown', (e) => {
    handleMenuKeydown(e, freezeMenu, { close: closeFreezeMenu, restoreFocusTo: freezeBtn });
  });

  document.addEventListener('keydown', (event) => {
    const target = event.target as Element | null;
    const menu = target?.closest<HTMLDivElement>('#menu-freeze');
    if (!menu || menu === freezeMenu) return;
    handleMenuKeydown(event, menu, { close: closeFreezeMenu, restoreFocusTo: getFreezeBtn() });
  });

  const applyFreezeMenuAction = (action: string | undefined): void => {
    const i = getInst();
    if (!i) return;
    const s = i.store.getState();

    let rows = s.layout.freezeRows;
    let cols = s.layout.freezeCols;
    if (action === 'row') {
      rows = 1;
      cols = 0;
    } else if (action === 'col') {
      rows = 0;
      cols = 1;
    } else if (action === 'selection') {
      // Freeze rows above and columns left of the active cell.
      rows = s.selection.active.row;
      cols = s.selection.active.col;
    } else if (action === 'off') {
      rows = 0;
      cols = 0;
    } else {
      return;
    }

    setFreezePanes(i.store, i.history, rows, cols, i.workbook);
    closeFreezeMenu();
    getSheetEl().focus();
  };

  freezeMenu?.querySelectorAll<HTMLButtonElement>('[data-freeze]').forEach((btn) => {
    btn.addEventListener('click', () => {
      applyFreezeMenuAction(btn.dataset.freeze);
    });
  });

  document.addEventListener('click', (event) => {
    const target = event.target as Element | null;
    const menu = target?.closest<HTMLElement>('#menu-freeze');
    if (!menu || menu === freezeMenu) return;
    const item = target?.closest<HTMLButtonElement>('[data-freeze]');
    if (!item) return;
    event.preventDefault();
    applyFreezeMenuAction(item.dataset.freeze);
  });

  return {
    applySortMenuAction,
    applyCalcOptionAction,
    applyFreezeMenuAction,
    updateCalcOptionsMenu,
    openSortFilterHomeMenu,
    closeSortFilterHomeMenu,
    openDataSortMenu,
    closeDataSortMenu,
    openFindSelectMenu,
    closeFindSelectMenu,
    openCalcOptionsMenu,
    closeCalcOptionsMenu,
    openChartInsertMenu,
    closeChartInsertMenu,
    openFreezeMenu,
    closeFreezeMenu,
    getFreezeBtn,
    getFreezeMenu,
  };
};
