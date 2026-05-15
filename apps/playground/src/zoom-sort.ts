import {
  clearFilter,
  mutators,
  removeDuplicates,
  type SpreadsheetInstance,
  setSheetZoom,
  sortRange,
} from '@libraz/formulon-cell';
import { focusMenuItem, handleMenuKeydown, prepareMenu } from './menu-a11y.js';

export function setupZoomControls(getInst: () => SpreadsheetInstance | null): {
  refreshZoom(): void;
} {
  const zoomDisplay = document.getElementById('zoom-display');
  const zoomRailFill = document.getElementById('zoom-rail-fill');
  const zoomRailThumb = document.getElementById('zoom-rail-thumb');
  const zMin = 0.25;
  const zMax = 3.0;
  const refreshZoom = (): void => {
    const inst = getInst();
    if (!inst) return;
    const z = inst.store.getState().viewport.zoom;
    if (zoomDisplay) zoomDisplay.textContent = `${Math.round(z * 100)}%`;
    const pct = Math.max(0, Math.min(1, (z - zMin) / (zMax - zMin))) * 100;
    if (zoomRailFill) zoomRailFill.style.width = `${pct}%`;
    if (zoomRailThumb) zoomRailThumb.style.left = `${pct}%`;
  };
  const stepZoom = (delta: number): void => {
    const inst = getInst();
    if (!inst) return;
    const z = inst.store.getState().viewport.zoom;
    const next = Math.max(zMin, Math.min(zMax, Math.round((z + delta) * 100) / 100));
    if (next === z) return;
    setSheetZoom(inst.store, next, inst.workbook);
    refreshZoom();
  };
  zoomDisplay?.addEventListener('click', () => {
    const inst = getInst();
    if (!inst) return;
    const z = inst.store.getState().viewport.zoom;
    const next = z >= 1.5 ? 0.75 : Math.round((z + 0.25) * 100) / 100;
    setSheetZoom(inst.store, next, inst.workbook);
    refreshZoom();
  });
  document.getElementById('btn-zoom-out')?.addEventListener('click', () => stepZoom(-0.1));
  document.getElementById('btn-zoom-in')?.addEventListener('click', () => stepZoom(0.1));
  return { refreshZoom };
}

export function setupSortMenu(input: {
  getFilterDropdown: () => {
    open(range: unknown, col: number, anchor: { x: number; y: number; h: number }): void;
  } | null;
  getInst: () => SpreadsheetInstance | null;
  sheetEl: HTMLElement;
  statusMetric: HTMLElement | null;
}): void {
  const sortBtn = document.getElementById('btn-sort');
  const sortMenu = document.getElementById('menu-sort');
  if (sortMenu) prepareMenu(sortMenu, 'Sort and filter');

  const closeSortMenu = (restoreFocus = false): void => {
    if (!sortMenu) return;
    sortMenu.hidden = true;
    sortBtn?.setAttribute('aria-expanded', 'false');
    if (restoreFocus) sortBtn?.focus();
  };
  const openSortMenu = (): void => {
    if (!sortMenu) return;
    sortMenu.hidden = false;
    sortBtn?.setAttribute('aria-expanded', 'true');
    focusMenuItem(sortMenu);
  };
  sortBtn?.addEventListener('click', (e) => {
    e.stopPropagation();
    if (!sortMenu) return;
    if (sortMenu.hidden) openSortMenu();
    else closeSortMenu();
  });
  document.addEventListener('mousedown', (e) => {
    if (!sortMenu || sortMenu.hidden) return;
    if (sortMenu.contains(e.target as Node)) return;
    if (sortBtn?.contains(e.target as Node)) return;
    closeSortMenu();
  });
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && !sortMenu?.hidden) closeSortMenu(true);
  });
  sortMenu?.addEventListener('keydown', (e) => {
    handleMenuKeydown(e, sortMenu, { close: closeSortMenu, restoreFocusTo: sortBtn });
  });

  sortMenu?.querySelectorAll<HTMLButtonElement>('[data-sort]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const inst = input.getInst();
      if (!inst) return;
      const action = btn.dataset.sort;
      closeSortMenu();
      const state = inst.store.getState();
      const r = state.selection.range;
      if (r.r0 === r.r1 && r.c0 === r.c1) return;
      if (action === 'asc' || action === 'desc') {
        sortRange(state, inst.store, inst.workbook, r, { byCol: r.c0, direction: action });
        mutators.replaceCells(inst.store, inst.workbook.cells(state.data.sheetIndex));
      } else if (action === 'dedupe') {
        const removed = removeDuplicates(state, inst.store, inst.workbook, r);
        mutators.replaceCells(inst.store, inst.workbook.cells(state.data.sheetIndex));
        if (input.statusMetric)
          input.statusMetric.textContent = `Removed ${removed} duplicate row${removed === 1 ? '' : 's'}`;
      } else if (action === 'filter') {
        mutators.setFilterRange(inst.store, r);
        const sheetRect = input.sheetEl.getBoundingClientRect();
        input
          .getFilterDropdown()
          ?.open(r, r.c0, { x: sheetRect.left + 80, y: sheetRect.top, h: 32 });
      } else if (action === 'filter-clear') {
        clearFilter(state, inst.store, r);
      } else if (action === 'conditional') {
        inst.openConditionalDialog();
      } else if (action === 'named') {
        inst.openNamedRangeDialog();
      }
      input.sheetEl.focus();
    });
  });
}
