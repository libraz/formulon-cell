import { totalsCellAt } from '../commands/format-as-table.js';
import type { History } from '../commands/history.js';
import { setTableColumnFunction, TABLE_TOTALS_FUNCTIONS } from '../commands/table-design.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import { getTotalsChevron } from '../render/grid.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { overlayPortalFor } from './overlay-portal.js';

export interface TableTotalsDeps {
  /** The grid surface that paints the chevron and receives clicks. */
  grid: HTMLElement;
  store: SpreadsheetStore;
  history: History | null;
  wb: WorkbookHandle;
  strings: Strings;
  /** Refresh cached cells after the engine rewrote a totals cell. */
  onAfterCommit: () => void;
  /** Leave the totals cell open for typing a formula ("More Functions"). */
  onMoreFunctions: () => void;
}

export interface TableTotalsHandle {
  /** Open the picker for the active cell; false when it is not a totals cell. */
  openForActiveCell(): boolean;
  detach(): void;
}

const functionLabel = (strings: Strings, fn: (typeof TABLE_TOTALS_FUNCTIONS)[number]): string => {
  const pivot = strings.pivotTableDialog;
  switch (fn) {
    case 'none':
      return strings.ribbon.tableTotalsNone;
    case 'average':
      return pivot.average;
    case 'count':
      return pivot.count;
    case 'countNums':
      return pivot.countNumbers;
    case 'max':
      return pivot.max;
    case 'min':
      return pivot.min;
    case 'sum':
      return pivot.sum;
    case 'stdDev':
      return pivot.stdDev;
    case 'var':
      return pivot.variance;
    default:
      return strings.ribbonMenu.autosumMoreFunctions;
  }
};

/**
 * Owns the totals-row function picker: a chevron on the active cell of a
 * table's total row opens a list of functions; picking one sets the column's
 * totals function, which the store mirror writes to the engine.
 */
export function attachTableTotals(deps: TableTotalsDeps): TableTotalsHandle {
  const { grid, store, strings } = deps;
  let popover: HTMLDivElement | null = null;
  let restoreFocus: HTMLElement | null = null;

  const close = (shouldRestoreFocus = false): void => {
    if (!popover) return;
    popover.remove();
    popover = null;
    document.removeEventListener('mousedown', onDocMouseDown, true);
    document.removeEventListener('keydown', onDocKey, true);
    const target = restoreFocus;
    restoreFocus = null;
    if (shouldRestoreFocus) target?.focus({ preventScroll: true });
  };
  const onDocMouseDown = (e: MouseEvent): void => {
    if (popover && !popover.contains(e.target as Node)) close();
  };
  const onDocKey = (e: KeyboardEvent): void => {
    if (e.key !== 'Escape') return;
    e.preventDefault();
    close(true);
  };

  const items = (): HTMLElement[] =>
    popover ? Array.from(popover.querySelectorAll<HTMLElement>('.fc-validation-list__item')) : [];
  const focusItem = (idx: number): void => {
    const options = items();
    if (options.length === 0) return;
    const next = (idx + options.length) % options.length;
    for (const [i, option] of options.entries()) {
      option.tabIndex = i === next ? 0 : -1;
      option.setAttribute('aria-selected', i === next ? 'true' : 'false');
    }
    options[next]?.focus({ preventScroll: true });
  };

  const pick = (tableId: string, column: number, fn: (typeof TABLE_TOTALS_FUNCTIONS)[number]) => {
    close();
    if (fn === 'custom') {
      deps.onMoreFunctions();
      return;
    }
    if (setTableColumnFunction(store, deps.history, deps.wb, tableId, column, fn)) {
      deps.onAfterCommit();
    }
    grid.focus({ preventScroll: true });
  };

  const open = (): boolean => {
    close();
    const chevron = getTotalsChevron();
    if (!chevron) return false;
    const s = store.getState();
    const totals = totalsCellAt(s, s.data.sheetIndex, chevron.row, chevron.col);
    if (!totals) return false;
    const current = totals.table.totals?.[totals.column]?.function ?? 'none';
    const rect = grid.getBoundingClientRect();
    const div = document.createElement('div');
    div.className = 'fc-validation-list fc-table-totals';
    div.style.position = 'fixed';
    div.style.left = `${rect.left + chevron.rect.x}px`;
    div.style.top = `${rect.top + chevron.rect.y + chevron.rect.h}px`;
    div.setAttribute('role', 'listbox');
    div.setAttribute('aria-label', strings.ribbon.tableTotalsFunction);
    div.tabIndex = -1;
    restoreFocus =
      document.activeElement instanceof HTMLElement && document.activeElement !== document.body
        ? document.activeElement
        : grid;
    const selected = Math.max(0, TABLE_TOTALS_FUNCTIONS.indexOf(current));
    for (const [idx, fn] of TABLE_TOTALS_FUNCTIONS.entries()) {
      const item = document.createElement('div');
      item.className = 'fc-validation-list__item';
      item.dataset.totalsFunction = fn;
      item.setAttribute('role', 'option');
      item.setAttribute('aria-selected', idx === selected ? 'true' : 'false');
      item.tabIndex = idx === selected ? 0 : -1;
      item.textContent = functionLabel(strings, fn);
      const choose = (e: Event): void => {
        e.preventDefault();
        pick(totals.table.id, totals.column, fn);
      };
      item.addEventListener('mousedown', choose);
      item.addEventListener('click', choose);
      div.appendChild(item);
    }
    div.addEventListener('keydown', (e) => {
      const options = items();
      const active = document.activeElement instanceof HTMLElement ? document.activeElement : null;
      const idx = active ? options.indexOf(active) : -1;
      if (e.key === 'ArrowDown') focusItem(idx + 1);
      else if (e.key === 'ArrowUp') focusItem(idx - 1);
      else if (e.key === 'Home') focusItem(0);
      else if (e.key === 'End') focusItem(options.length - 1);
      else if (e.key === 'Enter' || e.key === ' ') {
        const fn = TABLE_TOTALS_FUNCTIONS[idx];
        if (!fn) return;
        pick(totals.table.id, totals.column, fn);
      } else if (e.key === 'Escape') close(true);
      else return;
      e.preventDefault();
    });
    overlayPortalFor(grid).appendChild(div);
    popover = div;
    focusItem(selected);
    document.addEventListener('mousedown', onDocMouseDown, true);
    document.addEventListener('keydown', onDocKey, true);
    return true;
  };

  const onDown = (e: PointerEvent): void => {
    if (e.button !== 0) return;
    const chevron = getTotalsChevron();
    if (!chevron) return;
    const rect = grid.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;
    if (
      x < chevron.rect.x ||
      x > chevron.rect.x + chevron.rect.w ||
      y < chevron.rect.y ||
      y > chevron.rect.y + chevron.rect.h
    ) {
      return;
    }
    e.preventDefault();
    e.stopPropagation();
    const sheet = store.getState().data.sheetIndex;
    mutators.setActive(store, { sheet, row: chevron.row, col: chevron.col });
    open();
  };
  const onGridKeyDown = (e: KeyboardEvent): void => {
    if (!e.altKey || e.key !== 'ArrowDown') return;
    if (!open()) return;
    e.preventDefault();
    e.stopPropagation();
  };

  // Capture phase so we beat the regular pointer.ts handler.
  grid.addEventListener('pointerdown', onDown, true);
  grid.addEventListener('keydown', onGridKeyDown, true);

  return {
    openForActiveCell: open,
    detach() {
      close();
      grid.removeEventListener('pointerdown', onDown, true);
      grid.removeEventListener('keydown', onGridKeyDown, true);
    },
  };
}
