import type { Addr } from '../engine/types.js';
import { formatCell } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { type Strings, defaultStrings } from '../i18n/strings.js';
import { type SpreadsheetStore, mutators } from '../store/store.js';

export interface WatchPanelDeps {
  /** Element the panel docks under. The panel appends itself as a child of
   *  `host` so it inherits theme + tab focus scope. */
  host: HTMLElement;
  store: SpreadsheetStore;
  /** Lazy workbook getter — callers should pass the same closure mount.ts
   *  uses elsewhere so a `setWorkbook` swap is picked up automatically. */
  getWb: () => WorkbookHandle;
  strings?: Strings;
  /** Optional hook for after-write style follow-ups. The watch panel itself
   *  performs no writes, so this is reserved for symmetry with the rest of
   *  the interact modules. */
  onAfterCommit?: () => void;
}

export interface WatchPanelHandle {
  open(): void;
  close(): void;
  toggle(): void;
  /** Re-read every watched cell from the workbook and repaint the rows.
   *  Cheap — runs on every recalc batch from the engine. */
  refresh(): void;
  detach(): void;
}

/** Excel column-letter conversion (0-indexed). */
const colLetter = (col: number): string => {
  let n = col;
  let out = '';
  do {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
};

const a1 = (addr: Addr): string => `${colLetter(addr.col)}${addr.row + 1}`;

/**
 * Excel-style Watch Window. Lists pinned cells with live values; updates on
 * every store change (covers both direct cell mutation and the recalc-fed
 * value refresh path in mount.ts). Click a row to jump the active selection
 * to that cell, switching sheet if needed.
 */
export function attachWatchPanel(deps: WatchPanelDeps): WatchPanelHandle {
  const { host, store, getWb } = deps;
  let strings = deps.strings ?? defaultStrings;

  const root = document.createElement('div');
  root.className = 'fc-watch';
  root.dataset.fcWatch = '1';
  root.setAttribute('role', 'region');
  root.setAttribute('aria-label', strings.watchPanel.title);
  root.hidden = !store.getState().ui.watchPanelOpen;

  const header = document.createElement('div');
  header.className = 'fc-watch__header';
  const title = document.createElement('span');
  title.className = 'fc-watch__title';
  title.textContent = strings.watchPanel.title;
  const actions = document.createElement('span');
  actions.className = 'fc-watch__actions';
  const addBtn = document.createElement('button');
  addBtn.type = 'button';
  addBtn.className = 'fc-watch__btn';
  addBtn.textContent = strings.watchPanel.addWatch;
  const clearBtn = document.createElement('button');
  clearBtn.type = 'button';
  clearBtn.className = 'fc-watch__btn';
  clearBtn.textContent = strings.watchPanel.clearAll;
  const closeBtn = document.createElement('button');
  closeBtn.type = 'button';
  closeBtn.className = 'fc-watch__btn fc-watch__close';
  closeBtn.setAttribute('aria-label', strings.watchPanel.close);
  closeBtn.textContent = '×';
  actions.append(addBtn, clearBtn, closeBtn);
  header.append(title, actions);

  const body = document.createElement('div');
  body.className = 'fc-watch__body';

  const table = document.createElement('table');
  table.className = 'fc-watch__table';
  const thead = document.createElement('thead');
  const headRow = document.createElement('tr');
  const headers = [
    'sheetHeader',
    'cellHeader',
    'nameHeader',
    'valueHeader',
    'formulaHeader',
  ] as const;
  for (const key of headers) {
    const th = document.createElement('th');
    th.scope = 'col';
    th.dataset.fcCol = key;
    th.textContent = strings.watchPanel[key];
    headRow.appendChild(th);
  }
  // Trailing column for the per-row remove (×) button. No header label.
  const thRemove = document.createElement('th');
  thRemove.scope = 'col';
  thRemove.dataset.fcCol = 'remove';
  thRemove.setAttribute('aria-hidden', 'true');
  headRow.appendChild(thRemove);
  thead.appendChild(headRow);
  const tbody = document.createElement('tbody');
  table.append(thead, tbody);
  body.appendChild(table);

  const empty = document.createElement('div');
  empty.className = 'fc-watch__empty';
  empty.textContent = strings.watchPanel.empty;

  root.append(header, body, empty);
  host.appendChild(root);

  /** Find a defined name (workbook scope) whose ref points exactly at `addr`.
   *  Returns null when there's no exact match — which keeps the column quiet
   *  for cells that aren't named. */
  const nameOf = (addr: Addr): string => {
    try {
      const wb = getWb();
      const target = a1(addr).toUpperCase();
      for (const dn of wb.definedNames()) {
        const eq = dn.formula.replace(/^=/, '').replace(/\$/g, '').toUpperCase();
        const bang = eq.lastIndexOf('!');
        const tail = bang >= 0 ? eq.slice(bang + 1) : eq;
        if (tail === target) return dn.name;
      }
    } catch {
      // Engine read failed (e.g. mid-dispose); treat as no name available.
    }
    return '';
  };

  const sheetLabel = (sheet: number): string => {
    try {
      return getWb().sheetName(sheet);
    } catch {
      return `Sheet${sheet + 1}`;
    }
  };

  const readValue = (addr: Addr): string => {
    try {
      return formatCell(getWb().getValue(addr));
    } catch {
      return '';
    }
  };

  const formulaOf = (addr: Addr): string => {
    try {
      return getWb().cellFormula(addr) ?? '';
    } catch {
      return '';
    }
  };

  const renderRows = (): void => {
    const watches = store.getState().watch.watches;
    tbody.replaceChildren();
    if (watches.length === 0) {
      empty.hidden = false;
      table.hidden = true;
      return;
    }
    empty.hidden = true;
    table.hidden = false;
    for (const addr of watches) {
      const tr = document.createElement('tr');
      tr.className = 'fc-watch__row';
      tr.dataset.fcSheet = String(addr.sheet);
      tr.dataset.fcRow = String(addr.row);
      tr.dataset.fcCol = String(addr.col);
      tr.tabIndex = 0;

      const tdSheet = document.createElement('td');
      tdSheet.textContent = sheetLabel(addr.sheet);
      const tdCell = document.createElement('td');
      tdCell.textContent = a1(addr);
      const tdName = document.createElement('td');
      tdName.textContent = nameOf(addr);
      const tdValue = document.createElement('td');
      tdValue.className = 'fc-watch__value';
      tdValue.textContent = readValue(addr);
      const tdFormula = document.createElement('td');
      tdFormula.className = 'fc-watch__formula';
      tdFormula.textContent = formulaOf(addr);

      const tdRemove = document.createElement('td');
      tdRemove.className = 'fc-watch__remove-cell';
      const removeBtn = document.createElement('button');
      removeBtn.type = 'button';
      removeBtn.className = 'fc-watch__remove';
      removeBtn.setAttribute('aria-label', strings.watchPanel.removeWatch);
      removeBtn.textContent = '×';
      removeBtn.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        mutators.removeWatch(store, addr);
      });
      tdRemove.appendChild(removeBtn);

      tr.append(tdSheet, tdCell, tdName, tdValue, tdFormula, tdRemove);
      tr.addEventListener('click', () => jumpTo(addr));
      tr.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault();
          jumpTo(addr);
        }
      });
      tbody.appendChild(tr);
    }
  };

  /** Jump the active selection to `addr`. Switches sheet (and re-hydrates
   *  the cell map) when the watched cell lives elsewhere. */
  const jumpTo = (addr: Addr): void => {
    const s = store.getState();
    if (s.data.sheetIndex !== addr.sheet) {
      mutators.setSheetIndex(store, addr.sheet);
      try {
        const wb = getWb();
        mutators.replaceCells(store, wb.cells(addr.sheet));
      } catch {
        // If the sheet swap can't be hydrated (engine torn down), the
        // setActive below still moves the selection so the user sees
        // the consistent selection state.
      }
    }
    mutators.setActive(store, addr);
  };

  const refresh = (): void => {
    renderRows();
  };

  const open = (): void => {
    mutators.setWatchPanelOpen(store, true);
  };
  const close = (): void => {
    mutators.setWatchPanelOpen(store, false);
  };
  const toggle = (): void => {
    mutators.setWatchPanelOpen(store, !store.getState().ui.watchPanelOpen);
  };

  const onAdd = (): void => {
    mutators.addWatch(store, store.getState().selection.active);
  };
  const onClear = (): void => {
    mutators.clearWatches(store);
  };
  const onClose = (): void => close();
  addBtn.addEventListener('click', onAdd);
  clearBtn.addEventListener('click', onClear);
  closeBtn.addEventListener('click', onClose);

  // Re-render on any store change. Cheap — the table is at most a few rows.
  // Tracks watches list, panel visibility, and live cell-value updates that
  // mount.ts pipes into `data.cells` after recalc.
  let lastWatches = store.getState().watch.watches;
  let lastVisible = store.getState().ui.watchPanelOpen;
  let lastCells = store.getState().data.cells;
  let lastSheetIdx = store.getState().data.sheetIndex;
  const unsub = store.subscribe(() => {
    const s = store.getState();
    const visible = s.ui.watchPanelOpen;
    if (visible !== lastVisible) {
      lastVisible = visible;
      root.hidden = !visible;
      if (visible) refresh();
    }
    const watchesChanged = s.watch.watches !== lastWatches;
    const cellsChanged = s.data.cells !== lastCells;
    const sheetChanged = s.data.sheetIndex !== lastSheetIdx;
    if (watchesChanged) lastWatches = s.watch.watches;
    if (cellsChanged) lastCells = s.data.cells;
    if (sheetChanged) lastSheetIdx = s.data.sheetIndex;
    if (visible && (watchesChanged || cellsChanged || sheetChanged)) refresh();
  });

  if (!root.hidden) refresh();

  return {
    open,
    close,
    toggle,
    refresh,
    detach(): void {
      addBtn.removeEventListener('click', onAdd);
      clearBtn.removeEventListener('click', onClear);
      closeBtn.removeEventListener('click', onClose);
      unsub();
      root.remove();
    },
    setStrings(next: Strings): void {
      strings = next;
      title.textContent = strings.watchPanel.title;
      addBtn.textContent = strings.watchPanel.addWatch;
      clearBtn.textContent = strings.watchPanel.clearAll;
      closeBtn.setAttribute('aria-label', strings.watchPanel.close);
      empty.textContent = strings.watchPanel.empty;
      const ths = headRow.querySelectorAll('th');
      ths.forEach((th, i) => {
        const key = headers[i];
        if (key) th.textContent = strings.watchPanel[key];
      });
      root.setAttribute('aria-label', strings.watchPanel.title);
      refresh();
    },
  } as WatchPanelHandle & { setStrings(next: Strings): void };
}
