import { inputLocaleForWorkbook } from '../commands/coerce-input.js';
import type { History } from '../commands/history.js';
import { recordWatchesChange } from '../commands/watch.js';
import { addrKey, formatA1Cell } from '../engine/address.js';
import type { Addr } from '../engine/types.js';
import { formatCell } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import { selectionRangeList } from '../store/selection-geometry.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import { sheetFocusTarget } from './dialog-focus.js';
import {
  appendDialogButton,
  bindPaneEscape,
  deferInitialFocus,
  markSurface,
} from './dialog-shell.js';
import { bindRovingRows, setRovingStop } from './roving-rows.js';
import { syncTaskpaneDock } from './taskpane-dock.js';

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
  /** Optional history stack used to keep panel Add/Delete/Clear aligned with
   *  ribbon Watch Window commands. */
  history?: History | null;
  /** Workbook label shown in the Excel-compatible Book column. */
  workbookName?: string;
  /** Optional live workbook label. Explicit `workbookName` remains highest
   * priority; an absent getter keeps the existing neutral `Book1` fallback. */
  getWorkbookName?: () => string | null | undefined;
  /** Current UI locale used only when the native display renderer is absent. */
  getLocale?: () => string;
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

/**
 * Spreadsheet-style Watch Window. Lists pinned cells with live values; updates on
 * every store change (covers both direct cell mutation and the recalc-fed
 * value refresh path in mount.ts). Click a row to jump the active selection
 * to that cell, switching sheet if needed.
 */
export function attachWatchPanel(deps: WatchPanelDeps): WatchPanelHandle {
  const { host, store, getWb } = deps;
  let strings = deps.strings ?? defaultStrings;
  // The Mac host docks the window as a task pane beside the grid, with a tab,
  // a + / - toolbar acting on the selected watch, and no per-row remove column.
  const fcHost = host.closest<HTMLElement>('.fc-host') ?? host;
  const mac = fcHost.dataset.fcPlatform === 'mac';
  const dock = mac ? fcHost.querySelector<HTMLElement>('.fc-host__taskpane-dock') : null;
  const macLabels = () => strings.mac.watchWindow;
  let selectedKey: string | null = null;

  const root = document.createElement('div');
  root.className = 'fc-watch';
  root.classList.toggle('fc-watch--mac', mac);
  root.dataset.fcWatch = '1';
  root.setAttribute('role', 'region');
  root.setAttribute('aria-label', mac ? macLabels().title : strings.watchPanel.title);
  root.hidden = !store.getState().ui.watchPanelOpen;
  root.tabIndex = -1;

  const header = document.createElement('div');
  header.className = 'fc-watch__header';
  const title = document.createElement('span');
  title.className = 'fc-watch__title';
  title.textContent = mac ? macLabels().title : strings.watchPanel.title;
  const actions = document.createElement('span');
  actions.className = 'fc-watch__actions';
  const toolbar = document.createElement('div');
  toolbar.className = 'fc-watch__toolbar';
  const addBtn = appendDialogButton(mac ? toolbar : actions, {
    label: mac ? '' : strings.watchPanel.addWatch,
    baseClass: mac ? 'fc-watch__tool fc-watch__tool--add' : 'fc-watch__btn',
  });
  const removeTool = mac
    ? appendDialogButton(toolbar, {
        label: '',
        baseClass: 'fc-watch__tool fc-watch__tool--remove',
      })
    : null;
  const clearBtn = mac
    ? null
    : appendDialogButton(actions, {
        label: strings.watchPanel.clearAll,
        baseClass: 'fc-watch__btn',
      });
  const closeBtn = appendDialogButton(actions, { label: '', baseClass: 'fc-watch__btn' });
  closeBtn.classList.add('fc-watch__close');
  const labelMacTools = (): void => {
    if (!mac) return;
    addBtn.setAttribute('aria-label', macLabels().add);
    addBtn.title = macLabels().add;
    removeTool?.setAttribute('aria-label', macLabels().remove);
    if (removeTool) removeTool.title = macLabels().remove;
  };
  closeBtn.setAttribute('aria-label', mac ? macLabels().close : strings.watchPanel.close);
  labelMacTools();
  header.append(title, actions);

  const body = document.createElement('div');
  body.className = 'fc-watch__body';

  const table = document.createElement('table');
  table.className = 'fc-watch__table';
  const thead = document.createElement('thead');
  const headRow = document.createElement('tr');
  const headers = [
    'bookHeader',
    'sheetHeader',
    'nameHeader',
    'cellHeader',
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
  // Trailing column for the per-row remove button. No header label.
  if (!mac) {
    const thRemove = document.createElement('th');
    thRemove.scope = 'col';
    thRemove.dataset.fcCol = 'remove';
    thRemove.setAttribute('aria-hidden', 'true');
    headRow.appendChild(thRemove);
  }
  thead.appendChild(headRow);
  const tbody = document.createElement('tbody');
  table.append(thead, tbody);
  body.appendChild(table);

  const empty = document.createElement('div');
  empty.className = 'fc-watch__empty';
  empty.textContent = strings.watchPanel.empty;

  if (mac) root.append(header, toolbar, body, empty);
  else root.append(header, body, empty);
  (dock ?? host).appendChild(root);
  markSurface(root, 'pane');
  const syncDock = (): void => {
    if (dock) syncTaskpaneDock(dock);
  };
  syncDock();

  let restoreFocusEl: HTMLElement | null = null;

  const focusInitial = (): void => {
    if (root.contains(document.activeElement)) return;
    deferInitialFocus(root, () => addBtn.focus({ preventScroll: true }));
  };

  const restoreFocus = (): void => {
    const target = restoreFocusEl;
    restoreFocusEl = null;
    if (!target) return;
    if (root.contains(document.activeElement) || document.activeElement === document.body) {
      target.focus({ preventScroll: true });
    }
  };

  /** Map each defined name's exact single-cell target (A1, sheet prefix and
   *  `$` stripped) to its name; the first name wins. Empty when the engine
   *  read fails (e.g. mid-dispose). */
  const definedNameLookup = (): Map<string, string> => {
    const byTarget = new Map<string, string>();
    try {
      for (const dn of getWb().definedNames()) {
        const eq = dn.formula.replace(/^=/, '').replace(/\$/g, '').toUpperCase();
        const bang = eq.lastIndexOf('!');
        const tail = bang >= 0 ? eq.slice(bang + 1) : eq;
        if (!byTarget.has(tail)) byTarget.set(tail, dn.name);
      }
    } catch {
      // Engine read failed (e.g. mid-dispose); treat as no names available.
    }
    return byTarget;
  };

  const sheetLabel = (sheet: number): string => {
    try {
      return getWb().sheetName(sheet);
    } catch {
      return `Sheet${sheet + 1}`;
    }
  };

  const bookLabel = (): string => {
    if (deps.workbookName !== undefined) return deps.workbookName;
    try {
      const live = deps.getWorkbookName?.();
      if (live) return live;
    } catch {
      // A host-provided label is presentation-only; keep the stable fallback
      // while a workbook is being swapped or disposed.
    }
    return 'Book1';
  };

  const workbookDisplayLocale = (wb: WorkbookHandle): string => {
    // en-US defers to the host locale like an unmapped profile.
    const profileLocale = inputLocaleForWorkbook(wb);
    if (profileLocale && profileLocale !== 'en-US') return profileLocale;
    return deps.getLocale?.() ?? 'en-US';
  };

  const readValue = (addr: Addr): string => {
    try {
      const wb = getWb();
      // getDisplayText is workbook-aware: it applies the active profile,
      // number format, date system, localized booleans, and error spelling.
      // An empty string is a valid display for a blank cell, so test null.
      const native =
        typeof wb.getDisplayText === 'function'
          ? wb.getDisplayText(addr.sheet, addr.row, addr.col)
          : null;
      if (native !== null) return native;
      return formatCell(wb.getValue(addr), workbookDisplayLocale(wb));
    } catch {
      return '';
    }
  };

  const formulaOf = (addr: Addr): string => {
    try {
      const wb = getWb();
      const formula = wb.cellFormula(addr) ?? '';
      if (!formula) return '';
      // Stored formulas stay canonical; the Watch Window is a display surface
      // and follows the active profile's aliases/separators when available.
      return typeof wb.getFormulaForDisplay === 'function'
        ? wb.getFormulaForDisplay(formula)
        : formula;
    } catch {
      return '';
    }
  };

  const syncRemoveTool = (): void => {
    if (!removeTool) return;
    const watches = store.getState().watch.watches;
    if (selectedKey !== null && !watches.some((addr) => addrKey(addr) === selectedKey)) {
      selectedKey = null;
    }
    projectDisabledState(removeTool, selectedKey === null, null);
  };

  const watchRows = (): HTMLElement[] =>
    Array.from(tbody.querySelectorAll<HTMLElement>(':scope > .fc-watch__row'));
  const detachRovingRows = bindRovingRows(tbody, watchRows);

  // Removing the focused row hands focus to the next row, else the previous, else Add.
  const removeRow = (row: HTMLElement, addr: Addr): void => {
    const hadFocus = row.contains(document.activeElement);
    const at = watchRows().indexOf(row);
    recordWatchesChange(deps.history ?? null, store, () => {
      mutators.removeWatch(store, addr);
    });
    if (!hadFocus) return;
    const rows = watchRows();
    const next = rows[at] ?? rows[at - 1];
    if (!next) {
      addBtn.focus();
      return;
    }
    setRovingStop(rows, next);
    next.focus();
  };

  const renderRows = (): void => {
    const watches = store.getState().watch.watches;
    syncRemoveTool();
    tbody.replaceChildren();
    // The Mac pane keeps its column headers over an empty, striped list.
    if (watches.length === 0) {
      empty.hidden = mac;
      table.hidden = !mac;
      return;
    }
    empty.hidden = true;
    table.hidden = false;
    const names = definedNameLookup();
    const book = bookLabel();
    const active = store.getState().selection.active;
    for (const addr of watches) {
      const selected = mac
        ? addrKey(addr) === selectedKey
        : active.sheet === addr.sheet && active.row === addr.row && active.col === addr.col;
      const tr = document.createElement('tr');
      tr.className = 'fc-watch__row';
      tr.classList.toggle('fc-watch__row--selected', selected);
      tr.dataset.fcSheet = String(addr.sheet);
      tr.dataset.fcRow = String(addr.row);
      tr.dataset.fcCol = String(addr.col);
      tr.setAttribute('aria-selected', String(selected));

      const tdBook = document.createElement('td');
      tdBook.textContent = book;
      const tdSheet = document.createElement('td');
      tdSheet.textContent = sheetLabel(addr.sheet);
      const tdName = document.createElement('td');
      tdName.textContent = names.get(formatA1Cell(addr.row, addr.col).toUpperCase()) ?? '';
      const tdCell = document.createElement('td');
      tdCell.textContent = formatA1Cell(addr.row, addr.col);
      const tdValue = document.createElement('td');
      tdValue.className = 'fc-watch__value';
      tdValue.textContent = readValue(addr);
      const tdFormula = document.createElement('td');
      tdFormula.className = 'fc-watch__formula';
      tdFormula.textContent = formulaOf(addr);

      tr.append(tdBook, tdSheet, tdName, tdCell, tdValue, tdFormula);
      if (!mac) {
        const tdRemove = document.createElement('td');
        tdRemove.className = 'fc-watch__remove-cell';
        const removeBtn = appendDialogButton(tdRemove, {
          label: '',
          baseClass: 'fc-watch__remove',
        });
        removeBtn.setAttribute('aria-label', strings.watchPanel.removeWatch);
        // Reached from its row, which takes the list's one tab stop.
        removeBtn.tabIndex = -1;
        removeBtn.addEventListener('click', (e) => {
          e.preventDefault();
          e.stopPropagation();
          removeRow(tr, addr);
        });
        tr.appendChild(tdRemove);
      }
      const choose = (): void => {
        if (mac) selectedKey = addrKey(addr);
        jumpTo(addr);
        if (mac) refresh();
      };
      tr.addEventListener('click', choose);
      tr.addEventListener('keydown', (e) => {
        if (e.target !== tr) return;
        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault();
          choose();
        } else if (e.key === 'Delete') {
          e.preventDefault();
          removeRow(tr, addr);
        }
      });
      tbody.appendChild(tr);
    }
    // The rows take one tab stop, on the selected watch.
    const rendered = watchRows();
    setRovingStop(
      rendered,
      rendered.find((row) => row.getAttribute('aria-selected') === 'true'),
    );
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
    restoreFocusEl =
      document.activeElement instanceof HTMLElement
        ? document.activeElement
        : sheetFocusTarget(host);
    mutators.setWatchPanelOpen(store, true);
    focusInitial();
  };
  const close = (restore = true): void => {
    mutators.setWatchPanelOpen(store, false);
    if (restore) restoreFocus();
  };
  const toggle = (): void => {
    if (store.getState().ui.watchPanelOpen) {
      close(true);
    } else {
      open();
    }
  };

  const onAdd = (): void => {
    const selection = store.getState().selection;
    recordWatchesChange(deps.history ?? null, store, () => {
      mutators.addWatchRanges(store, selectionRangeList(selection));
    });
  };
  const onRemoveSelected = (): void => {
    const key = selectedKey;
    const addr = store.getState().watch.watches.find((watch) => addrKey(watch) === key);
    if (!addr) return;
    selectedKey = null;
    recordWatchesChange(deps.history ?? null, store, () => {
      mutators.removeWatch(store, addr);
    });
    syncRemoveTool();
  };
  const onClear = (): void => {
    recordWatchesChange(deps.history ?? null, store, () => {
      mutators.clearWatches(store);
    });
  };
  const onClose = (): void => close(true);
  addBtn.addEventListener('click', onAdd);
  removeTool?.addEventListener('click', onRemoveSelected);
  clearBtn?.addEventListener('click', onClear);
  closeBtn.addEventListener('click', onClose);
  const detachPaneEscape = bindPaneEscape(root, fcHost);

  // Re-render on any store change. Cheap — the table is at most a few rows.
  // Tracks watches list, panel visibility, and live cell-value updates that
  // mount.ts pipes into `data.cells` after recalc.
  let lastWatches = store.getState().watch.watches;
  let lastVisible = store.getState().ui.watchPanelOpen;
  let lastCells = store.getState().data.cells;
  let lastSheetIdx = store.getState().data.sheetIndex;
  let lastActive = store.getState().selection.active;
  const unsub = store.subscribe(() => {
    const s = store.getState();
    const visible = s.ui.watchPanelOpen;
    if (visible !== lastVisible) {
      lastVisible = visible;
      root.hidden = !visible;
      syncDock();
      if (visible) {
        refresh();
        if (restoreFocusEl) focusInitial();
      }
    }
    const watchesChanged = s.watch.watches !== lastWatches;
    const cellsChanged = s.data.cells !== lastCells;
    const sheetChanged = s.data.sheetIndex !== lastSheetIdx;
    const activeChanged = s.selection.active !== lastActive;
    if (watchesChanged) lastWatches = s.watch.watches;
    if (cellsChanged) lastCells = s.data.cells;
    if (sheetChanged) lastSheetIdx = s.data.sheetIndex;
    if (activeChanged) lastActive = s.selection.active;
    if (visible && (watchesChanged || cellsChanged || sheetChanged || activeChanged)) refresh();
  });

  if (!root.hidden) refresh();

  return {
    open,
    close,
    toggle,
    refresh,
    detach(): void {
      addBtn.removeEventListener('click', onAdd);
      removeTool?.removeEventListener('click', onRemoveSelected);
      clearBtn?.removeEventListener('click', onClear);
      closeBtn.removeEventListener('click', onClose);
      detachPaneEscape();
      detachRovingRows();
      unsub();
      root.remove();
      syncDock();
    },
    setStrings(next: Strings): void {
      strings = next;
      title.textContent = mac ? macLabels().title : strings.watchPanel.title;
      if (!mac) addBtn.textContent = strings.watchPanel.addWatch;
      if (clearBtn) clearBtn.textContent = strings.watchPanel.clearAll;
      closeBtn.setAttribute('aria-label', mac ? macLabels().close : strings.watchPanel.close);
      labelMacTools();
      empty.textContent = strings.watchPanel.empty;
      const ths = headRow.querySelectorAll('th');
      ths.forEach((th, i) => {
        const key = headers[i];
        if (key) th.textContent = strings.watchPanel[key];
      });
      root.setAttribute('aria-label', mac ? macLabels().title : strings.watchPanel.title);
      refresh();
    },
  } as WatchPanelHandle & { setStrings(next: Strings): void };
}
