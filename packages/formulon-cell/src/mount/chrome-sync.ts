import { inputLocaleForWorkbook } from '../commands/coerce-input.js';
import type { History } from '../commands/history.js';
import {
  definedNameDisplayFormula,
  isValidDefinedName,
  recordDefinedNamesChange,
  upsertDefinedName,
} from '../commands/named-ranges.js';
import { formatA1Cell, formatA1Range, MAX_COL, MAX_ROW } from '../engine/address.js';
import { formulaForEdit } from '../engine/edit-seed.js';
import { syncActiveSheet } from '../engine/layout-sync.js';
import { formatBoolean, formatEditableNumber } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { type SpreadsheetEmitter, selectionEquals } from '../events.js';
import type { Strings } from '../i18n/strings.js';
import { overlayPortalFor } from '../interact/overlay-portal.js';
import type { SpreadsheetStore, State } from '../store/store.js';
import { mutators } from '../store/store.js';
import { createHostButton } from './chrome-buttons.js';
import {
  formatSelectionRef,
  nameForRange,
  rangeDefinedNames,
  resolveAreaText,
  type SheetRange,
  splitAreas,
} from './ref-utils.js';
import type { SheetTabsController } from './sheet-tabs-controller.js';

interface AttachChromeSyncInput {
  a11y: HTMLElement;
  a11yLive: HTMLElement;
  fxInput: HTMLTextAreaElement;
  getFormulaEditing: () => boolean;
  getSheetTabs: () => SheetTabsController | null;
  getStrings: () => Strings;
  getWb: () => WorkbookHandle;
  /** Shared undo history; names defined from the box are recorded on it. */
  getHistory?: () => History | null;
  grid: HTMLElement;
  host: HTMLElement;
  invalidate: () => void;
  /** Surfaces a modal-style notice such as the invalid-reference message. */
  onAlert?: (outcome: { severity: 'stop'; message: string }) => void;
  store: SpreadsheetStore;
  tag: HTMLInputElement;
  emitter: SpreadsheetEmitter;
}

export interface ChromeSyncController {
  detach(): void;
  updateChrome(): void;
}

const quoteSheetName = (name: string): string => {
  if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) return name;
  return `'${name.replace(/'/g, "''")}'`;
};

const MAX_A11Y_VIEWPORT_ROWS = 20;
const MAX_A11Y_VIEWPORT_COLS = 10;

const selectionFormula = (
  sheetName: string,
  range: { r0: number; c0: number; r1: number; c1: number },
): string => {
  const ref = formatA1Range(range, { absolute: true });
  return `=${quoteSheetName(sheetName)}!${ref}`;
};

const cellDisplayText = (
  state: State,
  wb: WorkbookHandle,
  addr: { sheet: number; row: number; col: number },
): string => {
  const key = `${addr.sheet}:${addr.row}:${addr.col}`;
  const cell = state.data.cells.get(key);
  const fmt = state.format.formats.get(key);
  const formula = cell?.formula ?? '';
  if (formula && fmt?.formulaHidden === true && state.protection.protectedSheets.has(addr.sheet)) {
    return '';
  }
  if (formula) return formulaForEdit(wb, addr, formula, state.ui.r1c1) ?? formula;
  if (!cell) {
    const lambda = wb.getLambdaText(addr);
    return lambda ? wb.getFormulaForDisplay(`=${lambda}`) : '';
  }
  const v = cell.value;
  switch (v.kind) {
    case 'number':
      return formatEditableNumber(v.value, inputLocaleForWorkbook(wb));
    case 'bool':
      return formatBoolean(v.value, inputLocaleForWorkbook(wb));
    case 'text':
      return v.value;
    case 'error':
      return v.text;
    default: {
      const lambda = wb.getLambdaText(addr);
      return lambda ? wb.getFormulaForDisplay(`=${lambda}`) : '';
    }
  }
};

export function attachChromeSync(input: AttachChromeSyncInput): ChromeSyncController {
  const {
    a11y,
    a11yLive,
    emitter,
    fxInput,
    getFormulaEditing,
    getSheetTabs,
    getStrings,
    getHistory,
    getWb,
    grid,
    host,
    invalidate,
    onAlert,
    store,
    tag,
  } = input;
  const mirrorId = a11y.id || 'fc-a11y';
  const activeCellMirror = document.createElement('div');
  activeCellMirror.id = `${mirrorId}-active-cell`;
  activeCellMirror.setAttribute('role', 'gridcell');
  activeCellMirror.setAttribute('aria-selected', 'true');
  /** One `role="row"` wrapper — `gridcell` is only valid inside a row. */
  const createRowMirror = (row: number): HTMLElement => {
    const rowMirror = document.createElement('div');
    rowMirror.id = `${mirrorId}-row-${row}`;
    rowMirror.setAttribute('role', 'row');
    rowMirror.setAttribute('aria-rowindex', String(row + 1));
    return rowMirror;
  };
  grid.setAttribute('aria-activedescendant', activeCellMirror.id);
  grid.setAttribute('aria-rowcount', String(MAX_ROW + 1));
  grid.setAttribute('aria-colcount', String(MAX_COL + 1));
  grid.setAttribute('aria-multiselectable', 'true');

  /** The address, or the defined name when the selection is exactly a named range. */
  const nameBoxText = (ref: string): string => {
    const s = store.getState();
    if (s.selection.extraRanges?.length) return ref;
    return nameForRange(getWb(), s.data.sheetIndex, s.selection.range) ?? ref;
  };

  const selectionRef = (s: State): string => {
    const a = s.selection.active;
    const range = s.selection.range;
    const merge = s.merges.byAnchor.get(`${a.sheet}:${a.row}:${a.col}`);
    const logicalCell =
      merge &&
      merge.sheet === range.sheet &&
      merge.r0 === range.r0 &&
      merge.c0 === range.c0 &&
      merge.r1 === range.r1 &&
      merge.c1 === range.c1;
    return formatSelectionRef(
      logicalCell ? { sheet: a.sheet, r0: a.row, c0: a.col, r1: a.row, c1: a.col } : range,
      a,
      s.ui.r1c1 === true,
    );
  };

  const updateChrome = (): void => {
    const wb = getWb();
    const s = store.getState();
    host.dataset.fcWorkbookView = s.ui.workbookView;
    const a = s.selection.active;
    const ref = selectionRef(s);
    // A selected chart or picture names itself and leaves the formula bar empty.
    const object = s.ui.selectedObject;
    if (document.activeElement !== tag) tag.value = object ? object.name : nameBoxText(ref);
    const display = cellDisplayText(s, wb, a);
    if (!getFormulaEditing()) fxInput.value = object ? '' : display;
    activeCellMirror.setAttribute('aria-rowindex', String(a.row + 1));
    activeCellMirror.setAttribute('aria-colindex', String(a.col + 1));
    activeCellMirror.setAttribute('aria-label', display ? `${ref} ${display}` : ref);
    activeCellMirror.textContent = `${ref} ${display}`;
    a11yLive.textContent = `${ref} ${display}`;
    const rowMirrors: HTMLElement[] = [];
    let activePlaced = false;
    const rowEnd = Math.min(
      MAX_ROW,
      s.viewport.rowStart + Math.max(1, Math.min(s.viewport.rowCount, MAX_A11Y_VIEWPORT_ROWS)) - 1,
    );
    const colEnd = Math.min(
      MAX_COL,
      s.viewport.colStart + Math.max(1, Math.min(s.viewport.colCount, MAX_A11Y_VIEWPORT_COLS)) - 1,
    );
    for (let row = s.viewport.rowStart; row <= rowEnd; row += 1) {
      const cells: HTMLElement[] = [];
      for (let col = s.viewport.colStart; col <= colEnd; col += 1) {
        if (row === a.row && col === a.col) {
          cells.push(activeCellMirror);
          activePlaced = true;
          continue;
        }
        const cellMirror = document.createElement('div');
        const cellRef = formatA1Cell(row, col);
        const cellDisplay = cellDisplayText(s, wb, { sheet: s.data.sheetIndex, row, col });
        cellMirror.id = `${mirrorId}-cell-${row}-${col}`;
        cellMirror.setAttribute('role', 'gridcell');
        cellMirror.setAttribute('aria-rowindex', String(row + 1));
        cellMirror.setAttribute('aria-colindex', String(col + 1));
        cellMirror.setAttribute('aria-selected', 'false');
        cellMirror.setAttribute('aria-label', cellDisplay ? `${cellRef} ${cellDisplay}` : cellRef);
        cellMirror.textContent = `${cellRef} ${cellDisplay}`;
        cells.push(cellMirror);
      }
      // The active cell can sit outside the mirrored column window; it still
      // belongs to its own row so `aria-activedescendant` stays resolvable.
      if (row === a.row && !activePlaced) {
        if (a.col < s.viewport.colStart) cells.unshift(activeCellMirror);
        else cells.push(activeCellMirror);
        activePlaced = true;
      }
      const rowMirror = createRowMirror(row);
      rowMirror.replaceChildren(...cells);
      rowMirrors.push(rowMirror);
    }
    if (!activePlaced) {
      const activeRow = createRowMirror(a.row);
      activeRow.replaceChildren(activeCellMirror);
      if (a.row < s.viewport.rowStart) rowMirrors.unshift(activeRow);
      else rowMirrors.push(activeRow);
    }
    a11y.replaceChildren(...rowMirrors);
  };

  let nameMenu: HTMLDivElement | null = null;
  const closeNameMenu = (): void => {
    nameMenu?.remove();
    nameMenu = null;
    document.removeEventListener('pointerdown', onNameMenuDocPointer, true);
    document.removeEventListener('keydown', onNameMenuDocKey, true);
  };

  const resolveNameBoxAreas = (raw: string): SheetRange[] | null => {
    const wb = getWb();
    const sheetIdx = store.getState().data.sheetIndex;
    const areas: SheetRange[] = [];
    for (const text of splitAreas(raw)) {
      const area = resolveAreaText(wb, text, sheetIdx);
      if (!area) return null;
      areas.push(area);
    }
    const first = areas[0];
    if (!first || areas.some((area) => area.sheet !== first.sheet)) return null;
    return areas;
  };

  const applyNameBoxValue = (raw: string): boolean => {
    const areas = resolveNameBoxAreas(raw);
    const first = areas?.[0];
    if (!areas || !first) return false;
    const sheetIdx = first.sheet;
    if (sheetIdx !== store.getState().data.sheetIndex) {
      const tabs = getSheetTabs();
      if (tabs) tabs.switchSheet(sheetIdx);
      else mutators.setSheetIndex(store, sheetIdx);
      if (store.getState().data.sheetIndex !== sheetIdx) return false;
    }
    if (areas.length === 1 && first.r0 === first.r1 && first.c0 === first.c1) {
      mutators.setActive(store, { sheet: sheetIdx, row: first.r0, col: first.c0 });
    } else {
      const start = { sheet: sheetIdx, row: first.r0, col: first.c0 };
      store.setState((s) => ({
        ...s,
        selection: { active: start, anchor: start, range: first, extraRanges: areas.slice(1) },
      }));
    }
    host.focus();
    return true;
  };

  const defineNameBoxValue = (raw: string): boolean => {
    if (!isValidDefinedName(raw)) return false;
    const wb = getWb();
    const state = store.getState();
    const formula = selectionFormula(
      wb.sheetName(state.selection.range.sheet),
      state.selection.range,
    );
    const result = recordDefinedNamesChange(getHistory?.() ?? null, wb, () =>
      upsertDefinedName(wb, raw, formula),
    );
    if (!result.ok) return false;
    host.focus();
    return true;
  };

  const addNameBoxValue = (raw: string): boolean => {
    const areas = resolveNameBoxAreas(raw);
    const range = areas?.[0];
    if (areas?.length !== 1 || !range) return false;
    const sheetIdx = store.getState().data.sheetIndex;
    if (range.sheet !== sheetIdx) return false;
    const collapsed = range.r0 === range.r1 && range.c0 === range.c1;
    if (collapsed) {
      mutators.addExtraCell(store, { sheet: sheetIdx, row: range.r0, col: range.c0 });
    } else {
      mutators.addExtraRange(store, range, { sheet: sheetIdx, row: range.r0, col: range.c0 });
    }
    return true;
  };

  /** Tell the user the entry is unusable, then put the box back as it was. */
  const rejectNameBoxValue = (): void => {
    onAlert?.({ severity: 'stop', message: getStrings().a11y.nameBoxInvalidReference });
    tag.value = nameBoxText(selectionRef(store.getState()));
    tag.focus();
    tag.select();
  };

  function onNameMenuDocPointer(e: PointerEvent): void {
    if (!nameMenu) return;
    const target = e.target;
    if (target instanceof Node && (nameMenu.contains(target) || tag.contains(target))) return;
    closeNameMenu();
  }

  function onNameMenuDocKey(e: KeyboardEvent): void {
    if (!nameMenu) return;
    const items = Array.from(nameMenu.querySelectorAll<HTMLButtonElement>('[role="option"]'));
    const active =
      document.activeElement instanceof HTMLButtonElement ? document.activeElement : null;
    const idx = active ? items.indexOf(active) : -1;
    const focusAt = (next: number): void => {
      if (items.length === 0) return;
      e.preventDefault();
      e.stopPropagation();
      const wrapped = (next + items.length) % items.length;
      items[wrapped]?.focus();
    };
    if (e.key === 'Escape') {
      e.preventDefault();
      e.stopPropagation();
      closeNameMenu();
      tag.focus();
    } else if (e.key === 'ArrowDown') {
      focusAt(idx < 0 ? 0 : idx + 1);
    } else if (e.key === 'ArrowUp') {
      focusAt(idx < 0 ? items.length - 1 : idx - 1);
    } else if (e.key === 'Home') {
      focusAt(0);
    } else if (e.key === 'End') {
      focusAt(items.length - 1);
    } else if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      e.stopPropagation();
      (idx >= 0 ? items[idx] : items[0])?.click();
    }
  }

  const openNameMenu = (): void => {
    const rows = rangeDefinedNames(getWb(), store.getState().data.sheetIndex);
    closeNameMenu();
    const menu = document.createElement('div');
    menu.className = 'fc-namebox-menu';
    menu.setAttribute('role', 'listbox');
    menu.setAttribute('aria-label', tag.getAttribute('aria-label') ?? getStrings().a11y.nameBox);
    if (rows.length === 0) {
      const empty = document.createElement('div');
      empty.className = 'fc-namebox-menu__empty';
      empty.textContent = getStrings().ribbonMenu.noDefinedNames;
      menu.appendChild(empty);
    }
    for (const row of rows) {
      const item = createHostButton({
        className: 'fc-namebox-menu__item',
        text: row.name,
      });
      item.setAttribute('role', 'option');
      item.title = definedNameDisplayFormula(row.formula);
      item.addEventListener('click', (e) => {
        tag.value = row.name;
        if (e.ctrlKey || e.metaKey) {
          addNameBoxValue(row.name);
          updateChrome();
          return;
        }
        closeNameMenu();
        applyNameBoxValue(row.name);
      });
      menu.appendChild(item);
    }
    overlayPortalFor(host).appendChild(menu);
    const r = tag.getBoundingClientRect();
    menu.style.left = `${Math.max(4, r.left)}px`;
    menu.style.top = `${r.bottom + 2}px`;
    menu.style.minWidth = `${Math.max(116, r.width)}px`;
    nameMenu = menu;
    document.addEventListener('pointerdown', onNameMenuDocPointer, true);
    document.addEventListener('keydown', onNameMenuDocKey, true);
    menu.querySelector<HTMLButtonElement>('[role="option"]')?.focus();
  };

  const onTagFocus = (): void => tag.select();
  const onTagPointerDown = (e: PointerEvent): void => {
    const rect = tag.getBoundingClientRect();
    if (e.clientX >= rect.right - 24) {
      e.preventDefault();
      openNameMenu();
    }
  };
  const onTagKey = (e: KeyboardEvent): void => {
    if ((e.altKey && e.key === 'ArrowDown') || e.key === 'F4') {
      e.preventDefault();
      e.stopPropagation();
      openNameMenu();
    } else if (e.key === 'Enter') {
      e.preventDefault();
      e.stopPropagation();
      const raw = tag.value;
      if (!applyNameBoxValue(raw) && !defineNameBoxValue(raw)) rejectNameBoxValue();
    } else if (e.key === 'Escape') {
      e.preventDefault();
      e.stopPropagation();
      closeNameMenu();
      host.focus();
      updateChrome();
    }
  };
  const onTagBlur = (): void => {
    if (!nameMenu) updateChrome();
  };

  tag.addEventListener('focus', onTagFocus);
  tag.addEventListener('pointerdown', onTagPointerDown);
  tag.addEventListener('keydown', onTagKey);
  tag.addEventListener('blur', onTagBlur);

  let lastSheetIdx = store.getState().data.sheetIndex;
  let lastHiddenSheets = store.getState().layout.hiddenSheets;
  let lastSelection = store.getState().selection;
  const unsub = store.subscribe(() => {
    const s = store.getState();
    const sheetChanged = s.data.sheetIndex !== lastSheetIdx;
    if (sheetChanged) {
      lastSheetIdx = s.data.sheetIndex;
      syncActiveSheet(getWb(), lastSheetIdx);
    }
    if (sheetChanged || s.layout.hiddenSheets !== lastHiddenSheets) {
      lastHiddenSheets = s.layout.hiddenSheets;
      getSheetTabs()?.update();
    }
    if (!selectionEquals(lastSelection, s.selection)) {
      lastSelection = s.selection;
      emitter.emit('selectionChange', {
        active: s.selection.active,
        anchor: s.selection.anchor,
        range: s.selection.range,
      });
    }
    invalidate();
    updateChrome();
  });

  syncActiveSheet(getWb(), lastSheetIdx);
  updateChrome();

  return {
    detach(): void {
      closeNameMenu();
      tag.removeEventListener('focus', onTagFocus);
      tag.removeEventListener('pointerdown', onTagPointerDown);
      tag.removeEventListener('keydown', onTagKey);
      tag.removeEventListener('blur', onTagBlur);
      unsub();
    },
    updateChrome,
  };
}
