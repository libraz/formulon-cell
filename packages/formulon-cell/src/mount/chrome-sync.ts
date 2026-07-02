import { isValidDefinedName, upsertDefinedName } from '../commands/named-ranges.js';
import { formatA1FormulaAsR1C1 } from '../commands/refs.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { type SpreadsheetEmitter, selectionEquals } from '../events.js';
import type { Strings } from '../i18n/strings.js';
import type { SpreadsheetStore, State } from '../store/store.js';
import { mutators } from '../store/store.js';
import { createHostButton } from './chrome-buttons.js';
import {
  colName,
  formatSelectionRef,
  lookupDefinedName,
  parseCellRef,
  parseRangeRef,
} from './ref-utils.js';
import type { SheetTabsController } from './sheet-tabs-controller.js';

interface AttachChromeSyncInput {
  a11y: HTMLElement;
  fxInput: HTMLTextAreaElement;
  getFormulaEditing: () => boolean;
  getSheetTabs: () => SheetTabsController | null;
  getStrings: () => Strings;
  getWb: () => WorkbookHandle;
  grid: HTMLElement;
  host: HTMLElement;
  invalidate: () => void;
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

const absoluteCellRef = (row: number, col: number): string => `$${colName(col)}$${row + 1}`;
const MAX_A11Y_VIEWPORT_ROWS = 20;
const MAX_A11Y_VIEWPORT_COLS = 10;

const selectionFormula = (
  sheetName: string,
  range: { r0: number; c0: number; r1: number; c1: number },
): string => {
  const start = absoluteCellRef(range.r0, range.c0);
  const end = absoluteCellRef(range.r1, range.c1);
  const ref = start === end ? start : `${start}:${end}`;
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
  if (formula) return state.ui.r1c1 ? formatA1FormulaAsR1C1(formula, addr) : formula;
  if (!cell) {
    const lambda = wb.getLambdaText(addr);
    return lambda ? `=${lambda}` : '';
  }
  const v = cell.value;
  switch (v.kind) {
    case 'number':
      return String(v.value);
    case 'bool':
      return v.value ? 'TRUE' : 'FALSE';
    case 'text':
      return v.value;
    case 'error':
      return v.text;
    default: {
      const lambda = wb.getLambdaText(addr);
      return lambda ? `=${lambda}` : '';
    }
  }
};

export function attachChromeSync(input: AttachChromeSyncInput): ChromeSyncController {
  const {
    a11y,
    emitter,
    fxInput,
    getFormulaEditing,
    getSheetTabs,
    getStrings,
    getWb,
    grid,
    host,
    invalidate,
    store,
    tag,
  } = input;
  const activeCellMirror = document.createElement('div');
  activeCellMirror.id = `${a11y.id || 'fc-a11y'}-active-cell`;
  activeCellMirror.setAttribute('role', 'gridcell');
  activeCellMirror.setAttribute('aria-selected', 'true');
  const viewportMirror = document.createElement('div');
  viewportMirror.id = `${a11y.id || 'fc-a11y'}-viewport`;
  viewportMirror.setAttribute('role', 'rowgroup');
  a11y.replaceChildren(activeCellMirror, viewportMirror);
  grid.setAttribute('aria-activedescendant', activeCellMirror.id);
  grid.setAttribute('aria-rowcount', String(1_048_576));
  grid.setAttribute('aria-colcount', String(16_384));
  grid.setAttribute('aria-multiselectable', 'true');

  const updateChrome = (): void => {
    const wb = getWb();
    const s = store.getState();
    host.dataset.fcWorkbookView = s.ui.workbookView;
    const a = s.selection.active;
    const ref = formatSelectionRef(s.selection.range, a, s.ui.r1c1 === true);
    if (document.activeElement !== tag) tag.value = ref;
    const display = cellDisplayText(s, wb, a);
    if (!getFormulaEditing()) fxInput.value = display;
    activeCellMirror.setAttribute('aria-rowindex', String(a.row + 1));
    activeCellMirror.setAttribute('aria-colindex', String(a.col + 1));
    activeCellMirror.setAttribute('aria-label', display ? `${ref} ${display}` : ref);
    activeCellMirror.textContent = `${ref} ${display}`;
    const ownedIds = [activeCellMirror.id];
    const cells: HTMLElement[] = [];
    const rowEnd = Math.min(
      1_048_575,
      s.viewport.rowStart + Math.max(1, Math.min(s.viewport.rowCount, MAX_A11Y_VIEWPORT_ROWS)) - 1,
    );
    const colEnd = Math.min(
      16_383,
      s.viewport.colStart + Math.max(1, Math.min(s.viewport.colCount, MAX_A11Y_VIEWPORT_COLS)) - 1,
    );
    for (let row = s.viewport.rowStart; row <= rowEnd; row += 1) {
      for (let col = s.viewport.colStart; col <= colEnd; col += 1) {
        if (row === a.row && col === a.col) continue;
        const cellMirror = document.createElement('div');
        const cellRef = `${colName(col)}${row + 1}`;
        const cellDisplay = cellDisplayText(s, wb, { sheet: s.data.sheetIndex, row, col });
        cellMirror.id = `${viewportMirror.id}-cell-${row}-${col}`;
        cellMirror.setAttribute('role', 'gridcell');
        cellMirror.setAttribute('aria-rowindex', String(row + 1));
        cellMirror.setAttribute('aria-colindex', String(col + 1));
        cellMirror.setAttribute('aria-selected', 'false');
        cellMirror.setAttribute('aria-label', cellDisplay ? `${cellRef} ${cellDisplay}` : cellRef);
        cellMirror.textContent = `${cellRef} ${cellDisplay}`;
        ownedIds.push(cellMirror.id);
        cells.push(cellMirror);
      }
    }
    viewportMirror.replaceChildren(...cells);
    grid.setAttribute('aria-owns', ownedIds.join(' '));
  };

  let nameMenu: HTMLDivElement | null = null;
  const definedNameRows = (): { name: string; formula: string }[] =>
    [...getWb().definedNames()]
      .filter((dn) => dn.name.trim() && dn.formula.trim())
      .sort((a, b) => a.name.localeCompare(b.name));

  const closeNameMenu = (): void => {
    nameMenu?.remove();
    nameMenu = null;
    document.removeEventListener('pointerdown', onNameMenuDocPointer, true);
    document.removeEventListener('keydown', onNameMenuDocKey, true);
  };

  const resolveNameBoxRange = (
    raw: string,
    sheetIdx: number,
  ): { sheet: number; r0: number; c0: number; r1: number; c1: number } | null => {
    const asRange = (range: { r0: number; c0: number; r1: number; c1: number }) => ({
      sheet: sheetIdx,
      ...range,
    });
    const asCell = (cell: { row: number; col: number }) => ({
      sheet: sheetIdx,
      r0: cell.row,
      c0: cell.col,
      r1: cell.row,
      c1: cell.col,
    });
    const range = parseRangeRef(raw);
    if (range) return asRange(range);
    const parsed = parseCellRef(raw);
    if (parsed) return asCell(parsed);
    const dn = lookupDefinedName(getWb(), raw.trim());
    if (!dn) return null;
    const subRange = parseRangeRef(dn);
    if (subRange) return asRange(subRange);
    const subCell = parseCellRef(dn);
    return subCell ? asCell(subCell) : null;
  };

  const applyNameBoxValue = (raw: string): boolean => {
    const sheetIdx = store.getState().data.sheetIndex;
    const range = resolveNameBoxRange(raw, sheetIdx);
    if (!range) return false;
    const collapsed = range.r0 === range.r1 && range.c0 === range.c1;
    if (collapsed) {
      mutators.setActive(store, {
        sheet: sheetIdx,
        row: range.r0,
        col: range.c0,
      });
    } else {
      store.setState((s) => ({
        ...s,
        selection: {
          active: { sheet: sheetIdx, row: range.r0, col: range.c0 },
          anchor: { sheet: sheetIdx, row: range.r0, col: range.c0 },
          range,
          extraRanges: [],
        },
      }));
    }
    host.focus();
    return true;
  };

  const defineNameBoxValue = (raw: string): boolean => {
    if (!isValidDefinedName(raw)) return false;
    const wb = getWb();
    const state = store.getState();
    const formula = selectionFormula(wb.sheetName(state.data.sheetIndex), state.selection.range);
    const result = upsertDefinedName(wb, raw, formula);
    if (!result.ok) return false;
    host.focus();
    return true;
  };

  const addNameBoxValue = (raw: string): boolean => {
    const sheetIdx = store.getState().data.sheetIndex;
    const range = resolveNameBoxRange(raw, sheetIdx);
    if (!range) return false;
    const collapsed = range.r0 === range.r1 && range.c0 === range.c1;
    if (collapsed) {
      mutators.addExtraCell(store, { sheet: sheetIdx, row: range.r0, col: range.c0 });
    } else {
      mutators.addExtraRange(store, range, { sheet: sheetIdx, row: range.r0, col: range.c0 });
    }
    return true;
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
    const rows = definedNameRows();
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
      item.title = row.formula;
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
    document.body.appendChild(menu);
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
      if (!applyNameBoxValue(tag.value)) defineNameBoxValue(tag.value);
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
      getWb().clearViewportHint();
      lastSheetIdx = s.data.sheetIndex;
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
