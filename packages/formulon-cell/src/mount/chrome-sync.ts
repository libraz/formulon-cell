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
import { isDataTableFormula } from '../engine/workbook-handle-data-tables.js';
import { type SpreadsheetEmitter, selectionEquals } from '../events.js';
import type { Strings } from '../i18n/strings.js';
import { scrollActiveIntoView } from '../interact/navigation-policy.js';
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
  if (formula) {
    const text = formulaForEdit(wb, addr, formula, state.ui.r1c1) ?? formula;
    // Data table results read as the array formula the desktop app shows.
    return isDataTableFormula(formula) ? `{${text}}` : text;
  }
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
  const activeOnlyRowMirror = createRowMirror(-1);
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

  interface MirrorViewport {
    sheet: number;
    rowStart: number;
    rowEnd: number;
    colStart: number;
    colEnd: number;
  }
  interface ContentSnapshot {
    cells: State['data']['cells'];
    formats: State['format']['formats'];
    protectedSheets: State['protection']['protectedSheets'];
    r1c1: boolean | undefined;
    wb: WorkbookHandle;
    profileId: ReturnType<WorkbookHandle['spreadsheetProfileId']>;
    inputLocale: string | undefined;
    strings: Strings;
  }

  const mirrorViewport = (s: State): MirrorViewport => ({
    sheet: s.data.sheetIndex,
    rowStart: s.viewport.rowStart,
    rowEnd: Math.min(
      MAX_ROW,
      s.viewport.rowStart + Math.max(1, Math.min(s.viewport.rowCount, MAX_A11Y_VIEWPORT_ROWS)) - 1,
    ),
    colStart: s.viewport.colStart,
    colEnd: Math.min(
      MAX_COL,
      s.viewport.colStart + Math.max(1, Math.min(s.viewport.colCount, MAX_A11Y_VIEWPORT_COLS)) - 1,
    ),
  });

  const sameViewport = (left: MirrorViewport | undefined, right: MirrorViewport): boolean =>
    left?.sheet === right.sheet &&
    left.rowStart === right.rowStart &&
    left.rowEnd === right.rowEnd &&
    left.colStart === right.colStart &&
    left.colEnd === right.colEnd;

  const contentSnapshot = (s: State, wb: WorkbookHandle): ContentSnapshot => ({
    cells: s.data.cells,
    formats: s.format.formats,
    protectedSheets: s.protection.protectedSheets,
    r1c1: s.ui.r1c1,
    wb,
    profileId: wb.spreadsheetProfileId(),
    inputLocale: inputLocaleForWorkbook(wb),
    strings: getStrings(),
  });

  const sameContent = (left: ContentSnapshot | undefined, right: ContentSnapshot): boolean =>
    left?.cells === right.cells &&
    left.formats === right.formats &&
    left.protectedSheets === right.protectedSheets &&
    left.r1c1 === right.r1c1 &&
    left.wb === right.wb &&
    left.profileId === right.profileId &&
    left.inputLocale === right.inputLocale &&
    left.strings === right.strings;

  const cellKey = (sheet: number, row: number, col: number): string => `${sheet}:${row}:${col}`;
  const ordinaryCells = new Map<string, HTMLElement>();
  const ordinaryDisplays = new Map<string, string>();
  const rowMirrors = new Map<number, HTMLElement>();
  let currentViewport: MirrorViewport | undefined;
  let currentContent: ContentSnapshot | undefined;
  let activePlaceholder: HTMLElement | null = null;
  let activeCoordinateKey: string | undefined;
  let currentActiveDisplay: { key: string; text: string } | undefined;

  const setOrdinaryCellText = (
    node: HTMLElement,
    row: number,
    col: number,
    display: string,
  ): void => {
    const cellRef = formatA1Cell(row, col);
    node.id = `${mirrorId}-cell-${row}-${col}`;
    node.setAttribute('role', 'gridcell');
    node.setAttribute('aria-rowindex', String(row + 1));
    node.setAttribute('aria-colindex', String(col + 1));
    node.setAttribute('aria-selected', 'false');
    node.setAttribute('aria-label', display ? `${cellRef} ${display}` : cellRef);
    node.textContent = `${cellRef} ${display}`;
  };

  const reconcileChildren = (parent: HTMLElement, desired: readonly HTMLElement[]): void => {
    const desiredSet = new Set(desired);
    for (const child of Array.from(parent.children)) {
      if (!desiredSet.has(child as HTMLElement)) child.remove();
    }
    for (let index = 0; index < desired.length; index += 1) {
      const desiredChild = desired[index];
      if (!desiredChild || parent.children.item(index) === desiredChild) continue;
      parent.insertBefore(desiredChild, parent.children.item(index));
    }
  };

  const refreshOrdinaryCells = (s: State, wb: WorkbookHandle, viewport: MirrorViewport): void => {
    ordinaryDisplays.clear();
    for (let row = viewport.rowStart; row <= viewport.rowEnd; row += 1) {
      for (let col = viewport.colStart; col <= viewport.colEnd; col += 1) {
        const key = cellKey(viewport.sheet, row, col);
        const node = ordinaryCells.get(key);
        if (!node) continue;
        const display = cellDisplayText(s, wb, { sheet: viewport.sheet, row, col });
        ordinaryDisplays.set(key, display);
        setOrdinaryCellText(node, row, col, display);
      }
    }
  };

  const rebuildMirror = (
    s: State,
    wb: WorkbookHandle,
    viewport: MirrorViewport,
    refreshAll: boolean,
  ): void => {
    const sameSheet = currentViewport?.sheet === viewport.sheet;
    if (!sameSheet) {
      ordinaryCells.clear();
      ordinaryDisplays.clear();
      rowMirrors.clear();
      activePlaceholder = null;
    } else if (activePlaceholder && activeCellMirror.parentElement) {
      activeCellMirror.replaceWith(activePlaceholder);
      activePlaceholder = null;
    }
    activeCoordinateKey = cellKey(
      s.selection.active.sheet,
      s.selection.active.row,
      s.selection.active.col,
    );
    const rowOrder: number[] = [];
    const nextCellKeys = new Set<string>();
    const nextRows = new Set<number>();
    let activePlaced = false;
    for (let row = viewport.rowStart; row <= viewport.rowEnd; row += 1) {
      let rowMirror = rowMirrors.get(row);
      if (!rowMirror) {
        rowMirror = createRowMirror(row);
        rowMirrors.set(row, rowMirror);
      }
      nextRows.add(row);
      rowOrder.push(row);
      const cells: HTMLElement[] = [];
      for (let col = viewport.colStart; col <= viewport.colEnd; col += 1) {
        const key = cellKey(viewport.sheet, row, col);
        nextCellKeys.add(key);
        let cellMirror = ordinaryCells.get(key);
        if (!cellMirror) {
          cellMirror = document.createElement('div');
          ordinaryCells.set(key, cellMirror);
        }
        if (refreshAll || !ordinaryDisplays.has(key)) {
          const display = cellDisplayText(s, wb, { sheet: viewport.sheet, row, col });
          ordinaryDisplays.set(key, display);
          setOrdinaryCellText(cellMirror, row, col, display);
        }
        if (row === s.selection.active.row && col === s.selection.active.col) {
          activePlaceholder = cellMirror;
          cells.push(activeCellMirror);
          activePlaced = true;
        } else {
          cells.push(cellMirror);
        }
      }
      if (row === s.selection.active.row && !activePlaced) {
        if (s.selection.active.col < viewport.colStart) cells.unshift(activeCellMirror);
        else cells.push(activeCellMirror);
        activePlaced = true;
      }
      reconcileChildren(rowMirror, cells);
    }
    for (const key of ordinaryCells.keys()) {
      if (nextCellKeys.has(key)) continue;
      ordinaryCells.delete(key);
      ordinaryDisplays.delete(key);
    }
    for (const row of rowMirrors.keys()) {
      if (!nextRows.has(row)) rowMirrors.delete(row);
    }
    if (!activePlaced) {
      activeOnlyRowMirror.id = `${mirrorId}-row-${s.selection.active.row}`;
      activeOnlyRowMirror.setAttribute('aria-rowindex', String(s.selection.active.row + 1));
      reconcileChildren(activeOnlyRowMirror, [activeCellMirror]);
      if (s.selection.active.row < viewport.rowStart) rowOrder.unshift(s.selection.active.row);
      else rowOrder.push(s.selection.active.row);
    }
    const orderedRows = rowOrder.map((row) =>
      row === s.selection.active.row && !rowMirrors.has(row)
        ? activeOnlyRowMirror
        : (rowMirrors.get(row) as HTMLElement),
    );
    reconcileChildren(a11y, orderedRows);
    currentViewport = viewport;
  };

  const moveActiveMirror = (s: State, viewport: MirrorViewport): void => {
    const active = s.selection.active;
    const nextKey = cellKey(active.sheet, active.row, active.col);
    if (nextKey === activeCoordinateKey) return;

    if (activePlaceholder && activeCellMirror.parentElement) {
      activeCellMirror.replaceWith(activePlaceholder);
    } else if (activeCellMirror.parentElement !== activeOnlyRowMirror) {
      activeCellMirror.remove();
    }
    activePlaceholder = null;
    activeCoordinateKey = nextKey;

    const targetPlaceholder = ordinaryCells.get(nextKey);
    if (targetPlaceholder) {
      targetPlaceholder.replaceWith(activeCellMirror);
      activePlaceholder = targetPlaceholder;
      activeOnlyRowMirror.remove();
      return;
    }

    if (active.row >= viewport.rowStart && active.row <= viewport.rowEnd) {
      activeOnlyRowMirror.remove();
      const rowMirror = rowMirrors.get(active.row);
      if (!rowMirror) return;
      if (active.col < viewport.colStart)
        rowMirror.insertBefore(activeCellMirror, rowMirror.firstChild);
      else rowMirror.append(activeCellMirror);
      return;
    }

    activeOnlyRowMirror.id = `${mirrorId}-row-${active.row}`;
    activeOnlyRowMirror.setAttribute('aria-rowindex', String(active.row + 1));
    reconcileChildren(activeOnlyRowMirror, [activeCellMirror]);
    if (active.row < viewport.rowStart) {
      if (a11y.firstElementChild !== activeOnlyRowMirror)
        a11y.insertBefore(activeOnlyRowMirror, a11y.firstChild);
    } else if (a11y.lastElementChild !== activeOnlyRowMirror) {
      a11y.append(activeOnlyRowMirror);
    }
  };

  const refreshChrome = (forceContent: boolean): void => {
    const wb = getWb();
    const s = store.getState();
    const viewport = mirrorViewport(s);
    const snapshot = contentSnapshot(s, wb);
    const viewportChanged = !sameViewport(currentViewport, viewport);
    const contentChanged = forceContent || !sameContent(currentContent, snapshot);
    host.dataset.fcWorkbookView = s.ui.workbookView;

    if (viewportChanged) {
      rebuildMirror(s, wb, viewport, contentChanged);
    } else if (contentChanged) {
      refreshOrdinaryCells(s, wb, viewport);
      moveActiveMirror(s, viewport);
    } else {
      moveActiveMirror(s, viewport);
    }
    if (contentChanged) currentContent = snapshot;

    const active = s.selection.active;
    const activeKey = cellKey(active.sheet, active.row, active.col);
    let display = ordinaryDisplays.get(activeKey);
    if (display === undefined) {
      if (currentActiveDisplay?.key === activeKey && !contentChanged) {
        display = currentActiveDisplay.text;
      } else {
        display = cellDisplayText(s, wb, active);
      }
    }
    currentActiveDisplay = { key: activeKey, text: display };
    const ref = selectionRef(s);
    const object = s.ui.selectedObject;
    if (document.activeElement !== tag) tag.value = object ? object.name : nameBoxText(ref);
    if (!getFormulaEditing()) {
      fxInput.value = object ? '' : display;
      fxInput.readOnly =
        !object &&
        isDataTableFormula(
          s.data.cells.get(`${active.sheet}:${active.row}:${active.col}`)?.formula,
        );
    }
    activeCellMirror.setAttribute('aria-rowindex', String(active.row + 1));
    activeCellMirror.setAttribute('aria-colindex', String(active.col + 1));
    activeCellMirror.setAttribute('aria-label', display ? `${ref} ${display}` : ref);
    activeCellMirror.textContent = `${ref} ${display}`;
    a11yLive.textContent = `${ref} ${display}`;
  };

  const updateChrome = (): void => refreshChrome(true);

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
    scrollActiveIntoView(store, store.getState().selection.active);
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
    const rows = rangeDefinedNames(getWb(), store.getState().data.sheetIndex).filter(
      (row) => !row.hidden,
    );
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
    refreshChrome(false);
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
