// Sort / filter / advanced-filter / remove-duplicates / text-to-columns actions.
// Extracted from main.ts to keep ribbon wiring slim. The factory pattern lets
// the host pass in live references (instance, status bar, dialog helpers)
// without coupling this module to global state.

import {
  applyAdvancedFilter,
  clearFilter,
  colLetter,
  copyAdvancedFilterResult,
  type FilterDropdownHandle,
  inferAutoFilterRange,
  inferSortHasHeader,
  type Range,
  recordFilterChange,
  recordFormatChange,
  removeDuplicates,
  type SpreadsheetInstance,
  setAutoFilter,
  sortRange,
  textToColumns,
  type toolbarMenuText,
} from '@libraz/formulon-cell';

import {
  showAdvancedFilterDialog,
  showPrompt,
  showRemoveDuplicatesDialog,
  showSortDialog,
} from './dialogs.js';

export interface SortFilterCtx {
  getInst: () => SpreadsheetInstance | null;
  ribbonLang: 'ja' | 'en';
  ribbonMenuText: ReturnType<typeof toolbarMenuText>;
  sheetEl: HTMLElement;
  statusMetric: HTMLElement | null;
  getFilterDropdown: () => FilterDropdownHandle | null;
  focusSheet: () => void;
  refreshWorkbookCells: () => void;
}

export interface SortFilterApi {
  openFilterForSelection: () => void;
  applyAdvancedFilterAction: () => Promise<void>;
  sortSelection: (direction: 'asc' | 'desc') => void;
  customSortSelection: () => Promise<void>;
  removeDuplicateRows: () => Promise<void>;
  splitTextToColumns: (delimiter?: string) => void;
  splitTextToColumnsCustom: () => Promise<void>;
}

export const createSortFilter = (ctx: SortFilterCtx): SortFilterApi => {
  const {
    getInst,
    ribbonLang,
    ribbonMenuText,
    sheetEl,
    statusMetric,
    getFilterDropdown,
    focusSheet,
    refreshWorkbookCells,
  } = ctx;

  const sortTargetRange = (state: ReturnType<SpreadsheetInstance['store']['getState']>): Range => {
    const r = state.selection.range;
    if (r.r0 === r.r1 && r.c0 === r.c1) return inferAutoFilterRange(state);
    return r;
  };

  const sortCellDisplayText = (
    state: ReturnType<SpreadsheetInstance['store']['getState']>,
    row: number,
    col: number,
  ): string => {
    const value = state.data.cells.get(`${state.selection.active.sheet}:${row}:${col}`)?.value;
    if (!value) return '';
    if (value.kind === 'number') return String(value.value);
    if (value.kind === 'text') return value.value;
    if (value.kind === 'bool') return value.value ? 'TRUE' : 'FALSE';
    if (value.kind === 'error') return value.text;
    return '';
  };

  const colFromLetters = (letters: string): number => {
    let col = 0;
    const upper = letters.toUpperCase();
    for (let i = 0; i < upper.length; i += 1) {
      const code = upper.charCodeAt(i);
      if (code < 65 || code > 90) return -1;
      col = col * 26 + (code - 64);
    }
    return col - 1;
  };

  const parseA1Range = (raw: string, sheet: number): Range | null => {
    const normalized = raw.replace(/\$/g, '').trim().toUpperCase();
    const match = normalized.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
    if (!match) return null;
    const c0 = colFromLetters(match[1] ?? '');
    const r0 = Number(match[2]) - 1;
    const c1 = match[3] ? colFromLetters(match[3]) : c0;
    const r1 = match[4] ? Number(match[4]) - 1 : r0;
    if (c0 < 0 || c1 < 0 || r0 < 0 || r1 < 0) return null;
    return {
      sheet,
      r0: Math.min(r0, r1),
      c0: Math.min(c0, c1),
      r1: Math.max(r0, r1),
      c1: Math.max(c0, c1),
    };
  };

  const rangeRef = (range: Range): string => {
    const start = `${colLetter(range.c0)}${range.r0 + 1}`;
    const end = `${colLetter(range.c1)}${range.r1 + 1}`;
    return start === end ? start : `${start}:${end}`;
  };

  const syncStoreCellsToWorkbook = (
    i: SpreadsheetInstance,
    sheet: number,
    row: number,
    col: number,
    height: number,
    width: number,
  ): void => {
    const cells = i.store.getState().data.cells;
    for (let r = row; r < row + height; r += 1) {
      for (let c = col; c < col + width; c += 1) {
        const addr = { sheet, row: r, col: c };
        const cell = cells.get(`${sheet}:${r}:${c}`);
        if (!cell) {
          i.workbook.setBlank(addr);
        } else if (cell.formula) {
          i.workbook.setFormula(addr, cell.formula);
        } else if (cell.value.kind === 'number') {
          i.workbook.setNumber(addr, cell.value.value);
        } else if (cell.value.kind === 'text') {
          i.workbook.setText(addr, cell.value.value);
        } else if (cell.value.kind === 'bool') {
          i.workbook.setBool(addr, cell.value.value);
        } else {
          i.workbook.setBlank(addr);
        }
      }
    }
  };

  const openFilterForSelection = (): void => {
    const i = getInst();
    if (!i) return;
    const r = inferAutoFilterRange(i.store.getState());
    const active = i.store.getState().ui.filterRange;
    const sameActive =
      active != null &&
      active.sheet === r.sheet &&
      active.r0 === r.r0 &&
      active.c0 === r.c0 &&
      active.r1 === r.r1 &&
      active.c1 === r.c1;
    recordFilterChange(i.history, i.store, () => {
      if (sameActive) clearFilter(i.store.getState(), i.store, r);
      else setAutoFilter(i.store, r);
    });
    if (sameActive) {
      focusSheet();
      return;
    }
    const sheetRect = sheetEl.getBoundingClientRect();
    getFilterDropdown()?.open(r, r.c0, { x: sheetRect.left + 80, y: sheetRect.top, h: 32 });
    focusSheet();
  };

  const applyAdvancedFilterAction = async (): Promise<void> => {
    const i = getInst();
    if (!i) return;
    const state = i.store.getState();
    const listInitial = rangeRef(state.ui.filterRange ?? inferAutoFilterRange(state));
    const result = await showAdvancedFilterDialog({
      title: ribbonMenuText.advancedFilterDialogTitle,
      listRangeLabel: ribbonMenuText.advancedFilterListRange,
      criteriaRangeLabel: ribbonMenuText.advancedFilterCriteriaRange,
      copyToLabel: ribbonMenuText.advancedFilterCopyTo,
      uniqueOnlyLabel: ribbonMenuText.advancedFilterUniqueOnly,
      initialListRange: listInitial,
      okLabel: 'OK',
      cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
      validateListRange: (value) =>
        parseA1Range(value, state.selection.active.sheet)
          ? null
          : ribbonLang === 'ja'
            ? 'A1:B10 の形式で入力してください。'
            : 'Enter a list range such as A1:B10.',
      validateCriteriaRange: (value) =>
        parseA1Range(value, state.selection.active.sheet)
          ? null
          : ribbonLang === 'ja'
            ? 'A1:B3 の形式で入力してください。'
            : 'Enter a criteria range such as A1:B3.',
      validateCopyTo: (value) => {
        if (!value.trim()) return null;
        return parseA1Range(value, state.selection.active.sheet)
          ? null
          : ribbonLang === 'ja'
            ? 'A1 の形式で入力してください。'
            : 'Enter a cell such as A1.';
      },
    });
    if (result === null) {
      focusSheet();
      return;
    }
    const listRange = parseA1Range(result.listRange, state.selection.active.sheet);
    const criteriaRange = parseA1Range(result.criteriaRange, state.selection.active.sheet);
    if (!listRange || !criteriaRange) return;
    const copyRange = result.copyTo
      ? parseA1Range(result.copyTo, state.selection.active.sheet)
      : null;
    if (copyRange) {
      let copied = 0;
      i.history.begin();
      try {
        copied = copyAdvancedFilterResult(
          i.store.getState(),
          i.store,
          listRange,
          criteriaRange,
          { sheet: copyRange.sheet, row: copyRange.r0, col: copyRange.c0 },
          { uniqueOnly: result.uniqueOnly },
        );
        syncStoreCellsToWorkbook(
          i,
          copyRange.sheet,
          copyRange.r0,
          copyRange.c0,
          copied,
          listRange.c1 - listRange.c0 + 1,
        );
      } finally {
        i.history.end();
      }
      if (statusMetric) {
        statusMetric.textContent = ribbonMenuText.advancedFilterCopiedStatus.replace(
          '{count}',
          String(copied),
        );
      }
    } else {
      recordFilterChange(i.history, i.store, () => {
        applyAdvancedFilter(i.store.getState(), i.store, listRange, criteriaRange);
      });
    }
    focusSheet();
  };

  const sortSelection = (direction: 'asc' | 'desc'): void => {
    const i = getInst();
    if (!i) return;
    const state = i.store.getState();
    const r = sortTargetRange(state);
    if (r.r0 === r.r1) return;
    const activeCol = state.selection.active.col;
    const byCol = activeCol >= r.c0 && activeCol <= r.c1 ? activeCol : r.c0;
    const hasHeader = inferSortHasHeader(state, r);
    let sorted = false;
    i.history.begin();
    try {
      recordFormatChange(i.history, i.store, () => {
        sorted = sortRange(state, i.store, i.workbook, r, { byCol, direction, hasHeader });
      });
    } finally {
      i.history.end();
    }
    if (sorted) refreshWorkbookCells();
    focusSheet();
  };

  const customSortSelection = async (): Promise<void> => {
    const i = getInst();
    if (!i) return;
    const state = i.store.getState();
    const range = sortTargetRange(state);
    if (range.r0 === range.r1) {
      focusSheet();
      return;
    }
    const inferredHeader = inferSortHasHeader(state, range);
    const activeCol =
      state.selection.active.col >= range.c0 && state.selection.active.col <= range.c1
        ? state.selection.active.col
        : range.c0;
    const columns = Array.from({ length: range.c1 - range.c0 + 1 }, (_, offset) => {
      const col = range.c0 + offset;
      const letter = colLetter(col);
      const header = inferredHeader ? sortCellDisplayText(state, range.r0, col).trim() : '';
      return {
        value: String(col),
        label: header ? `${header} (${letter})` : letter,
      };
    });
    const result = await showSortDialog({
      title: ribbonMenuText.sortCustom,
      columnLabel: ribbonMenuText.sortColumn,
      thenByLabel: ribbonMenuText.sortThenBy,
      noThenByLabel: ribbonMenuText.sortNoThenBy,
      orderLabel: ribbonMenuText.sortOrder,
      headerLabel: ribbonMenuText.sortMyDataHasHeaders,
      ascendingLabel: ribbonMenuText.sortAscendingMenu,
      descendingLabel: ribbonMenuText.sortDescendingMenu,
      columns,
      initialColumn: String(activeCol),
      initialDirection: 'asc',
      initialHasHeader: inferredHeader,
      okLabel: 'OK',
      cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
    });
    if (!result) {
      focusSheet();
      return;
    }
    let sorted = false;
    const byCol = Number(result.column);
    i.history.begin();
    try {
      recordFormatChange(i.history, i.store, () => {
        sorted = sortRange(i.store.getState(), i.store, i.workbook, range, {
          byCol,
          direction: result.direction,
          hasHeader: result.hasHeader,
          keys: result.levels.map((level) => ({
            byCol: Number(level.column),
            direction: level.direction,
          })),
        });
      });
    } finally {
      i.history.end();
    }
    if (sorted) {
      refreshWorkbookCells();
      if (statusMetric) {
        const columnLabel = columns.find((column) => column.value === result.column)?.label ?? '';
        statusMetric.textContent = ribbonMenuText.sortStatus.replace('{column}', columnLabel);
      }
    }
    focusSheet();
  };

  const removeDuplicateRows = async (): Promise<void> => {
    const i = getInst();
    if (!i) return;
    const state = i.store.getState();
    const range = sortTargetRange(state);
    const inferredHeader = inferSortHasHeader(state, range);
    const columns = Array.from({ length: range.c1 - range.c0 + 1 }, (_, offset) => {
      const col = range.c0 + offset;
      const letter = colLetter(col);
      const header = inferredHeader ? sortCellDisplayText(state, range.r0, col).trim() : '';
      return {
        value: String(col),
        label: header ? `${header} (${letter})` : letter,
      };
    });
    const result = await showRemoveDuplicatesDialog({
      title: ribbonMenuText.removeDuplicatesDialogTitle,
      columnsLabel: ribbonMenuText.removeDuplicatesColumns,
      headerLabel: ribbonMenuText.sortMyDataHasHeaders,
      selectAllLabel: ribbonMenuText.removeDuplicatesSelectAll,
      unselectAllLabel: ribbonMenuText.removeDuplicatesUnselectAll,
      noColumnsLabel: ribbonMenuText.removeDuplicatesNoColumns,
      columns,
      initialColumns: columns.map((column) => column.value),
      initialHasHeader: inferredHeader,
      okLabel: 'OK',
      cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
    });
    if (!result) {
      focusSheet();
      return;
    }
    let removed = 0;
    i.history.begin();
    try {
      recordFormatChange(i.history, i.store, () => {
        removed = removeDuplicates(i.store.getState(), i.store, i.workbook, range, {
          columns: result.columns.map(Number),
          hasHeader: result.hasHeader,
        });
      });
    } finally {
      i.history.end();
    }
    if (removed > 0) refreshWorkbookCells();
    if (statusMetric) {
      statusMetric.textContent = ribbonMenuText.removeDuplicatesStatus.replace(
        '{count}',
        String(removed),
      );
    }
    focusSheet();
  };

  const splitTextToColumns = (delimiter = ','): void => {
    const i = getInst();
    if (!i) return;
    const state = i.store.getState();
    let max = 0;
    i.history.begin();
    try {
      recordFormatChange(i.history, i.store, () => {
        max = textToColumns(state, i.store, i.workbook, state.selection.range, delimiter);
      });
    } finally {
      i.history.end();
    }
    if (max > 0) refreshWorkbookCells();
    if (statusMetric)
      statusMetric.textContent =
        max > 0
          ? ribbonMenuText.textToColumnsStatus.replace('{count}', String(max))
          : ribbonMenuText.textToColumnsNoDelimited;
    focusSheet();
  };

  const splitTextToColumnsCustom = async (): Promise<void> => {
    const delimiter = await showPrompt({
      title: ribbonMenuText.textToColumnsDialogTitle,
      label: ribbonMenuText.textToColumnsDialogDelimiters,
      initial: ',',
    });
    if (delimiter === null) return;
    splitTextToColumns(delimiter || ',');
  };

  return {
    openFilterForSelection,
    applyAdvancedFilterAction,
    sortSelection,
    customSortSelection,
    removeDuplicateRows,
    splitTextToColumns,
    splitTextToColumnsCustom,
  };
};
