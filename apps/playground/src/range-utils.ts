// A1 / range helpers + workbook cell sync + ribbon zoom dialog.
// Extracted from main.ts. The factory pattern lets the host pass in live
// references (instance accessor, dialog helpers, zoom refresh, locale) without
// coupling this module to global state.

import {
  colLetter,
  inferAutoFilterRange,
  type Range,
  type SpreadsheetInstance,
  setSheetZoom,
} from '@libraz/formulon-cell';

import { showNumberPrompt } from './dialogs.js';

export interface RangeUtilsCtx {
  getInst: () => SpreadsheetInstance | null;
  ribbonLang: 'ja' | 'en';
  refreshZoom: () => void;
  focusSheet: () => void;
}

export interface RangeUtilsApi {
  selectedRowCount: () => number;
  selectedColCount: () => number;
  sortTargetRange: (state: ReturnType<SpreadsheetInstance['store']['getState']>) => Range;
  sortCellDisplayText: (
    state: ReturnType<SpreadsheetInstance['store']['getState']>,
    row: number,
    col: number,
  ) => string;
  colFromLetters: (letters: string) => number;
  parseA1Range: (raw: string, sheet: number) => Range | null;
  rangeRef: (range: Range) => string;
  syncStoreCellsToWorkbook: (
    i: SpreadsheetInstance,
    sheet: number,
    row: number,
    col: number,
    height: number,
    width: number,
  ) => void;
  showZoomDialogFromRibbon: () => Promise<void>;
}

export const createRangeUtils = (ctx: RangeUtilsCtx): RangeUtilsApi => {
  const { getInst, ribbonLang, refreshZoom, focusSheet } = ctx;

  const selectedRowCount = (): number => {
    const inst = getInst();
    if (!inst) return 1;
    const r = inst.store.getState().selection.range;
    return Math.max(1, r.r1 - r.r0 + 1);
  };

  const selectedColCount = (): number => {
    const inst = getInst();
    if (!inst) return 1;
    const r = inst.store.getState().selection.range;
    return Math.max(1, r.c1 - r.c0 + 1);
  };

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

  const showZoomDialogFromRibbon = async (): Promise<void> => {
    const i = getInst();
    if (!i) return;
    const current = Math.round(i.store.getState().viewport.zoom * 100);
    const value = await showNumberPrompt({
      title: ribbonLang === 'ja' ? 'ズーム' : 'Zoom',
      label: ribbonLang === 'ja' ? '倍率' : 'Magnification',
      initial: current,
      min: 10,
      max: 400,
      step: 1,
      okLabel: 'OK',
      cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
    });
    if (value === null) {
      focusSheet();
      return;
    }
    setSheetZoom(i.store, Math.max(0.1, Math.min(4, value / 100)), i.workbook);
    refreshZoom();
    focusSheet();
  };

  return {
    selectedRowCount,
    selectedColCount,
    sortTargetRange,
    sortCellDisplayText,
    colFromLetters,
    parseA1Range,
    rangeRef,
    syncStoreCellsToWorkbook,
    showZoomDialogFromRibbon,
  };
};
