import {
  type AutoSumFormulaName,
  addSheet,
  autoSum,
  boundingRange,
  type CellBorderStyle,
  type ConditionalRule,
  cellValueIsFormulaError,
  circleInvalidValidationDataInSheet,
  clearComment,
  clearFormat,
  clearHyperlink,
  clearPrintArea,
  clearPrintTitles,
  clearSheetBackgroundImage,
  clearTraceArrowsByKind,
  clearValidationCircles,
  clearValidationInRangeWithEngine,
  clearVisualFormat,
  clearWatchedCells,
  colGroupRangeAt,
  colLetter,
  collapseColGroup,
  collapseRowGroup,
  createRibbonChartFromSelection,
  deleteCells,
  deleteCols,
  deleteRows,
  expandColGroup,
  expandRowGroup,
  fillRange,
  fillSeriesSourceRange,
  filterBySelectedCellValue,
  findMatchingCells,
  getPageSetup,
  groupCols,
  groupRows,
  ignoreCellError,
  inferAutoFilterRange,
  inferFillSeriesDirection,
  inferRecommendedChartKind,
  insertCells,
  insertCols,
  insertManualPageBreak,
  insertRows,
  isCellWritable,
  isWorkbookStructureProtected,
  listComments,
  mutators,
  type Range,
  type RibbonFillDirection,
  type RibbonFillSeriesMode,
  recordCommentChange,
  recordConditionalRulesChange,
  recordFilterChange,
  recordFormatChange,
  recordIgnoredErrorsChange,
  recordPageSetupChange,
  recordValidationCirclesChange,
  recordWatchesChange,
  removeManualPageBreak,
  removeSheet,
  resetManualPageBreaks,
  rowGroupRangeAt,
  type SessionChartKind,
  type SpreadsheetInstance,
  selectNextFormulaError,
  setPrintArea,
  setPrintTitleCols,
  setPrintTitleRows,
  setRotation,
  setSheetBackgroundImage,
  showFillSeriesDialog,
  type ToolbarMenuText,
  type ToolbarText,
  ungroupCols,
  ungroupRows,
  unwatchCell,
  warnProtected,
  watchRange,
} from '@libraz/formulon-cell';
import { showChoiceDialog, showMessage, showNumberPrompt, showPrompt } from './dialogs.js';
import type { SessionShapeKind } from './illustrations.js';

type PageBreakAction =
  | 'insert-auto'
  | 'insert-row'
  | 'insert-col'
  | 'remove-row'
  | 'remove-col'
  | 'reset-all';

export type PrintTitlesAction = 'rows' | 'cols' | 'clear';

export interface RibbonActionsCtx {
  getInst: () => SpreadsheetInstance | null;
  ribbonLang: 'ja' | 'en';
  ribbonText: ToolbarText;
  ribbonMenuText: ToolbarMenuText;
  sheetEl: HTMLElement;
  getStatusMetric: () => HTMLElement | null;
  refreshWorkbookCells: () => void;
  focusSheet: () => void;
  projectFormatToolbar: () => void;
  applyRibbonFormat: (
    fn: (
      state: ReturnType<SpreadsheetInstance['store']['getState']>,
      store: SpreadsheetInstance['store'],
    ) => void,
  ) => void;
  renderSheetTabs: () => void;
  switchSheet: (idx: number) => void;
  selectedRowCount: () => number;
  selectedColCount: () => number;
  clearHyperlinksInSelection: (mode?: 'clear' | 'remove') => void;
  addSessionIllustration: (
    kind: 'image' | 'shape' | 'screenshot',
    detail: Record<string, unknown>,
  ) => void;
  runFormulaErrorChecking: () => void;
}

export interface RibbonActionsApi {
  selectMatchingAddresses: (
    matches: readonly { sheet: number; row: number; col: number }[],
  ) => void;
  applyFindSelectAction: (action: string) => void;
  applyAutoSumFormula: (fn?: AutoSumFormulaName) => void;
  cfSelectionRange: () => {
    sheet: number;
    r0: number;
    c0: number;
    r1: number;
    c1: number;
  } | null;
  normalizedSelectionRange: () => {
    sheet: number;
    r0: number;
    c0: number;
    r1: number;
    c1: number;
  } | null;
  clearSelectionContents: () => void;
  applyFillDirection: (direction: 'down' | 'right' | 'up' | 'left') => void;
  runFillSeries: (range: Range, direction: RibbonFillDirection, mode: RibbonFillSeriesMode) => void;
  applyFillSeries: (modeOverride?: RibbonFillSeriesMode) => Promise<void>;
  applyClearAction: (action: string) => void;
  promptDimension: (
    title: string,
    label: string,
    initial: number,
    max: number,
  ) => Promise<number | null>;
  applyCellInsertAction: (action: string) => Promise<void>;
  applyCellDeleteAction: (action: string) => Promise<void>;
  sheetTabColorByAction: (action: string) => string | null | undefined;
  applyTextOrientationAction: (action: string) => void;
  addConditionalRuleFromRibbon: (rule: ConditionalRule) => void;
  promptCfNumber: (
    title: string,
    initial?: number,
    options?: { min?: number; max?: number; step?: number },
  ) => Promise<number | null>;
  promptCfText: (title: string, label: string, initial?: string) => Promise<string | null>;
  selectionToA1Range: () => string | null;
  applyPrintAreaAction: (action: 'set' | 'clear') => void;
  applyPageBreakAction: (action?: string) => void;
  applySheetBackgroundAction: (action?: 'set' | 'clear') => Promise<void>;
  applyPrintTitlesAction: (action: PrintTitlesAction) => void;
  selectionOutlineAxis: () => 'row' | 'col';
  selectionDetailOutlineAxis: () => 'row' | 'col';
  selectedRowOutlineRange: () => { r0: number; r1: number } | null;
  selectedColOutlineRange: () => { c0: number; c1: number } | null;
  applyOutlineAction: (action: 'group' | 'ungroup' | 'show-detail' | 'hide-detail') => void;
  selectReviewComment: (direction: 1 | -1) => void;
  deleteActiveReviewComment: () => void;
  applyReviewCommentAction: (action: string) => void;
  insertSymbolIntoActiveCell: (symbol: string) => void;
  insertCustomSymbolIntoActiveCell: () => Promise<void>;
  applyDataValidationAction: (action: string) => void;
  applyFormulaAuditAction: (action: string) => void;
  applyWatchAction: (action: string) => void;
  insertPictureFromRibbon: (action: string) => Promise<void>;
  insertShapeFromRibbon: (shape: SessionShapeKind) => void;
  insertScreenshotFromRibbon: () => void;
  createChartFromSelection: (kind?: SessionChartKind) => void;
  recommendedChartKind: () => SessionChartKind;
  chartLabel: (kind: SessionChartKind) => string;
  createRecommendedChartFromSelection: () => Promise<void>;
  chartKindFromAction: (action: string) => SessionChartKind;
}

const DEFAULT_SHEET_BACKGROUND =
  'linear-gradient(135deg, rgba(33,115,70,0.12), rgba(0,120,212,0.08)), repeating-linear-gradient(45deg, rgba(255,255,255,0.36) 0 12px, rgba(255,255,255,0.12) 12px 24px)';

export const createRibbonActions = (ctx: RibbonActionsCtx): RibbonActionsApi => {
  const {
    getInst,
    ribbonLang,
    ribbonText,
    ribbonMenuText,
    sheetEl,
    getStatusMetric,
    refreshWorkbookCells,
    focusSheet,
    projectFormatToolbar,
    applyRibbonFormat,
    renderSheetTabs,
    switchSheet,
    selectedRowCount,
    selectedColCount,
    clearHyperlinksInSelection,
    addSessionIllustration,
    runFormulaErrorChecking,
  } = ctx;

  const selectMatchingAddresses = (
    matches: readonly { sheet: number; row: number; col: number }[],
  ): void => {
    const i = getInst();
    if (!i) return;
    if (matches.length === 0) {
      void showMessage({
        title: ribbonMenuText.findSelect,
        message: ribbonMenuText.findNoMatches,
      });
      return;
    }
    const range = boundingRange(matches);
    const active = matches[0];
    if (!active) return;
    i.store.setState((state) => ({
      ...state,
      selection: {
        active,
        anchor: active,
        range,
        extraRanges: [],
      },
    }));
    focusSheet();
  };

  const applyFindSelectAction = (action: string): void => {
    const i = getInst();
    if (!i) return;
    if (action === 'find') {
      i.openFindReplace('find');
      return;
    }
    if (action === 'replace') {
      i.openFindReplace('replace');
      return;
    }
    if (action === 'go-to') {
      i.openGoTo();
      return;
    }
    if (action === 'go-to-special') {
      i.openGoToSpecial();
      return;
    }
    if (action === 'conditional-format') {
      selectMatchingAddresses(
        findMatchingCells(i.workbook, i.store, 'sheet', 'conditional-format'),
      );
      return;
    }
    if (action === 'formulas' || action === 'constants' || action === 'data-validation') {
      selectMatchingAddresses(findMatchingCells(i.workbook, i.store, 'sheet', action));
      return;
    }
    if (action === 'comments') {
      selectMatchingAddresses(listComments(i.store.getState()).map((entry) => entry.addr));
    }
  };

  const applyAutoSumFormula = (fn: AutoSumFormulaName = 'SUM'): void => {
    const i = getInst();
    if (!i) return;
    if (fn === 'MORE') {
      i.openFunctionArguments();
      return;
    }
    i.history.begin();
    let result: ReturnType<typeof autoSum> = null;
    try {
      result = autoSum(i.store.getState(), i.workbook, fn);
    } finally {
      i.history.end();
    }
    if (result) {
      refreshWorkbookCells();
      mutators.setActive(i.store, result.addr);
    }
    focusSheet();
  };

  const cfSelectionRange = () => {
    const i = getInst();
    if (!i) return null;
    const r = i.store.getState().selection.range;
    return {
      sheet: r.sheet,
      r0: Math.min(r.r0, r.r1),
      c0: Math.min(r.c0, r.c1),
      r1: Math.max(r.r0, r.r1),
      c1: Math.max(r.c0, r.c1),
    };
  };

  const normalizedSelectionRange = () => {
    const i = getInst();
    if (!i) return null;
    const r = i.store.getState().selection.range;
    return {
      sheet: r.sheet,
      r0: Math.min(r.r0, r.r1),
      c0: Math.min(r.c0, r.c1),
      r1: Math.max(r.r0, r.r1),
      c1: Math.max(r.c0, r.c1),
    };
  };

  const clearSelectionContents = (): void => {
    const i = getInst();
    const range = normalizedSelectionRange();
    if (!i || !range) return;
    i.history.begin();
    try {
      for (let row = range.r0; row <= range.r1; row += 1) {
        for (let col = range.c0; col <= range.c1; col += 1) {
          i.workbook.setBlank({ sheet: range.sheet, row, col });
        }
      }
    } finally {
      i.history.end();
    }
    refreshWorkbookCells();
    sheetEl.focus();
  };

  const applyFillDirection = (direction: 'down' | 'right' | 'up' | 'left'): void => {
    const i = getInst();
    const range = normalizedSelectionRange();
    if (!i || !range) return;
    let src = range;
    if (direction === 'down') src = { ...range, r1: range.r0 };
    else if (direction === 'up') src = { ...range, r0: range.r1 };
    else if (direction === 'right') src = { ...range, c1: range.c0 };
    else src = { ...range, c0: range.c1 };
    if (src.r0 === range.r0 && src.r1 === range.r1 && src.c0 === range.c0 && src.c1 === range.c1) {
      return;
    }
    i.history.begin();
    try {
      recordFormatChange(i.history, i.store, () => {
        fillRange(i.store.getState(), i.workbook, src, range, {
          formatting: 'with',
          store: i.store,
        });
      });
    } finally {
      i.history.end();
    }
    refreshWorkbookCells();
    sheetEl.focus();
  };

  const runFillSeries = (
    range: Range,
    direction: RibbonFillDirection,
    mode: RibbonFillSeriesMode,
  ): void => {
    const i = getInst();
    if (!i) return;
    const src = fillSeriesSourceRange(range, direction);
    if (src.r0 === range.r0 && src.r1 === range.r1 && src.c0 === range.c0 && src.c1 === range.c1) {
      return;
    }
    const dateUnit =
      mode === 'days' || mode === 'weekdays' || mode === 'months' || mode === 'years'
        ? mode
        : undefined;
    i.history.begin();
    try {
      recordFormatChange(i.history, i.store, () => {
        fillRange(i.store.getState(), i.workbook, src, range, {
          copyOnly: mode === 'copy',
          dateUnit,
          formatting: 'with',
          store: i.store,
        });
      });
    } finally {
      i.history.end();
    }
    refreshWorkbookCells();
    sheetEl.focus();
  };

  const applyFillSeries = async (modeOverride?: RibbonFillSeriesMode): Promise<void> => {
    const i = getInst();
    const range = normalizedSelectionRange();
    if (!i || !range) return;
    if (modeOverride) {
      runFillSeries(range, inferFillSeriesDirection(range), modeOverride);
      return;
    }
    const choice = await showFillSeriesDialog(range, ribbonLang);
    if (!choice) return;
    runFillSeries(range, choice.direction, choice.mode);
  };

  const applyClearAction = (action: string): void => {
    const i = getInst();
    const range = normalizedSelectionRange();
    if (!i || !range) return;
    if (action === 'contents') {
      clearSelectionContents();
      return;
    }
    if (action === 'formats') {
      applyRibbonFormat((s, store) => clearVisualFormat(s, store));
      return;
    }
    if (action === 'conditional') {
      recordConditionalRulesChange(i.history, i.store, () => {
        mutators.clearConditionalRulesInRange(i.store, range);
      });
      refreshWorkbookCells();
      sheetEl.focus();
      return;
    }
    if (action === 'comments') {
      const addrs: Array<{ sheet: number; row: number; col: number }> = [];
      for (let row = range.r0; row <= range.r1; row += 1) {
        for (let col = range.c0; col <= range.c1; col += 1) {
          addrs.push({ sheet: range.sheet, row, col });
        }
      }
      recordCommentChange(i.history, i.store, i.workbook, addrs, () => {
        for (let row = range.r0; row <= range.r1; row += 1) {
          for (let col = range.c0; col <= range.c1; col += 1) {
            clearComment(i.store, { sheet: range.sheet, row, col }, i.workbook);
          }
        }
      });
      refreshWorkbookCells();
      sheetEl.focus();
      return;
    }
    if (action === 'hyperlinks') {
      clearHyperlinksInSelection('clear');
      return;
    }
    if (action === 'remove-hyperlinks') {
      clearHyperlinksInSelection('remove');
      return;
    }
    if (action === 'all') {
      i.history.begin();
      try {
        recordFormatChange(i.history, i.store, () => {
          clearFormat(i.store.getState(), i.store);
        });
        recordConditionalRulesChange(i.history, i.store, () => {
          mutators.clearConditionalRulesInRange(i.store, range);
        });
        for (let row = range.r0; row <= range.r1; row += 1) {
          for (let col = range.c0; col <= range.c1; col += 1) {
            i.workbook.setBlank({ sheet: range.sheet, row, col });
          }
        }
      } finally {
        i.history.end();
      }
      refreshWorkbookCells();
      sheetEl.focus();
    }
  };

  const promptDimension = async (
    title: string,
    label: string,
    initial: number,
    max: number,
  ): Promise<number | null> => {
    return showNumberPrompt({
      title,
      label,
      initial,
      min: 1,
      max,
      step: 1,
      okLabel: 'OK',
      cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
      invalidMessage:
        ribbonLang === 'ja'
          ? `1 から ${max} までの数値を入力してください`
          : `Enter a number from 1 to ${max}.`,
    });
  };

  const applyCellInsertAction = async (action: string): Promise<void> => {
    const i = getInst();
    const range = normalizedSelectionRange();
    if (!i || !range) return;
    if (action === 'rows') {
      insertRows(i.store, i.workbook, i.history, range.r0, selectedRowCount());
    } else if (action === 'cols') {
      insertCols(i.store, i.workbook, i.history, range.c0, selectedColCount());
    } else if (action === 'sheet') {
      const added = addSheet(i.store, i.workbook, i.history);
      if (added >= 0) {
        renderSheetTabs();
        switchSheet(added);
      } else {
        const statusMetric = getStatusMetric();
        if (statusMetric && isWorkbookStructureProtected(i.store.getState())) {
          statusMetric.textContent = ribbonMenuText.workbookStructureProtectedBlocked;
        }
      }
    } else if (action === 'shift-down' || action === 'shift-right') {
      insertCells(
        i.store,
        i.workbook,
        i.history,
        range,
        action === 'shift-down' ? 'down' : 'right',
      );
    } else {
      const choice = await showChoiceDialog<'down' | 'right'>({
        title: ribbonLang === 'ja' ? 'セルを挿入' : 'Insert Cells',
        label: ribbonLang === 'ja' ? '挿入後のセルの移動方向' : 'Shift cells',
        initial: 'down',
        okLabel: 'OK',
        cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
        options: [
          { value: 'down', label: ribbonLang === 'ja' ? '下方向にシフト' : 'Shift cells down' },
          { value: 'right', label: ribbonLang === 'ja' ? '右方向にシフト' : 'Shift cells right' },
        ],
      });
      if (choice === 'down' || choice === 'right') {
        insertCells(i.store, i.workbook, i.history, range, choice);
      }
    }
    refreshWorkbookCells();
    sheetEl.focus();
  };

  const applyCellDeleteAction = async (action: string): Promise<void> => {
    const i = getInst();
    const range = normalizedSelectionRange();
    if (!i || !range) return;
    if (action === 'rows') {
      deleteRows(i.store, i.workbook, i.history, range.r0, selectedRowCount());
    } else if (action === 'cols') {
      deleteCols(i.store, i.workbook, i.history, range.c0, selectedColCount());
    } else if (action === 'sheet') {
      const before = i.store.getState().data.sheetIndex;
      if (removeSheet(i.store, i.workbook, before)) {
        renderSheetTabs();
        switchSheet(i.store.getState().data.sheetIndex);
      }
    } else if (action === 'shift-up' || action === 'shift-left') {
      deleteCells(i.store, i.workbook, i.history, range, action === 'shift-up' ? 'up' : 'left');
    } else {
      const choice = await showChoiceDialog<'up' | 'left'>({
        title: ribbonLang === 'ja' ? 'セルを削除' : 'Delete Cells',
        label: ribbonLang === 'ja' ? '削除後のセルの移動方向' : 'Shift cells',
        initial: 'up',
        okLabel: 'OK',
        cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
        options: [
          { value: 'up', label: ribbonLang === 'ja' ? '上方向にシフト' : 'Shift cells up' },
          { value: 'left', label: ribbonLang === 'ja' ? '左方向にシフト' : 'Shift cells left' },
        ],
      });
      if (choice === 'up' || choice === 'left') {
        deleteCells(i.store, i.workbook, i.history, range, choice);
      }
    }
    refreshWorkbookCells();
    sheetEl.focus();
  };

  const sheetTabColorByAction = (action: string): string | null | undefined => {
    const colors: Record<string, string | null> = {
      'tab-color-none': null,
      'tab-color-red': '#c00000',
      'tab-color-orange': '#ed7d31',
      'tab-color-yellow': '#ffc000',
      'tab-color-green': '#70ad47',
      'tab-color-blue': '#4472c4',
      'tab-color-purple': '#7030a0',
      'tab-color-gray': '#808080',
    };
    return Object.hasOwn(colors, action) ? colors[action] : undefined;
  };

  const applyTextOrientationAction = (action: string): void => {
    const i = getInst();
    if (!i) return;
    if (action === 'format') {
      i.openFormatDialog();
      return;
    }
    const rotations: Record<string, number> = {
      horizontal: 0,
      ccw: 45,
      cw: -45,
      vertical: 90,
      up: 90,
      down: -90,
    };
    const rotation = rotations[action];
    if (typeof rotation !== 'number') return;
    applyRibbonFormat((s, store) => setRotation(s, store, rotation));
  };

  const addConditionalRuleFromRibbon = (rule: ConditionalRule): void => {
    const i = getInst();
    if (!i) return;
    recordConditionalRulesChange(i.history, i.store, () => {
      mutators.addConditionalRule(i.store, rule);
    });
    refreshWorkbookCells();
    sheetEl.focus();
  };

  const promptCfNumber = async (
    title: string,
    initial = 0,
    options: { min?: number; max?: number; step?: number } = {},
  ): Promise<number | null> => {
    return showNumberPrompt({
      title,
      label: ribbonLang === 'ja' ? '値' : 'Value',
      initial,
      min: options.min,
      max: options.max,
      step: options.step,
      cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
      invalidMessage: ribbonLang === 'ja' ? '数値を入力してください' : 'Enter a number',
    });
  };

  const promptCfText = async (
    title: string,
    label: string,
    initial = '',
  ): Promise<string | null> => {
    const value = await showPrompt({
      title,
      label,
      initial,
      validate: (raw) =>
        raw.trim() ? null : ribbonLang === 'ja' ? '値を入力してください' : 'Enter a value',
    });
    return value === null ? null : value.trim();
  };

  const selectionToA1Range = (): string | null => {
    const i = getInst();
    if (!i) return null;
    const r = i.store.getState().selection.range;
    const start = `${colLetter(r.c0)}${r.r0 + 1}`;
    const end = `${colLetter(r.c1)}${r.r1 + 1}`;
    return start === end ? start : `${start}:${end}`;
  };

  const applyPrintAreaAction = (action: 'set' | 'clear'): void => {
    const i = getInst();
    if (!i) return;
    const sheet = i.store.getState().data.sheetIndex;
    recordPageSetupChange(i.history, i.store, () => {
      if (action === 'clear') {
        clearPrintArea(i.store, sheet);
        return;
      }
      const area = selectionToA1Range();
      if (area) setPrintArea(i.store, sheet, area);
    });
    const setup = getPageSetup(i.store.getState(), sheet);
    const message =
      action === 'clear'
        ? ribbonMenuText.printAreaStatusClear
        : ribbonMenuText.printAreaStatusSet.replace('{range}', setup.printArea ?? '');
    showMessage({
      title: ribbonText.printArea,
      message,
    });
    projectFormatToolbar();
    focusSheet();
  };

  const applyPageBreakAction = (action: string = 'insert-auto'): void => {
    const i = getInst();
    if (!i) return;
    const state = i.store.getState();
    const sheet = state.data.sheetIndex;
    const range = state.selection.range;
    recordPageSetupChange(i.history, i.store, () => {
      const pageBreakAction = action as PageBreakAction;
      if (pageBreakAction === 'reset-all') {
        resetManualPageBreaks(i.store, sheet);
        return;
      }
      if (pageBreakAction === 'remove-row') {
        removeManualPageBreak(i.store, sheet, 'row', range.r0);
        return;
      }
      if (pageBreakAction === 'remove-col') {
        removeManualPageBreak(i.store, sheet, 'col', range.c0);
        return;
      }
      if (pageBreakAction === 'insert-col') {
        if (range.c0 > 0) insertManualPageBreak(i.store, sheet, 'col', range.c0);
        return;
      }
      if (pageBreakAction === 'insert-row') {
        if (range.r0 > 0) insertManualPageBreak(i.store, sheet, 'row', range.r0);
        return;
      }
      if (range.r0 > 0) insertManualPageBreak(i.store, sheet, 'row', range.r0);
      else if (range.c0 > 0) insertManualPageBreak(i.store, sheet, 'col', range.c0);
    });
    projectFormatToolbar();
    focusSheet();
  };

  const applySheetBackgroundAction = async (action: 'set' | 'clear' = 'set'): Promise<void> => {
    const i = getInst();
    if (!i) return;
    const sheet = i.store.getState().data.sheetIndex;
    if (action === 'clear') {
      clearSheetBackgroundImage(i.store, sheet, i.history);
      projectFormatToolbar();
      focusSheet();
      return;
    }
    const current = i.store.getState().ui.sheetBackgroundImages.get(sheet);
    const value = await showPrompt({
      title: ribbonMenuText.sheetBackgroundSet,
      label: ribbonMenuText.sheetBackgroundPrompt,
      initial: current ?? DEFAULT_SHEET_BACKGROUND,
      okLabel: 'OK',
      cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
      validate: (raw) =>
        raw.trim()
          ? null
          : ribbonLang === 'ja'
            ? '背景画像のURLを入力してください。'
            : 'Enter a background image URL.',
    });
    if (value === null) {
      focusSheet();
      return;
    }
    setSheetBackgroundImage(i.store, sheet, value.trim(), i.history);
    projectFormatToolbar();
    focusSheet();
  };

  const applyPrintTitlesAction = (action: PrintTitlesAction): void => {
    const i = getInst();
    if (!i) return;
    const state = i.store.getState();
    const sheet = state.data.sheetIndex;
    const range = state.selection.range;
    if (action === 'clear') {
      clearPrintTitles(i.store, sheet, i.history);
    } else if (action === 'rows') {
      setPrintTitleRows(i.store, sheet, `${range.r0 + 1}:${range.r1 + 1}`, i.history);
    } else {
      setPrintTitleCols(i.store, sheet, `${colLetter(range.c0)}:${colLetter(range.c1)}`, i.history);
    }
    projectFormatToolbar();
    focusSheet();
  };

  const selectionOutlineAxis = (): 'row' | 'col' => {
    const i = getInst();
    if (!i) return 'row';
    const r = i.store.getState().selection.range;
    const rowSpan = r.r1 - r.r0;
    const colSpan = r.c1 - r.c0;
    return rowSpan >= colSpan ? 'row' : 'col';
  };

  const selectionDetailOutlineAxis = (): 'row' | 'col' => {
    const i = getInst();
    if (!i) return 'row';
    const state = i.store.getState();
    const activeRowLevel = state.layout.outlineRows.get(state.selection.active.row) ?? 0;
    const activeColLevel = state.layout.outlineCols.get(state.selection.active.col) ?? 0;
    if (activeRowLevel > 0 && activeColLevel === 0) return 'row';
    if (activeColLevel > 0 && activeRowLevel === 0) return 'col';
    return selectionOutlineAxis();
  };

  const selectedRowOutlineRange = (): { r0: number; r1: number } | null => {
    const i = getInst();
    if (!i) return null;
    const state = i.store.getState();
    const range = state.selection.range;
    let bestRow = state.selection.active.row;
    let bestLevel = state.layout.outlineRows.get(bestRow) ?? 0;
    for (let row = range.r0; row <= range.r1; row += 1) {
      const level = state.layout.outlineRows.get(row) ?? 0;
      if (level > bestLevel) {
        bestRow = row;
        bestLevel = level;
      }
    }
    return bestLevel > 0 ? rowGroupRangeAt(state.layout, bestRow, bestLevel) : null;
  };

  const selectedColOutlineRange = (): { c0: number; c1: number } | null => {
    const i = getInst();
    if (!i) return null;
    const state = i.store.getState();
    const range = state.selection.range;
    let bestCol = state.selection.active.col;
    let bestLevel = state.layout.outlineCols.get(bestCol) ?? 0;
    for (let col = range.c0; col <= range.c1; col += 1) {
      const level = state.layout.outlineCols.get(col) ?? 0;
      if (level > bestLevel) {
        bestCol = col;
        bestLevel = level;
      }
    }
    return bestLevel > 0 ? colGroupRangeAt(state.layout, bestCol, bestLevel) : null;
  };

  const applyOutlineAction = (
    action: 'group' | 'ungroup' | 'show-detail' | 'hide-detail',
  ): void => {
    const i = getInst();
    if (!i) return;
    const range = i.store.getState().selection.range;
    const axis =
      action === 'show-detail' || action === 'hide-detail'
        ? selectionDetailOutlineAxis()
        : selectionOutlineAxis();
    if (axis === 'row') {
      if (action === 'group') groupRows(i.store, i.history, range.r0, range.r1, i.workbook);
      else if (action === 'ungroup')
        ungroupRows(i.store, i.history, range.r0, range.r1, i.workbook);
      else {
        const group = selectedRowOutlineRange();
        if (!group) return;
        if (action === 'show-detail')
          expandRowGroup(i.store, i.history, group.r0, group.r1, i.workbook);
        else collapseRowGroup(i.store, i.history, group.r0, group.r1, i.workbook);
      }
    } else {
      if (action === 'group') groupCols(i.store, i.history, range.c0, range.c1, i.workbook);
      else if (action === 'ungroup')
        ungroupCols(i.store, i.history, range.c0, range.c1, i.workbook);
      else {
        const group = selectedColOutlineRange();
        if (!group) return;
        if (action === 'show-detail')
          expandColGroup(i.store, i.history, group.c0, group.c1, i.workbook);
        else collapseColGroup(i.store, i.history, group.c0, group.c1, i.workbook);
      }
    }
    refreshWorkbookCells();
    projectFormatToolbar();
    focusSheet();
  };

  const selectReviewComment = (direction: 1 | -1): void => {
    const i = getInst();
    if (!i) return;
    const comments = listComments(i.store.getState());
    const ja = ribbonLang === 'ja';
    if (comments.length === 0) {
      void showMessage({
        title: ja ? 'コメント' : 'Comments',
        message: ja ? 'コメントまたはメモが見つかりません。' : 'No comments or notes were found.',
      });
      return;
    }
    const active = i.store.getState().selection.active;
    const current = comments.findIndex(
      (entry) => entry.addr.row === active.row && entry.addr.col === active.col,
    );
    const nextIndex =
      current >= 0
        ? (current + direction + comments.length) % comments.length
        : direction > 0
          ? 0
          : comments.length - 1;
    const next = comments[nextIndex]?.addr;
    if (!next) return;
    mutators.setActive(i.store, next);
    i.openCommentDialog();
  };

  const deleteActiveReviewComment = (): void => {
    const i = getInst();
    if (!i) return;
    const addr = i.store.getState().selection.active;
    recordCommentChange(i.history, i.store, i.workbook, [addr], () => {
      clearComment(i.store, addr, i.workbook);
    });
    projectFormatToolbar();
    focusSheet();
  };

  const applyReviewCommentAction = (action: string): void => {
    const i = getInst();
    if (!i) return;
    if (action === 'delete-active') {
      deleteActiveReviewComment();
      return;
    }
    if (action !== 'delete-all') return;
    const comments = listComments(i.store.getState());
    if (comments.length === 0) {
      void showMessage({
        title: ribbonText.comments,
        message: ribbonMenuText.commentNone,
      });
      return;
    }
    recordCommentChange(
      i.history,
      i.store,
      i.workbook,
      comments.map((entry) => entry.addr),
      () => {
        for (const entry of comments) clearComment(i.store, entry.addr, i.workbook);
      },
    );
    const statusMetric = getStatusMetric();
    if (statusMetric) {
      statusMetric.textContent = ribbonMenuText.commentsDeleted.replace(
        '{count}',
        String(comments.length),
      );
    }
    projectFormatToolbar();
    focusSheet();
  };

  const insertSymbolIntoActiveCell = (symbol: string): void => {
    const i = getInst();
    if (!i) return;
    const addr = i.store.getState().selection.active;
    if (i.workbook.cellFormula(addr)) {
      void showMessage({
        title: ribbonLang === 'ja' ? '記号' : 'Symbol',
        message:
          ribbonLang === 'ja'
            ? '数式セルには記号を直接挿入できません。'
            : 'Symbols cannot be inserted directly into a formula cell.',
      });
      return;
    }
    if (!isCellWritable(i.store.getState(), addr)) {
      warnProtected(addr);
      void showMessage({
        title: ribbonLang === 'ja' ? '記号' : 'Symbol',
        message:
          ribbonLang === 'ja'
            ? '保護されたセルには記号を挿入できません。'
            : 'Symbols cannot be inserted into a protected cell.',
      });
      return;
    }
    const value = i.workbook.getValue(addr);
    const current = value.kind === 'text' ? value.value : '';
    i.history.begin();
    try {
      i.workbook.setText(addr, `${current}${symbol}`);
    } finally {
      i.history.end();
    }
    refreshWorkbookCells();
    focusSheet();
  };

  const insertCustomSymbolIntoActiveCell = async (): Promise<void> => {
    const value = await showPrompt({
      title: ribbonMenuText.symbolMore,
      label: ribbonMenuText.symbolPrompt,
      okLabel: 'OK',
      cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
      validate: (raw) => (raw.trim() ? null : ribbonMenuText.symbolInvalid),
    });
    if (value === null) {
      focusSheet();
      return;
    }
    insertSymbolIntoActiveCell(value.trim());
  };

  const applyDataValidationAction = (action: string): void => {
    const i = getInst();
    if (!i) return;
    const range = i.store.getState().selection.range;
    if (action === 'settings') {
      i.openDataValidationDialog();
      return;
    }
    if (action === 'circle-invalid') {
      const count = recordValidationCirclesChange(i.history, i.store, () =>
        circleInvalidValidationDataInSheet(i.store, range.sheet),
      );
      const statusMetric = getStatusMetric();
      if (statusMetric)
        statusMetric.textContent = `${ribbonMenuText.validationCircleInvalid} · ${count}`;
      projectFormatToolbar();
      focusSheet();
      return;
    }
    if (action === 'clear-circles') {
      recordValidationCirclesChange(i.history, i.store, () => clearValidationCircles(i.store));
      projectFormatToolbar();
      focusSheet();
      return;
    }
    if (action === 'clear-rules') {
      clearValidationInRangeWithEngine(i.store, i.history, i.workbook, range);
      refreshWorkbookCells();
      projectFormatToolbar();
      focusSheet();
    }
  };

  const applyFormulaAuditAction = (action: string): void => {
    const i = getInst();
    if (!i) return;
    const showTraceEmpty = (message: string): void => {
      void showMessage({ title: ribbonText.formulaAuditing, message });
    };
    if (action === 'clear-all') {
      i.clearTraces();
    } else if (action === 'clear-precedents') {
      clearTraceArrowsByKind(i.store, 'precedent', i.history);
      refreshWorkbookCells();
    } else if (action === 'clear-dependents') {
      clearTraceArrowsByKind(i.store, 'dependent', i.history);
      refreshWorkbookCells();
    } else if (action === 'error-checking') {
      runFormulaErrorChecking();
      return;
    } else if (action === 'trace-error') {
      const found = selectNextFormulaError(i.store);
      if (found) {
        if (i.tracePrecedents() === 0) showTraceEmpty(ribbonMenuText.traceNoPrecedents);
      } else {
        runFormulaErrorChecking();
      }
      return;
    } else if (action === 'ignore-error') {
      const active = i.store.getState().selection.active;
      const key = `${active.sheet}:${active.row}:${active.col}`;
      const cell = i.store.getState().data.cells.get(key);
      if (cell?.formula && cellValueIsFormulaError(cell.value)) {
        recordIgnoredErrorsChange(i.history, i.store, () => {
          ignoreCellError(i.store, active);
        });
        const statusMetric = getStatusMetric();
        if (statusMetric) statusMetric.textContent = ribbonMenuText.ignoreError;
        projectFormatToolbar();
        focusSheet();
        return;
      }
      runFormulaErrorChecking();
      return;
    }
    projectFormatToolbar();
    focusSheet();
  };

  const applyWatchAction = (action: string): void => {
    const i = getInst();
    if (!i) return;
    const state = i.store.getState();
    if (action === 'open') {
      i.openWatchWindow();
      return;
    }
    if (action === 'add') {
      recordWatchesChange(i.history, i.store, () => {
        watchRange(i.store, state.selection.range);
      });
      i.openWatchWindow();
      return;
    }
    if (action === 'delete') {
      recordWatchesChange(i.history, i.store, () => {
        unwatchCell(i.store, state.selection.active);
      });
      i.openWatchWindow();
      return;
    }
    if (action === 'delete-all') {
      recordWatchesChange(i.history, i.store, () => {
        clearWatchedCells(i.store);
      });
      i.openWatchWindow();
    }
  };

  const insertPictureFromRibbon = async (action: string): Promise<void> => {
    const title =
      action === 'device' ? ribbonMenuText.pictureThisDevice : ribbonMenuText.pictureOnline;
    const url = await showPrompt({
      title,
      label: ribbonMenuText.pictureUrlPrompt,
      placeholder: 'https://...',
      okLabel: 'OK',
      cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
      validate: (value) =>
        value.trim() ? null : ribbonLang === 'ja' ? 'URLを入力してください。' : 'Enter a URL.',
    });
    if (!url) {
      focusSheet();
      return;
    }
    addSessionIllustration('image', { url: url.trim(), w: 220, h: 140 });
  };

  const insertShapeFromRibbon = (shape: SessionShapeKind): void => {
    addSessionIllustration('shape', { shape });
  };

  const insertScreenshotFromRibbon = (): void => {
    addSessionIllustration('screenshot', { w: 230, h: 150 });
  };

  const createChartFromSelection = (kind: SessionChartKind = 'column'): void => {
    const i = getInst();
    if (!i) return;
    createRibbonChartFromSelection({
      store: i.store,
      history: i.history,
      range: i.store.getState().selection.range,
      action: kind,
    });
    focusSheet();
  };

  const recommendedChartKind = (): SessionChartKind => {
    const r = getInst()?.store.getState().selection.range;
    return r ? inferRecommendedChartKind(r) : 'column';
  };

  const chartLabel = (kind: SessionChartKind): string => {
    const t = ribbonMenuText;
    if (kind === 'bar') return t.chartBar;
    if (kind === 'line') return t.chartLine;
    if (kind === 'area') return t.chartArea;
    if (kind === 'pie') return t.chartPie;
    if (kind === 'scatter') return t.chartScatter;
    return t.chartColumn;
  };

  const createRecommendedChartFromSelection = async (): Promise<void> => {
    const t = ribbonMenuText;
    const initial = recommendedChartKind();
    const rawOptions: Array<{ value: SessionChartKind; label: string }> = [
      { value: initial, label: `${t.recommendedCharts}: ${chartLabel(initial)}` },
      { value: 'column', label: t.chartColumn },
      { value: 'bar', label: t.chartBar },
      { value: 'line', label: t.chartLine },
      { value: 'area', label: t.chartArea },
      { value: 'pie', label: t.chartPie },
      { value: 'scatter', label: t.chartScatter },
    ];
    const options = rawOptions.filter(
      (option, index, all) =>
        all.findIndex((candidate) => candidate.value === option.value) === index,
    );
    const choice = await showChoiceDialog<SessionChartKind>({
      title: t.recommendedCharts,
      label: t.chart,
      initial,
      okLabel: 'OK',
      cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
      options,
    });
    if (!choice) {
      focusSheet();
      return;
    }
    createChartFromSelection(choice);
  };

  const chartKindFromAction = (action: string): SessionChartKind =>
    action === 'bar' ||
    action === 'line' ||
    action === 'area' ||
    action === 'pie' ||
    action === 'scatter'
      ? action
      : action === 'recommended'
        ? recommendedChartKind()
        : 'column';

  return {
    selectMatchingAddresses,
    applyFindSelectAction,
    applyAutoSumFormula,
    cfSelectionRange,
    normalizedSelectionRange,
    clearSelectionContents,
    applyFillDirection,
    runFillSeries,
    applyFillSeries,
    applyClearAction,
    promptDimension,
    applyCellInsertAction,
    applyCellDeleteAction,
    sheetTabColorByAction,
    applyTextOrientationAction,
    addConditionalRuleFromRibbon,
    promptCfNumber,
    promptCfText,
    selectionToA1Range,
    applyPrintAreaAction,
    applyPageBreakAction,
    applySheetBackgroundAction,
    applyPrintTitlesAction,
    selectionOutlineAxis,
    selectionDetailOutlineAxis,
    selectedRowOutlineRange,
    selectedColOutlineRange,
    applyOutlineAction,
    selectReviewComment,
    deleteActiveReviewComment,
    applyReviewCommentAction,
    insertSymbolIntoActiveCell,
    insertCustomSymbolIntoActiveCell,
    applyDataValidationAction,
    applyFormulaAuditAction,
    applyWatchAction,
    insertPictureFromRibbon,
    insertShapeFromRibbon,
    insertScreenshotFromRibbon,
    createChartFromSelection,
    recommendedChartKind,
    chartLabel,
    createRecommendedChartFromSelection,
    chartKindFromAction,
  };
};
