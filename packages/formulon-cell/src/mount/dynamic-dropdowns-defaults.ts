// Default `DynamicDropdownsCtx` factory. Hosts that don't need a fully
// custom ribbon (React / Vue / quick embed) can call this and pass the
// result to `mountToolbar({ dynamicDropdowns })` — the toolbar's auto-wire
// then handles every menu-item click through these defaults. Hosts retain
// the ability to override any single handler via the `overrides` bag.
//
// Defaults are split into three buckets:
//   1. Pure instance handlers — derived entirely from the instance and the
//      already-exported command helpers (fill / clear / autosum / etc.).
//   2. Instance dispatch — forward to existing `instance.openX` methods.
//   3. Dialog / host-glue stubs — no-op (or browser fallback) when the host
//      doesn't supply a real implementation. Overriding via `overrides` lets
//      hosts plug in their own UI without re-wiring the click delegator.

// Imports use the `@libraz/formulon-cell` self-alias instead of relative
// paths because `dynamic-dropdowns.ts` and other ribbon modules already do.
// That keeps the type identity for `SpreadsheetInstance`, `History`,
// `WorkbookHandle` aligned with the public dist build — otherwise TypeScript
// flags the merged ctx as structurally identical but nominally distinct.
import {
  type AutoSumFunction,
  addConditionalRule,
  addPrintArea,
  applyCellFormatAction,
  applyTextScriptToRange,
  autoSum,
  buildRibbonAddInReport,
  type ConditionalRule,
  clearPrintArea,
  clearPrintTitles,
  clearSheetBackgroundImage,
  clearTraceArrowsByKind,
  clearValidationInRangeWithEngine,
  clearWatchedCells,
  colLetter,
  createDefinedNamesFromSelection,
  createRibbonChartFromSelection,
  dispatchHostClipboard,
  executeRibbonClearAction,
  executeRibbonCommentAction,
  executeRibbonFilterDataAction,
  executeRibbonFindAction,
  executeRibbonFormulaAuditingAction,
  executeRibbonHyperlinkAction,
  executeRibbonPivotTableAction,
  executeRibbonProtectionAction,
  type FreezeAction,
  fillRange,
  formatA1Range,
  handleDeleteCellsAction,
  handleFreezeAction,
  handleInsertCellsAction,
  handlePasteAction,
  inferAutoFilterRange,
  inferFillSeriesDirection,
  inferSortHasHeader,
  insertDefinedNameFormula,
  insertManualPageBreak,
  listDefinedNames,
  mutators,
  type PasteAction,
  parseScriptCommand,
  type Range,
  type RibbonAddInAction,
  type RibbonFillSeriesMode,
  type RibbonPdfAction,
  type RibbonPivotTableAction,
  recordConditionalRulesChange,
  recordDefinedNamesChange,
  recordFormatChange,
  recordTablesChange,
  recordWatchesChange,
  removeDuplicates,
  removeManualPageBreak,
  resetManualPageBreaks,
  resolveRibbonPdfAction,
  type SessionChartKind,
  type SpreadsheetInstance,
  setNumFmt,
  setPrintArea,
  setPrintTitleCols,
  setPrintTitleRows,
  setRotation,
  setSheetBackgroundImage,
  setWorkbookStructureProtected,
  sortActiveColumnAuto,
  sortRangeWithHistory,
  type ThemeName,
  textToColumns,
  unwatchCell,
  watchRange,
} from '@libraz/formulon-cell';
import {
  applyCellStyleByName,
  createCellStyleFromActiveFormat,
  mergeCellStylesFromWorkbook,
} from '../commands/cell-styles.js';
import {
  applyPivotTableStyleById,
  createPivotTableStyleFromActivePivot,
  createTableStyleFromActiveTable,
  DEFAULT_TABLE_COLOR,
  formatAsTableByStyleId,
  tableOverlayAt,
  tableVariantFromOptions,
} from '../commands/format-as-table.js';
import {
  arrangeSessionIllustration,
  createRibbonImageFromSelection,
  createRibbonShapeFromSelection,
} from '../commands/session-illustration.js';
import { cellValueViolatesValidation } from '../commands/validate.js';
import { addrKey } from '../engine/address.js';
import { findPivotTableAtCell } from '../engine/passthrough-sync.js';
import { showCellStyleDialog } from '../toolbar/dialogs/cell-style.js';
import { showChoiceDialog } from '../toolbar/dialogs/choice.js';
import { pickImageFileDataUrl } from '../toolbar/dialogs/image-file.js';
import { showMessage, showNumberPrompt, showPrompt } from '../toolbar/dialogs/prompt.js';
import { showRemoveDuplicatesDialog } from '../toolbar/dialogs/remove-duplicates.js';
import { showReport } from '../toolbar/dialogs/report.js';
import { type SortDialogColumn, showSortDialog } from '../toolbar/dialogs/sort.js';
import { showSymbolDialog } from '../toolbar/dialogs/symbol.js';
import { showTableStyleDialog } from '../toolbar/dialogs/table-style.js';
import { applyConditionalMenuAction } from '../toolbar/ribbon/conditional-menu-action.js';
import type { DynamicDropdownsCtx } from '../toolbar/ribbon/dynamic-dropdowns.js';
import { fillSeriesSourceRange, showFillSeriesDialog } from '../toolbar/ribbon/fill-series.js';

/** Options accepted alongside any host overrides. Lives separately from the
 *  partial-context bag so we can extend with cross-cutting knobs (e.g. a
 *  shared `focusSheet` closure) without polluting the dropdown ctx itself. */
export interface DefaultDynamicDropdownsOptions {
  /** Per-handler overrides. Merged on top of the defaults so the host only
   *  has to supply the ones that need a real dialog. Pass a getter when the
   *  overrides aren't ready at mount time (e.g. the playground builds its
   *  ctx after `mountToolbar` returns) — the ctx will lazily resolve each
   *  handler on every dispatch. */
  overrides?: Partial<DynamicDropdownsCtx> | (() => Partial<DynamicDropdownsCtx>);
}

const noop = (): void => undefined;

const normalizedSelectionRange = (instance: SpreadsheetInstance): Range => {
  const r = instance.store.getState().selection.range;
  return {
    sheet: r.sheet,
    r0: Math.min(r.r0, r.r1),
    c0: Math.min(r.c0, r.c1),
    r1: Math.max(r.r0, r.r1),
    c1: Math.max(r.c0, r.c1),
  };
};

const buildFillDirection =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyFillDirection'] =>
  (direction) => {
    const range = normalizedSelectionRange(instance);
    let src: Range = range;
    if (direction === 'down') src = { ...range, r1: range.r0 };
    else if (direction === 'up') src = { ...range, r0: range.r1 };
    else if (direction === 'right') src = { ...range, c1: range.c0 };
    else src = { ...range, c0: range.c1 };
    if (src.r0 === range.r0 && src.r1 === range.r1 && src.c0 === range.c0 && src.c1 === range.c1) {
      return;
    }
    instance.history.begin();
    try {
      recordFormatChange(instance.history, instance.store, () => {
        fillRange(instance.store.getState(), instance.workbook, src, range, {
          formatting: 'with',
          store: instance.store,
        });
      });
    } finally {
      instance.history.end();
    }
    instance.host.focus();
  };

const buildFillSeries =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyFillSeries'] =>
  async (mode) => {
    const range = normalizedSelectionRange(instance);
    const choice = mode
      ? { direction: inferFillSeriesDirection(range), mode }
      : await showFillSeriesDialog(range, instance.i18n.locale === 'en' ? 'en' : 'ja');
    if (!choice) return;
    const src = fillSeriesSourceRange(range, choice.direction);
    if (src.r0 === range.r0 && src.r1 === range.r1 && src.c0 === range.c0 && src.c1 === range.c1) {
      return;
    }
    const dateUnit: RibbonFillSeriesMode | undefined =
      choice.mode === 'days' ||
      choice.mode === 'weekdays' ||
      choice.mode === 'months' ||
      choice.mode === 'years'
        ? choice.mode
        : undefined;
    instance.history.begin();
    try {
      recordFormatChange(instance.history, instance.store, () => {
        fillRange(instance.store.getState(), instance.workbook, src, range, {
          copyOnly: choice.mode === 'copy',
          dateUnit,
          formatting: 'with',
          store: instance.store,
        });
      });
    } finally {
      instance.history.end();
    }
    instance.host.focus();
  };

const buildClearAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyClearAction'] =>
  (action) => {
    const clearAction = action === 'remove-hyperlinks' ? 'hyperlinks' : action;
    if (
      clearAction !== 'all' &&
      clearAction !== 'formats' &&
      clearAction !== 'contents' &&
      clearAction !== 'comments' &&
      clearAction !== 'hyperlinks' &&
      clearAction !== 'conditional'
    ) {
      return;
    }
    executeRibbonClearAction({
      store: instance.store,
      workbook: instance.workbook,
      history: instance.history,
      action: clearAction,
    });
    instance.host.focus();
  };

const freezeActionFromMenu = (action: string): FreezeAction | null => {
  if (action === 'row') return 'topRow';
  if (action === 'col') return 'firstColumn';
  if (action === 'selection') return 'panes';
  if (action === 'off') return 'none';
  return null;
};

const buildFreezeAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyFreezeAction'] =>
  (action) => {
    const next = freezeActionFromMenu(action);
    if (!next) return;
    handleFreezeAction(instance, next);
    instance.host.focus();
  };

const buildTextOrientation =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyTextOrientationAction'] =>
  (action) => {
    if (action === 'format') {
      instance.openFormatDialog();
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
    recordFormatChange(instance.history, instance.store, () => {
      setRotation(instance.store.getState(), instance.store, rotation);
    });
    instance.host.focus();
  };

const buildAutoSumFormula =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyAutoSumFormula'] =>
  (fn) => {
    if (fn === 'MORE') {
      instance.openFunctionArguments();
      return;
    }
    instance.history.begin();
    let result: ReturnType<typeof autoSum> = null;
    try {
      result = autoSum(instance.store.getState(), instance.workbook, fn as AutoSumFunction);
    } finally {
      instance.history.end();
    }
    if (result) mutators.setActive(instance.store, result.addr);
    instance.host.focus();
  };

const buildWatchAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyWatchAction'] =>
  (action) => {
    const state = instance.store.getState();
    if (action === 'open') {
      instance.openWatchWindow();
      return;
    }
    if (action === 'add') {
      recordWatchesChange(instance.history, instance.store, () => {
        watchRange(instance.store, state.selection.range);
      });
      instance.openWatchWindow();
      return;
    }
    if (action === 'delete') {
      recordWatchesChange(instance.history, instance.store, () => {
        unwatchCell(instance.store, state.selection.active);
      });
      instance.openWatchWindow();
      return;
    }
    if (action === 'delete-all') {
      recordWatchesChange(instance.history, instance.store, () => {
        clearWatchedCells(instance.store);
      });
      instance.openWatchWindow();
    }
  };

const buildCalcOptionAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyCalcOptionAction'] =>
  (action) => {
    if (action === 'auto' || action === 'manual' || action === 'auto-no-table') {
      const mode = action === 'auto' ? 0 : action === 'manual' ? 1 : 2;
      instance.workbook.setCalcMode(mode as 0 | 1 | 2);
      instance.host.focus();
      return;
    }
    if (action === 'calculate-now' || action === 'calculate-sheet') {
      instance.recalc();
      instance.host.focus();
      return;
    }
    if (action === 'iterative') {
      instance.openIterativeDialog();
    }
  };

const buildFindSelectAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyFindSelectAction'] =>
  async (action) => {
    const strings = instance.i18n.strings;
    const result = executeRibbonFindAction({
      store: instance.store,
      workbook: instance.workbook,
      action: action as Parameters<typeof executeRibbonFindAction>[0]['action'],
      strings: {
        findSelect: strings.ribbon.findSelect,
        findNoMatches: strings.ribbonMenu.findNoMatches,
        commentNone: strings.ribbonMenu.commentNone,
      },
    });
    if (result.kind === 'open-find') {
      instance.openFindReplace(result.mode);
      return;
    }
    if (result.kind === 'open-go-to') {
      instance.openGoTo();
      return;
    }
    if (result.kind === 'open-go-to-special') {
      instance.openGoToSpecial();
      return;
    }
    if (result.kind === 'report') {
      await showInstanceReport(instance, result.report.title, result.report.items);
      return;
    }
    if (result.kind === 'selected') instance.host.focus();
  };

const buildFormulaAuditAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyFormulaAuditAction'] =>
  async (action) => {
    if (action === 'precedents') {
      instance.tracePrecedents();
      return;
    }
    if (action === 'dependents') {
      instance.traceDependents();
      return;
    }
    if (action === 'clear-all') {
      instance.clearTraces();
      return;
    }
    if (action === 'clear-precedents' || action === 'clear-dependents') {
      clearTraceArrowsByKind(
        instance.store,
        action === 'clear-precedents' ? 'precedent' : 'dependent',
        instance.history,
      );
      instance.host.focus();
      return;
    }
    const map: Record<string, 'errorChecking' | 'traceError' | 'ignoreError'> = {
      'error-checking': 'errorChecking',
      'trace-error': 'traceError',
      'ignore-error': 'ignoreError',
    };
    const auditAction = map[action];
    if (!auditAction) return;
    const result = executeRibbonFormulaAuditingAction({
      store: instance.store,
      workbook: instance.workbook,
      history: instance.history,
      action: auditAction,
      strings: { errorChecking: instance.i18n.strings.ribbonMenu.errorChecking },
    });
    if (result.kind === 'trace-precedents') {
      instance.tracePrecedents();
      return;
    }
    if (result.kind === 'report') {
      await showInstanceReport(instance, result.report.title, result.report.items);
    }
    instance.host.focus();
  };

const buildPasteAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyRibbonPasteAction'] =>
  (action) => {
    if (action === 'dialog') {
      instance.openPasteSpecial();
      return;
    }
    // PASTE_SPECIAL_PRESETS is private to handlePasteAction; the action ids
    // accepted here are the ribbon "paste-action" attribute values:
    //   all | formulas | formulas-and-numfmt | values | values-and-numfmt |
    //   formats | transpose | dialog
    // Map them onto handlePasteAction's PasteAction string so we route
    // through the same `instance.pasteSpecial` / clipboard glue Phase 1.5
    // wired up — that takes care of snapshot fallback for `all` / `values`.
    const map: Record<string, PasteAction> = {
      all: 'paste',
      formulas: 'pasteFormulas',
      'formulas-and-numfmt': 'pasteFormulasNumFmt',
      values: 'pasteValues',
      'values-and-numfmt': 'pasteValuesNumFmt',
      formats: 'pasteFormatsOnly',
      transpose: 'pasteTranspose',
    };
    const mapped = map[action];
    if (mapped === 'paste') {
      dispatchHostClipboard(instance, 'paste');
      return;
    }
    if (mapped) handlePasteAction(instance, mapped);
  };

const buildCellInsertAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyCellInsertAction'] =>
  (action) => {
    const mapped: Record<string, Parameters<typeof handleInsertCellsAction>[1]> = {
      'shift-down': 'shiftDown',
      'shift-right': 'shiftRight',
      rows: 'rows',
      cols: 'cols',
      sheet: 'sheet',
    };
    const next = mapped[action];
    if (!next) return;
    handleInsertCellsAction(instance, next);
    instance.host.focus();
  };

const buildCellDeleteAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyCellDeleteAction'] =>
  (action) => {
    const mapped: Record<string, Parameters<typeof handleDeleteCellsAction>[1]> = {
      'shift-up': 'shiftUp',
      'shift-left': 'shiftLeft',
      rows: 'rows',
      cols: 'cols',
      sheet: 'sheet',
    };
    const next = mapped[action];
    if (!next) return;
    handleDeleteCellsAction(instance, next);
    instance.host.focus();
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
    'tab-color-gray': '#7f7f7f',
  };
  return colors[action];
};

const buildCellFormatAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyCellFormatAction'] =>
  async (action) => {
    const strings = instance.i18n.strings;
    await applyCellFormatAction(action, {
      inst: instance,
      ribbonLang: instance.i18n.locale === 'ja' ? 'ja' : 'en',
      range: normalizedSelectionRange(instance),
      statusMetric: null,
      ribbonMenuText: strings.ribbonMenu,
      renameSheetLabel: strings.sheetTabs.rename,
      runSheetProtectionFlow: async () => {
        instance.toggleSheetProtection();
      },
      showPrompt,
      promptDimension: (title, label, initial, max) =>
        showNumberPrompt({
          title,
          label,
          initial,
          min: 1,
          max,
          okLabel: strings.hyperlinkDialog.ok,
          cancelLabel: strings.hyperlinkDialog.cancel,
        }),
      renderSheetTabs: noop,
      switchSheet: (idx) => {
        mutators.setSheetIndex(instance.store, idx);
      },
      refreshWorkbookCells: () => {
        mutators.replaceCells(
          instance.store,
          instance.workbook.cells(instance.store.getState().data.sheetIndex),
        );
      },
      sheetTabColorByAction,
      projectFormatToolbar: noop,
      focusSheet: () => instance.host.focus(),
    });
    instance.host.focus();
  };

const buildDefinedNameAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyDefinedNameAction'] =>
  async (action) => {
    if (action === 'define') {
      instance.openDefineNameDialog();
      return;
    }
    if (action === 'manager') {
      instance.openNamedRangeDialog();
      return;
    }
    const source =
      action === 'create-top-row'
        ? 'top-row'
        : action === 'create-bottom-row'
          ? 'bottom-row'
          : action === 'create-left-column'
            ? 'left-column'
            : action === 'create-right-column'
              ? 'right-column'
              : null;
    if (source) {
      recordDefinedNamesChange(instance.history, instance.workbook, () =>
        createDefinedNamesFromSelection(instance.store.getState(), instance.workbook, source),
      );
      instance.host.focus();
      return;
    }
    if (action === 'use-formula') {
      const first = listDefinedNames(instance.workbook)[0];
      if (!first) {
        await showReport({
          title: instance.i18n.strings.ribbonMenu.useInFormula,
          items: [
            {
              severity: 'info',
              label: instance.i18n.strings.ribbonMenu.noDefinedNames,
              detail: '',
            },
          ],
          closeLabel: instance.i18n.strings.workbookObjects.close,
          infoLabel: instance.i18n.strings.reviewReports.info,
          warningLabel: instance.i18n.strings.reviewReports.warning,
        });
        return;
      }
      insertDefinedNameFormula(
        instance.store.getState(),
        instance.workbook,
        first.name,
        instance.store,
      );
      mutators.replaceCells(
        instance.store,
        instance.workbook.cells(instance.store.getState().data.sheetIndex),
      );
      instance.host.focus();
    }
  };

const buildLinksAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyLinksAction'] =>
  async (action) => {
    const linkAction =
      action === 'hyperlink'
        ? 'edit'
        : action === 'external' || action === 'open' || action === 'clear'
          ? action
          : null;
    if (!linkAction) return;
    const strings = instance.i18n.strings;
    const result = executeRibbonHyperlinkAction({
      store: instance.store,
      workbook: instance.workbook,
      history: instance.history,
      action: linkAction,
      strings: {
        linkOpen: strings.ribbonMenu.linkOpen,
        linkNoHyperlink: strings.ribbonMenu.linkNoHyperlink,
      },
    });
    if (result.kind === 'open-hyperlink-dialog') {
      instance.openHyperlinkDialog();
      return;
    }
    if (result.kind === 'open-external-dialog') {
      instance.openExternalLinksDialog();
      return;
    }
    if (result.kind === 'open-url') {
      window.open(result.url, '_blank', 'noopener,noreferrer');
      instance.host.focus();
      return;
    }
    if (result.kind === 'report') {
      await showReport({
        title: result.report.title,
        items: result.report.items,
        closeLabel: strings.workbookObjects.close,
        infoLabel: strings.reviewReports.info,
        warningLabel: strings.reviewReports.warning,
      });
      return;
    }
    instance.host.focus();
  };

const buildProtectAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyProtectAction'] =>
  async (action) => {
    if (action === 'protect-sheet' || action === 'unprotect-sheet') {
      instance.setSheetProtected(action === 'protect-sheet');
      instance.host.focus();
      return;
    }
    if (action === 'lock-cell' || action === 'unlock-cell') {
      await buildCellFormatAction(instance)(action);
      return;
    }
    if (action === 'protect-workbook' || action === 'unprotect-workbook') {
      setWorkbookStructureProtected(instance.store, action === 'protect-workbook');
      instance.host.focus();
      return;
    }
    if (action === 'allow-edit-ranges' || action === 'clear-allowed-edit-ranges') {
      const strings = instance.i18n.strings;
      const report = executeRibbonProtectionAction({
        store: instance.store,
        action: action === 'allow-edit-ranges' ? 'allow-edit-range' : 'clear-allowed-edit-ranges',
        strings: {
          allowEditRangesDialogTitle: strings.ribbonMenu.allowEditRangesDialogTitle,
          allowEditRangesCommand: strings.ribbonMenu.allowEditRangesCommand,
          allowEditRangesClearCommand: strings.ribbonMenu.allowEditRangesClearCommand,
          allowedEditRangeAddedStatus: strings.ribbonMenu.allowedEditRangeAddedStatus,
          allowedEditRangesClearedStatus: strings.ribbonMenu.allowedEditRangesClearedStatus,
        },
      });
      await showReport({
        title: report.title,
        items: report.items,
        closeLabel: strings.workbookObjects.close,
        infoLabel: strings.reviewReports.info,
        warningLabel: strings.reviewReports.warning,
      });
    }
  };

const buildTextToColumnsAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['splitTextToColumns'] =>
  (delimiter) => {
    const range = normalizedSelectionRange(instance);
    instance.history.begin();
    try {
      textToColumns(instance.store.getState(), instance.store, instance.workbook, range, delimiter);
      mutators.replaceCells(instance.store, instance.workbook.cells(range.sheet));
    } finally {
      instance.history.end();
    }
    instance.host.focus();
  };

const buildTextToColumnsCustom =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['splitTextToColumnsCustom'] =>
  async () => {
    const strings = instance.i18n.strings;
    const delimiter = await showPrompt({
      title: strings.ribbonMenu.textToColumnsDialogTitle,
      label: strings.ribbonMenu.textToColumnsDialogDelimiters,
      initial: ',',
      okLabel: strings.hyperlinkDialog.ok,
      cancelLabel: strings.hyperlinkDialog.cancel,
      validate: (value) => (value ? null : strings.ribbonMenu.textToColumnsNoDelimited),
    });
    if (delimiter === null) {
      instance.host.focus();
      return;
    }
    await buildTextToColumnsAction(instance)(delimiter);
  };

const buildReviewCommentAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyReviewCommentAction'] =>
  (action) => {
    if (action !== 'delete-active' && action !== 'delete-all') return;
    executeRibbonCommentAction({
      store: instance.store,
      workbook: instance.workbook,
      history: instance.history,
      action,
    });
    instance.host.focus();
  };

const chartKindFromAction = (action: string): SessionChartKind => {
  if (
    action === 'bar' ||
    action === 'line' ||
    action === 'area' ||
    action === 'pie' ||
    action === 'scatter'
  ) {
    return action;
  }
  return 'column';
};

const buildChartAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['createChartFromSelection'] =>
  (kind) => {
    createRibbonChartFromSelection({
      store: instance.store,
      range: normalizedSelectionRange(instance),
      action: kind,
      history: instance.history,
    });
    instance.host.focus();
  };

const buildRecommendedChartAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['createRecommendedChartFromSelection'] =>
  () => {
    createRibbonChartFromSelection({
      store: instance.store,
      range: normalizedSelectionRange(instance),
      action: 'recommended',
      history: instance.history,
    });
    instance.host.focus();
  };

const buildDataValidationAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyDataValidationAction'] =>
  (action) => {
    if (action === 'manage' || action === 'more' || action === 'open' || action === 'settings') {
      instance.openDataValidationDialog();
      return;
    }
    if (action === 'clear-circles') {
      mutators.clearValidationCircles(instance.store);
      instance.host.focus();
      return;
    }
    if (action === 'clear-rules') {
      clearValidationInRangeWithEngine(
        instance.store,
        instance.history,
        instance.workbook,
        normalizedSelectionRange(instance),
      );
      mutators.clearValidationCircles(instance.store);
      instance.host.focus();
      return;
    }
    if (action !== 'circle-invalid') return;
    const state = instance.store.getState();
    const range = normalizedSelectionRange(instance);
    const invalid = new Set<string>();
    for (let row = range.r0; row <= range.r1; row += 1) {
      for (let col = range.c0; col <= range.c1; col += 1) {
        const addr = { sheet: range.sheet, row, col };
        const key = addrKey(addr);
        const validation = state.format.formats.get(key)?.validation;
        if (!validation) continue;
        const value = instance.workbook.getValue(addr);
        if (cellValueViolatesValidation(value, validation)) invalid.add(key);
      }
    }
    mutators.setValidationCircles(instance.store, invalid);
    instance.host.focus();
  };

const buildConditionalMenuAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyConditionalMenuAction'] =>
  async (action, panel) => {
    const strings = instance.i18n.strings;
    const refreshWorkbookCells = (): void => {
      mutators.replaceCells(
        instance.store,
        instance.workbook.cells(instance.store.getState().data.sheetIndex),
      );
    };
    await applyConditionalMenuAction(
      {
        inst: instance,
        ribbonLang: instance.i18n.locale === 'ja' ? 'ja' : 'en',
        range: normalizedSelectionRange(instance),
        cfFill: { fill: '#ffc7ce', color: '#9c0006' },
        cfTopFill: { fill: '#c6efce', color: '#006100' },
        promptCfNumber: (title, initial, options) =>
          showNumberPrompt({
            title,
            label: title,
            initial,
            min: options?.min,
            max: options?.max,
            step: options?.step,
            okLabel: strings.hyperlinkDialog.ok,
            cancelLabel: strings.hyperlinkDialog.cancel,
          }),
        promptCfText: (title, label, initial) =>
          showPrompt({
            title,
            label,
            initial,
            okLabel: strings.hyperlinkDialog.ok,
            cancelLabel: strings.hyperlinkDialog.cancel,
          }),
        showChoiceDialog,
        showMessage,
        refreshWorkbookCells,
        addConditionalRuleFromRibbon: (rule: ConditionalRule) => {
          recordConditionalRulesChange(instance.history, instance.store, () => {
            addConditionalRule(instance.store, rule);
          });
          refreshWorkbookCells();
        },
      },
      action === 'clear' ? 'clear-selection' : action,
      panel,
    );
    instance.host.focus();
  };

const buildUiTheme =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyUiTheme'] =>
  (theme) => {
    instance.setTheme(theme as ThemeName);
  };

const buildPrintAreaAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyPrintAreaAction'] =>
  (action) => {
    const sheet = instance.store.getState().data.sheetIndex;
    if (action === 'clear') clearPrintArea(instance.store, sheet, instance.history);
    else if (action === 'add')
      addPrintArea(
        instance.store,
        sheet,
        formatA1Range(normalizedSelectionRange(instance)),
        instance.history,
      );
    else
      setPrintArea(
        instance.store,
        sheet,
        formatA1Range(normalizedSelectionRange(instance)),
        instance.history,
      );
    instance.host.focus();
  };

const buildPivotTableAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyPivotTableAction'] =>
  async (action) => {
    const pivotAction = (
      action === 'recommended' || action === 'new-sheet' || action === 'existing-sheet'
        ? action
        : 'dialog'
    ) as RibbonPivotTableAction;
    const strings = instance.i18n.strings;
    const result = executeRibbonPivotTableAction({
      store: instance.store,
      workbook: instance.workbook,
      action: pivotAction,
      history: instance.history,
      strings: {
        pivotTable: strings.ribbon.pivotTable,
        pivotTableNewSheet: strings.ribbonMenu.pivotTableNewSheet,
        recommendedPivotTables: strings.ribbonMenu.recommendedPivotTables,
        pivotAuthoringDetail: strings.workbookObjects.compatibilityDetails.pivotAuthoring,
        workbookStructureProtectedBlocked: strings.ribbonMenu.workbookStructureProtectedBlocked,
      },
    });
    if (result.kind === 'open-dialog') {
      instance.openPivotTableDialog();
      return;
    }
    if (result.kind === 'report') {
      await showReport({
        title: result.report.title,
        items: result.report.items,
        closeLabel: strings.workbookObjects.close,
        infoLabel: strings.reviewReports.info,
        warningLabel: strings.reviewReports.warning,
      });
      return;
    }
    instance.host.focus();
  };

const buildPageBreakAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyPageBreakAction'] =>
  (action) => {
    const range = normalizedSelectionRange(instance);
    if (action === 'insert-row')
      insertManualPageBreak(instance.store, range.sheet, 'row', range.r0, instance.history);
    else if (action === 'insert-col')
      insertManualPageBreak(instance.store, range.sheet, 'col', range.c0, instance.history);
    else if (action === 'remove-row')
      removeManualPageBreak(instance.store, range.sheet, 'row', range.r0, instance.history);
    else if (action === 'remove-col')
      removeManualPageBreak(instance.store, range.sheet, 'col', range.c0, instance.history);
    else if (action === 'reset-all')
      resetManualPageBreaks(instance.store, range.sheet, instance.history);
    instance.host.focus();
  };

const buildSheetBackgroundAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applySheetBackgroundAction'] =>
  async (action) => {
    const sheet = instance.store.getState().data.sheetIndex;
    if (action === 'clear') {
      clearSheetBackgroundImage(instance.store, sheet, instance.history);
      instance.host.focus();
      return;
    }
    const strings = instance.i18n.strings;
    const value = await showPrompt({
      title: strings.ribbon.background,
      label: strings.ribbonMenu.sheetBackgroundPrompt,
      initial: instance.store.getState().ui.sheetBackgroundImages.get(sheet) ?? '',
      okLabel: strings.hyperlinkDialog.ok,
      cancelLabel: strings.hyperlinkDialog.cancel,
      validate: (raw) => (raw.trim() ? null : strings.hyperlinkDialog.errorEmptyUrl),
    });
    if (value === null) {
      instance.host.focus();
      return;
    }
    setSheetBackgroundImage(instance.store, sheet, value, instance.history);
    instance.host.focus();
  };

const buildPrintTitlesAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyPrintTitlesAction'] =>
  (action) => {
    const range = normalizedSelectionRange(instance);
    if (action === 'clear') {
      clearPrintTitles(instance.store, range.sheet, instance.history);
    } else if (action === 'rows') {
      setPrintTitleRows(
        instance.store,
        range.sheet,
        `${range.r0 + 1}:${range.r1 + 1}`,
        instance.history,
      );
    } else if (action === 'cols') {
      setPrintTitleCols(
        instance.store,
        range.sheet,
        `${colLetter(range.c0)}:${colLetter(range.c1)}`,
        instance.history,
      );
    }
    instance.host.focus();
  };

const sortColumnsForRange = (range: Range): SortDialogColumn[] =>
  Array.from({ length: range.c1 - range.c0 + 1 }, (_, i) => {
    const col = range.c0 + i;
    const letter = colLetter(col);
    return { value: String(col), label: letter };
  });

const buildSortMenuAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applySortMenuAction'] =>
  async (action) => {
    if (action === 'asc' || action === 'desc') {
      sortActiveColumnAuto({
        store: instance.store,
        workbook: instance.workbook,
        history: instance.history,
        direction: action,
      });
      instance.host.focus();
      return;
    }
    if (action === 'custom') {
      const strings = instance.i18n.strings.ribbonMenu;
      const state = instance.store.getState();
      const range = inferAutoFilterRange(state);
      const columns = sortColumnsForRange(range);
      const result = await showSortDialog({
        title: strings.sortDialogTitle,
        columnLabel: strings.sortDialogColumn,
        orderLabel: strings.sortDialogOrder,
        headerLabel: strings.sortDialogHeader,
        ascendingLabel: strings.sortDialogAscending,
        descendingLabel: strings.sortDialogDescending,
        columns,
        initialColumn: String(Math.min(Math.max(state.selection.active.col, range.c0), range.c1)),
        initialDirection: 'asc',
        initialHasHeader: inferSortHasHeader(state, range),
        okLabel: strings.sortDialogApply,
        cancelLabel: strings.sortDialogCancel,
      });
      if (!result) {
        instance.host.focus();
        return;
      }
      sortRangeWithHistory({
        store: instance.store,
        workbook: instance.workbook,
        history: instance.history,
        range,
        options: {
          byCol: Number(result.column),
          direction: result.direction,
          keys: result.levels.map((level) => ({
            byCol: Number(level.column),
            direction: level.direction,
          })),
          hasHeader: result.hasHeader,
        },
      });
      instance.host.focus();
      return;
    }
    if (action === 'dedupe') {
      const strings = instance.i18n.strings.ribbonMenu;
      const state = instance.store.getState();
      const range = inferAutoFilterRange(state);
      const columns = sortColumnsForRange(range);
      const result = await showRemoveDuplicatesDialog({
        title: strings.removeDuplicatesDialogTitle,
        columnsLabel: strings.removeDuplicatesColumns,
        headerLabel: strings.sortDialogHeader,
        selectAllLabel: strings.removeDuplicatesSelectAll,
        unselectAllLabel: strings.removeDuplicatesUnselectAll,
        noColumnsLabel: strings.removeDuplicatesNoColumns,
        columns,
        initialColumns: columns.map((column) => column.value),
        initialHasHeader: inferSortHasHeader(state, range),
        okLabel: strings.sortDialogApply,
        cancelLabel: strings.sortDialogCancel,
      });
      if (!result) {
        instance.host.focus();
        return;
      }
      instance.history.begin();
      try {
        removeDuplicates(instance.store.getState(), instance.store, instance.workbook, range, {
          columns: result.columns.map(Number),
          hasHeader: result.hasHeader,
        });
      } finally {
        instance.history.end();
      }
      mutators.replaceCells(instance.store, instance.workbook.cells(range.sheet));
      instance.host.focus();
      return;
    }
    if (action === 'conditional') {
      instance.openConditionalDialog();
      return;
    }
    if (action === 'named') {
      instance.openNamedRangeDialog();
      return;
    }
    const filterAction =
      action === 'filter'
        ? 'toggle'
        : action === 'filter-clear'
          ? 'clear'
          : action === 'filter-reapply'
            ? 'reapply'
            : action === 'filter-by-value'
              ? 'filter-by-selected'
              : action === 'filter-advanced'
                ? 'advanced'
                : null;
    if (filterAction) {
      executeRibbonFilterDataAction({
        store: instance.store,
        history: instance.history,
        action: filterAction,
      });
      instance.host.focus();
    }
  };

const buildTableStyleAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['createTableFromSelection'] =>
  (style = 'medium', color, variant = 'banded') => {
    const range = normalizedSelectionRange(instance);
    recordTablesChange(instance.history, instance.store, () => {
      formatAsTableByStyleId(instance.store, range, style, color, variant);
    });
    instance.host.focus();
  };

const buildCellStyleAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyCellStyleFromRibbon'] =>
  (id) => {
    applyCellStyleByName(
      instance.store as unknown as Parameters<typeof applyCellStyleByName>[0],
      instance.history as unknown as Parameters<typeof applyCellStyleByName>[1],
      normalizedSelectionRange(instance),
      id,
    );
    instance.host.focus();
  };

const buildCurrencyPresetAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyCurrencyPreset'] =>
  (symbol) => {
    recordFormatChange(instance.history, instance.store, () => {
      setNumFmt(instance.store.getState(), instance.store, {
        kind: 'currency',
        decimals: 2,
        symbol,
      });
    });
    instance.host.focus();
  };

const buildCurrencyFooterAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['openCurrencyFooterAction'] =>
  (action) => {
    if (action === 'more') instance.openFormatDialog();
  };

const insertSymbolAtActiveCell = (instance: SpreadsheetInstance, symbol: string): void => {
  const addr = instance.store.getState().selection.active;
  instance.history.begin();
  try {
    instance.workbook.setText(addr, symbol);
    mutators.replaceCells(instance.store, instance.workbook.cells(addr.sheet));
  } finally {
    instance.history.end();
  }
};

const buildSymbolAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applySymbolAction'] =>
  async (symbol) => {
    if (symbol === 'more') {
      const selected = await showSymbolDialog({
        text: instance.i18n.strings.ribbonMenu,
        okLabel: instance.i18n.strings.hyperlinkDialog.ok,
        cancelLabel: instance.i18n.strings.hyperlinkDialog.cancel,
      });
      if (selected) insertSymbolAtActiveCell(instance, selected);
      instance.host.focus();
      return;
    }
    insertSymbolAtActiveCell(instance, symbol);
    instance.host.focus();
  };

const showInstanceReport = async (
  instance: SpreadsheetInstance,
  title: string,
  items: { severity: 'info' | 'warning'; label: string; detail: string }[],
): Promise<void> => {
  const strings = instance.i18n.strings;
  await showReport({
    title,
    items,
    closeLabel: strings.workbookObjects.close,
    infoLabel: strings.reviewReports.info,
    warningLabel: strings.reviewReports.warning,
  });
};

const buildScriptAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyScriptAction'] =>
  async (action) => {
    const strings = instance.i18n.strings;
    const raw =
      action === 'custom'
        ? await showPrompt({
            title: strings.ribbonMenu.scriptDialogTitle,
            label: strings.ribbonMenu.scriptDialogCommand,
            initial: 'uppercase',
            okLabel: strings.ribbonMenu.scriptDialogRun,
            cancelLabel: strings.hyperlinkDialog.cancel,
            validate: (value) =>
              parseScriptCommand(value) ? null : strings.ribbonMenu.scriptCommandInvalid,
          })
        : action;
    if (raw === null) {
      instance.host.focus();
      return;
    }
    const command = parseScriptCommand(raw);
    if (!command) return;
    const range = normalizedSelectionRange(instance);
    instance.history.begin();
    let count = 0;
    try {
      count = applyTextScriptToRange(instance.store.getState(), instance.workbook, range, command);
      mutators.replaceCells(instance.store, instance.workbook.cells(range.sheet));
    } finally {
      instance.history.end();
    }
    await showInstanceReport(instance, strings.ribbonMenu.automationScriptsTitle, [
      {
        severity: 'info',
        label: strings.ribbonMenu.automationRunStatus.replace('{count}', String(count)),
        detail: strings.ribbonMenu.automationRunDetail
          .replace('{command}', command)
          .replace('{range}', formatA1Range(range))
          .replace('{count}', String(count)),
      },
    ]);
    instance.host.focus();
  };

const buildPictureAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['insertPictureFromRibbon'] =>
  async (action) => {
    const strings = instance.i18n.strings;
    if (action === 'online') {
      const value = await showPrompt({
        title: strings.ribbonMenu.pictureOnline,
        label: strings.ribbonMenu.pictureUrlPrompt,
        initial: 'https://',
        okLabel: strings.hyperlinkDialog.ok,
        cancelLabel: strings.hyperlinkDialog.cancel,
        validate: (raw) => (raw.trim() ? null : strings.hyperlinkDialog.errorEmptyUrl),
      });
      if (value) {
        createRibbonImageFromSelection(
          instance.store as unknown as Parameters<typeof createRibbonImageFromSelection>[0],
          normalizedSelectionRange(instance),
          value.trim(),
          instance.history as unknown as Parameters<typeof createRibbonImageFromSelection>[3],
        );
      }
      instance.host.focus();
      return;
    }
    if (action === 'device') {
      const result = await pickImageFileDataUrl();
      if (result) {
        createRibbonImageFromSelection(
          instance.store as unknown as Parameters<typeof createRibbonImageFromSelection>[0],
          normalizedSelectionRange(instance),
          result.src,
          instance.history as unknown as Parameters<typeof createRibbonImageFromSelection>[3],
          result.alt,
        );
      }
      instance.host.focus();
      return;
    }
    const label =
      action === 'online' ? strings.ribbonMenu.pictureOnline : strings.ribbonMenu.pictureThisDevice;
    await showInstanceReport(instance, strings.ribbon.pictures, [
      {
        severity: 'info',
        label,
        detail: strings.workbookObjects.compatibilityDetails.chartsDrawings,
      },
    ]);
    instance.host.focus();
  };

const buildShapeAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['insertShapeFromRibbon'] =>
  (shape) => {
    createRibbonShapeFromSelection(
      instance.store as unknown as Parameters<typeof createRibbonShapeFromSelection>[0],
      normalizedSelectionRange(instance),
      shape,
      instance.history as unknown as Parameters<typeof createRibbonShapeFromSelection>[3],
    );
    instance.host.focus();
  };

const activeIllustrationId = (instance: SpreadsheetInstance): string | null => {
  const active = instance.host.querySelector<HTMLElement>('.fc-illustration[aria-selected="true"]');
  if (active?.dataset.illustrationId) return active.dataset.illustrationId;
  const state = instance.store.getState();
  const sheet = state.data.sheetIndex;
  const visible = state.illustrations.illustrations.filter((item) => item.sheet === sheet);
  return visible.at(-1)?.id ?? null;
};

const buildArrangeAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyArrangeAction'] =>
  (action) => {
    if (action === 'selection-pane') {
      instance.openWorkbookObjects();
      return;
    }
    const id = activeIllustrationId(instance);
    if (id) {
      arrangeSessionIllustration(
        instance.store as unknown as Parameters<typeof arrangeSessionIllustration>[0],
        id,
        action,
        instance.history as unknown as Parameters<typeof arrangeSessionIllustration>[3],
      );
    }
    instance.host.focus();
  };

const buildScreenshotAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['insertScreenshotFromRibbon'] =>
  async (action = 'current-view') => {
    const strings = instance.i18n.strings;
    if (action === 'current-view') {
      const canvas = instance.host.querySelector<HTMLCanvasElement>('canvas');
      const dataUrl = canvas?.toDataURL?.('image/png');
      if (dataUrl) {
        createRibbonImageFromSelection(
          instance.store as unknown as Parameters<typeof createRibbonImageFromSelection>[0],
          normalizedSelectionRange(instance),
          dataUrl,
          instance.history as unknown as Parameters<typeof createRibbonImageFromSelection>[3],
        );
        instance.host.focus();
        return;
      }
    } else if (action === 'screen-clipping') {
      const captureScreenClip = (
        instance as unknown as {
          captureScreenClip: () => Promise<{ src: string; alt?: string } | null>;
        }
      ).captureScreenClip;
      const clip = await captureScreenClip();
      if (clip) {
        createRibbonImageFromSelection(
          instance.store as unknown as Parameters<typeof createRibbonImageFromSelection>[0],
          normalizedSelectionRange(instance),
          clip.src,
          instance.history as unknown as Parameters<typeof createRibbonImageFromSelection>[3],
          clip.alt,
        );
        instance.host.focus();
        return;
      }
    }
    const compatibilityDetails = strings.workbookObjects
      .compatibilityDetails as typeof strings.workbookObjects.compatibilityDetails & {
      screenshotCurrentView?: string;
      screenClipping?: string;
    };
    await showInstanceReport(instance, strings.ribbon.screenshot, [
      {
        severity: 'info',
        label:
          action === 'screen-clipping'
            ? strings.ribbonMenu.screenshotScreenClipping
            : strings.ribbonMenu.screenshotCurrentView,
        detail:
          action === 'screen-clipping'
            ? (compatibilityDetails.screenClipping ?? compatibilityDetails.chartsDrawings)
            : (compatibilityDetails.screenshotCurrentView ?? compatibilityDetails.chartsDrawings),
      },
    ]);
    instance.host.focus();
  };

const buildTableStyleFooterAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['openTableStyleFooterAction'] =>
  async (action) => {
    const strings = instance.i18n.strings;
    const activePivot = findPivotTableAtCell(
      instance.workbook as unknown as Parameters<typeof findPivotTableAtCell>[0],
      instance.store.getState().selection.active,
    );
    const state = instance.store.getState();
    const active = state.selection.active;
    const activeTable = tableOverlayAt(state, active.sheet, active.row, active.col);
    const initial = {
      name: strings.ribbonMenu.tableStyleMedium,
      style: activeTable?.style ?? 'medium',
      color: activeTable?.color ?? DEFAULT_TABLE_COLOR,
      variant: tableVariantFromOptions({
        banded: activeTable?.banded ?? true,
        firstCol: activeTable?.firstCol ?? false,
      }),
    } as const;
    if (action === 'new-table-style') {
      const value = await showTableStyleDialog({
        title: strings.ribbonMenu.tableStyleNew,
        strings,
        initial,
      });
      if (value) {
        createTableStyleFromActiveTable(
          instance.store as unknown as Parameters<typeof createTableStyleFromActiveTable>[0],
          instance.history as unknown as Parameters<typeof createTableStyleFromActiveTable>[1],
          value.name,
          {
            style: value.style,
            color: value.color,
            variant: value.variant,
          },
        );
      }
      instance.host.focus();
      return;
    }
    if (action === 'new-pivot-style') {
      const value = await showTableStyleDialog({
        title: strings.ribbonMenu.tableStyleNewPivot,
        strings,
        initial,
      });
      if (value) {
        createPivotTableStyleFromActivePivot(
          instance.store as unknown as Parameters<typeof createPivotTableStyleFromActivePivot>[0],
          instance.history as unknown as Parameters<typeof createPivotTableStyleFromActivePivot>[1],
          value.name,
          activePivot
            ? { sheetIndex: activePivot.sheetIndex, pivotIndex: activePivot.pivotIndex }
            : null,
          {
            style: value.style,
            color: value.color,
            variant: value.variant,
          },
        );
      }
      instance.host.focus();
      return;
    }
    const label =
      action === 'new-pivot-style'
        ? strings.ribbonMenu.tableStyleNewPivot
        : strings.ribbonMenu.tableStyleNew;
    const detail =
      action === 'new-pivot-style'
        ? strings.workbookObjects.compatibilityDetails.pivotAuthoring
        : strings.workbookObjects.compatibilityDetails.formatAsTable;
    await showInstanceReport(instance, label, [{ severity: 'info', label, detail }]);
    instance.host.focus();
  };

const buildPivotTableStyleAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyPivotTableStyleFromRibbon'] =>
  (styleId) => {
    const pivot = findPivotTableAtCell(
      instance.workbook as unknown as Parameters<typeof findPivotTableAtCell>[0],
      instance.store.getState().selection.active,
    );
    if (pivot) {
      applyPivotTableStyleById(
        instance.store as unknown as Parameters<typeof applyPivotTableStyleById>[0],
        instance.history as unknown as Parameters<typeof applyPivotTableStyleById>[1],
        { sheetIndex: pivot.sheetIndex, pivotIndex: pivot.pivotIndex },
        styleId,
      );
    }
    instance.host.focus();
  };

const buildCellStyleFooterAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['openCellStyleFooterAction'] =>
  async (action) => {
    const strings = instance.i18n.strings;
    if (action === 'new-cell-style') {
      const value = await showCellStyleDialog({
        title: strings.ribbonMenu.cellStyleNew,
        strings,
        initialName: strings.ribbonMenu.cellStyleNormal,
      });
      if (value) {
        createCellStyleFromActiveFormat(
          instance.store as unknown as Parameters<typeof createCellStyleFromActiveFormat>[0],
          instance.history as unknown as Parameters<typeof createCellStyleFromActiveFormat>[1],
          normalizedSelectionRange(instance),
          value.name,
          { include: value.include },
        );
      }
      instance.host.focus();
      return;
    }
    if (action === 'merge-cell-style') {
      const result = mergeCellStylesFromWorkbook(
        instance.store as unknown as Parameters<typeof mergeCellStylesFromWorkbook>[0],
        instance.history as unknown as Parameters<typeof mergeCellStylesFromWorkbook>[1],
        instance.workbook as unknown as Parameters<typeof mergeCellStylesFromWorkbook>[2],
      );
      const ribbonMenu = strings.ribbonMenu as typeof strings.ribbonMenu & {
        cellStyleMergeImported?: string;
      };
      const detail =
        result.imported > 0
          ? (ribbonMenu.cellStyleMergeImported ?? '{count} style(s) imported.').replace(
              '{count}',
              String(result.imported),
            )
          : strings.workbookObjects.compatibilityDetails.cellFormatting;
      await showInstanceReport(instance, strings.ribbonMenu.cellStyleMerge, [
        {
          severity: result.imported > 0 ? 'info' : 'warning',
          label: strings.ribbonMenu.cellStyleMerge,
          detail,
        },
      ]);
      instance.host.focus();
      return;
    }
    const label = strings.ribbonMenu.cellStyleNew;
    await showInstanceReport(instance, label, [
      {
        severity: 'info',
        label,
        detail: strings.workbookObjects.compatibilityDetails.cellFormatting,
      },
    ]);
    instance.host.focus();
  };

const buildPdfAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyPdfAction'] =>
  async (action) => {
    const strings = instance.i18n.strings;
    const result = resolveRibbonPdfAction(action as RibbonPdfAction, {
      cellMenu: strings.ribbonMenu,
      pdfTitle: strings.ribbonMenu.pdfCreate,
    });
    if (result.kind === 'open-page-setup') {
      instance.openPageSetup();
      return;
    }
    instance.print('pdf');
    if (result.report) await showInstanceReport(instance, result.report.title, result.report.items);
    instance.host.focus();
  };

const buildAddInAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyAddInAction'] =>
  async (action) => {
    const strings = instance.i18n.strings;
    const report = buildRibbonAddInReport(action as RibbonAddInAction, {
      cellMenu: strings.ribbonMenu,
      addInDefaultTitle: strings.ribbon.addIn,
    });
    if (report) await showInstanceReport(instance, report.title, report.items);
    instance.host.focus();
  };

export function createDefaultDynamicDropdownsCtx(
  instance: SpreadsheetInstance,
  opts: DefaultDynamicDropdownsOptions = {},
): DynamicDropdownsCtx {
  const focusSheet = (): void => {
    instance.host.focus();
  };

  const base: DynamicDropdownsCtx = {
    getInst: () => instance,
    updateCalcOptionsMenu: noop,
    updateDefinedNamesMenu: noop,
    closeBorderMenu: noop,
    closeFreezeMenu: noop,
    closePrintAreaMenu: noop,
    closeSymbolMenu: noop,
    getConditionalMenu: () => document.getElementById('menu-conditional') as HTMLElement | null,
    focusSheet,

    // Pure / instance-derivable defaults.
    applyRibbonPasteAction: buildPasteAction(instance),
    applyFillSeries: buildFillSeries(instance),
    applyFillDirection: buildFillDirection(instance),
    applyClearAction: buildClearAction(instance),
    applyFreezeAction: buildFreezeAction(instance),
    applyTextOrientationAction: buildTextOrientation(instance),
    applyAutoSumFormula: buildAutoSumFormula(instance),
    applyFormulaAuditAction: buildFormulaAuditAction(instance),
    applyWatchAction: buildWatchAction(instance),
    applyCalcOptionAction: buildCalcOptionAction(instance),
    applyFindSelectAction: buildFindSelectAction(instance),
    applyDataValidationAction: buildDataValidationAction(instance),
    applyConditionalMenuAction: buildConditionalMenuAction(instance),
    applyUiTheme: buildUiTheme(instance),
    applySymbolAction: buildSymbolAction(instance),

    // Dialog / host-glue — host opts in via `overrides`. Defaults are no-op
    // so the click delegator returns true (event consumed) and the open
    // menu still closes, instead of falling through to the legacy fallback.
    applyPivotTableAction: buildPivotTableAction(instance),
    applyDefinedNameAction: buildDefinedNameAction(instance),
    applyLinksAction: buildLinksAction(instance),
    applyCellInsertAction: buildCellInsertAction(instance),
    applyCellDeleteAction: buildCellDeleteAction(instance),
    applyCellFormatAction: buildCellFormatAction(instance),
    applyPageBreakAction: buildPageBreakAction(instance),
    applySheetBackgroundAction: buildSheetBackgroundAction(instance),
    applyPrintAreaAction: buildPrintAreaAction(instance),
    applyPrintTitlesAction: buildPrintTitlesAction(instance),
    applyArrangeAction: buildArrangeAction(instance),
    applySortMenuAction: buildSortMenuAction(instance),
    applyReviewCommentAction: buildReviewCommentAction(instance),
    applyProtectAction: buildProtectAction(instance),
    createRecommendedChartFromSelection: buildRecommendedChartAction(instance),
    createChartFromSelection: buildChartAction(instance),
    chartKindFromAction,
    insertPictureFromRibbon: buildPictureAction(instance),
    insertShapeFromRibbon: buildShapeAction(instance),
    insertScreenshotFromRibbon: buildScreenshotAction(instance),
    applyScriptAction: buildScriptAction(instance),
    applyPdfAction: buildPdfAction(instance),
    createTableFromSelection: buildTableStyleAction(instance),
    openTableStyleFooterAction: buildTableStyleFooterAction(instance),
    applyPivotTableStyleFromRibbon: buildPivotTableStyleAction(instance),
    applyCellStyleFromRibbon: buildCellStyleAction(instance),
    openCellStyleFooterAction: buildCellStyleFooterAction(instance),
    applyCurrencyPreset: buildCurrencyPresetAction(instance),
    openCurrencyFooterAction: buildCurrencyFooterAction(instance),
    splitTextToColumns: buildTextToColumnsAction(instance),
    splitTextToColumnsCustom: buildTextToColumnsCustom(instance),
    applyAddInAction: buildAddInAction(instance),
  };

  // Object form: spread once at construction. Function form: build a live
  // ctx whose every property re-reads the latest override on each access so
  // hosts can swap handlers post-mount without re-creating the api.
  if (typeof opts.overrides !== 'function') {
    return { ...base, ...opts.overrides };
  }
  const getOverrides = opts.overrides;
  const ctx = {} as DynamicDropdownsCtx;
  for (const key of Object.keys(base) as (keyof DynamicDropdownsCtx)[]) {
    Object.defineProperty(ctx, key, {
      enumerable: true,
      get() {
        const override = (getOverrides() as Partial<DynamicDropdownsCtx>)[key];
        return override !== undefined ? override : base[key];
      },
    });
  }
  return ctx;
}
