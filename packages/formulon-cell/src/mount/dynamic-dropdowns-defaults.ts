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
  applyAdvancedFilter,
  applyMerge,
  applyTextScriptToRange,
  applyUnmerge,
  autoSum,
  buildRibbonAddInReport,
  type ConditionalRule,
  cellValueIsFormulaError,
  clearPrintArea,
  clearSheetBackgroundImage,
  clearTraceArrowsByKind,
  clearValidationInRangeWithEngine,
  clearWatchedCells,
  colLetter,
  commentAt,
  conditionalRulesForRange,
  copyAdvancedFilterResult,
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
  getPageSetup,
  handleDeleteCellsAction,
  handleFreezeAction,
  handleInsertCellsAction,
  handlePasteAction,
  hiddenInSelection,
  inferAutoFilterRange,
  inferFillSeriesDirection,
  inferSortHasHeader,
  insertDefinedNameFormula,
  insertManualPageBreak,
  listComments,
  listDefinedNames,
  mergeWillLoseData,
  mutators,
  type PasteAction,
  parseA1Range,
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
  setAlign,
  setNumFmt,
  setPrintArea,
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
import { applyFormatPatch } from '../commands/format.js';
import {
  applyPivotTableStyleById,
  type CustomTableStyle,
  createPivotTableStyleFromActivePivot,
  createTableStyleFromActiveTable,
  DEFAULT_TABLE_COLOR,
  formatAsTableByStyleId,
  inferTableHasHeaders,
  pivotTableStyleAssignment,
  tableOverlayAt,
  tableVariantFromOptions,
} from '../commands/format-as-table.js';
import { hyperlinkAt } from '../commands/hyperlinks.js';
import {
  addAllowedEditRange,
  isWorkbookStructureProtected,
  protectedSheetPassword,
  protectedSheetPasswordHash,
  protectedSheetPermissions,
  verifySheetProtectionPasswordHash,
} from '../commands/protection.js';
import {
  arrangeSessionIllustration,
  createRibbonImageFromSelection,
  createRibbonShapeFromSelection,
} from '../commands/session-illustration.js';
import { cellValueViolatesValidation } from '../commands/validate.js';
import { addrKey } from '../engine/address.js';
import { findPivotTableAtCell } from '../engine/passthrough-sync.js';
import { sheetTabColorActionForColor, sheetTabColorByAction } from '../sheet-tab-colors.js';
import { formatWithPending } from '../store/pending-format.js';
import { showAdvancedFilterDialog } from '../toolbar/dialogs/advanced-filter.js';
import { showCellStyleDialog } from '../toolbar/dialogs/cell-style.js';
import { showChoiceDialog } from '../toolbar/dialogs/choice.js';
import {
  showConditionalFormatNumberDialog,
  showConditionalFormatTextDialog,
} from '../toolbar/dialogs/conditional-format.js';
import { showDefinedNamePickerDialog } from '../toolbar/dialogs/defined-name-picker.js';
import { showDimensionDialog } from '../toolbar/dialogs/dimension.js';
import { showFormatAsTableDialog } from '../toolbar/dialogs/format-as-table.js';
import { pickImageFileDataUrl } from '../toolbar/dialogs/image-file.js';
import { confirmMergeLoseData } from '../toolbar/dialogs/merge-confirm.js';
import { showMessage } from '../toolbar/dialogs/prompt.js';
import {
  showAllowEditRangeDialog,
  showProtectSheetDialog,
  showUnprotectSheetDialog,
} from '../toolbar/dialogs/protection.js';
import { showRemoveDuplicatesDialog } from '../toolbar/dialogs/remove-duplicates.js';
import { showRenameSheetDialog } from '../toolbar/dialogs/rename-sheet.js';
import { reportDialogLabels, showReport } from '../toolbar/dialogs/report.js';
import { showScriptCommandDialog } from '../toolbar/dialogs/script-command.js';
import { type SortDialogColumn, showSortDialog } from '../toolbar/dialogs/sort.js';
import { showSymbolDialog } from '../toolbar/dialogs/symbol.js';
import { showTableStyleDialog } from '../toolbar/dialogs/table-style.js';
import { showTextToColumnsDialog } from '../toolbar/dialogs/text-to-columns.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import { applyCellFormatAction } from '../toolbar/ribbon/cell-format-action.js';
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
  focusSheet?: () => void;
  projectFormatToolbar?: () => void;
  refreshCells?: () => void;
  renderSheetTabs?: () => void;
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

const addrFromKey = (key: string): { sheet: number; row: number; col: number } | null => {
  const [sheetRaw, rowRaw, colRaw] = key.split(':');
  const sheet = Number(sheetRaw);
  const row = Number(rowRaw);
  const col = Number(colRaw);
  if (!Number.isInteger(sheet) || !Number.isInteger(row) || !Number.isInteger(col)) return null;
  return { sheet, row, col };
};

const addrInRange = (addr: { sheet: number; row: number; col: number }, range: Range): boolean =>
  addr.sheet === range.sheet &&
  addr.row >= range.r0 &&
  addr.row <= range.r1 &&
  addr.col >= range.c0 &&
  addr.col <= range.c1;

const MAX_MERGE_ACROSS_ROWS = 100_000;

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

const setMenuControlDisabled = (
  button: HTMLButtonElement,
  disabled: boolean,
  reason?: string,
): void => {
  const baseTitle = button.dataset.menuBaseTitle ?? button.title;
  button.dataset.menuBaseTitle = baseTitle;
  projectDisabledState(button, disabled, reason ?? null, {
    datasetKey: 'menuDisabledReason',
    titlePrefix: baseTitle,
  });
};

const updateFillMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateFillMenu'] =>
  (menu) => {
    const range = normalizedSelectionRange(instance);
    const hasMultipleRows = range.r1 > range.r0;
    const hasMultipleCols = range.c1 > range.c0;
    const strings = instance.i18n.strings.ribbonMenu;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-fill]')) {
      const action = button.dataset.fill;
      const disabled =
        ((action === 'down' || action === 'up') && !hasMultipleRows) ||
        ((action === 'right' || action === 'left') && !hasMultipleCols);
      const reason =
        action === 'down' || action === 'up'
          ? strings.fillRequiresMultipleRows
          : action === 'right' || action === 'left'
            ? strings.fillRequiresMultipleCols
            : undefined;
      setMenuControlDisabled(button, disabled, reason);
    }
  };

const buildFillSeries =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyFillSeries'] =>
  async (mode) => {
    const range = normalizedSelectionRange(instance);
    const fillSeriesDialogStrings = (
      instance.i18n.strings as typeof instance.i18n.strings & {
        fillSeriesDialog: {
          title: string;
          seriesIn: string;
          columns: string;
          rows: string;
          up: string;
          left: string;
          type: string;
          autoFill: string;
          copy: string;
          day: string;
          weekday: string;
          month: string;
          year: string;
          ok: string;
          cancel: string;
        };
      }
    ).fillSeriesDialog;
    const choice = mode
      ? { direction: inferFillSeriesDirection(range), mode }
      : await showFillSeriesDialog(range, fillSeriesDialogStrings);
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

const visualClearFormatKeys = new Set([
  'cellStyle',
  'numFmt',
  'bold',
  'italic',
  'underline',
  'strike',
  'align',
  'vAlign',
  'wrap',
  'shrinkToFit',
  'indent',
  'rotation',
  'textDirection',
  'borders',
  'color',
  'fill',
  'fillPattern',
  'fillPatternColor',
  'fontFamily',
  'fontSize',
]);

const hasClearableVisualFormat = (format: object): boolean =>
  Object.keys(format).some((key) => visualClearFormatKeys.has(key));

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

const updateClearMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateClearMenu'] =>
  (menu) => {
    const state = instance.store.getState();
    const range = state.selection.range;
    let hasContents = false;
    let hasFormats = false;
    let hasComments = false;
    let hasHyperlinks = false;
    const pending = (
      state.ui as typeof state.ui & {
        pendingFormat?: {
          addr: { sheet: number; row: number; col: number };
          format: object;
        } | null;
      }
    ).pendingFormat;
    if (
      pending &&
      pending.addr.sheet === range.sheet &&
      pending.addr.row >= range.r0 &&
      pending.addr.row <= range.r1 &&
      pending.addr.col >= range.c0 &&
      pending.addr.col <= range.c1 &&
      Object.keys(pending.format).length > 0
    ) {
      hasFormats = true;
    }
    for (const key of state.data.cells.keys()) {
      const addr = addrFromKey(key);
      if (addr && addrInRange(addr, range)) {
        hasContents = true;
        break;
      }
    }
    for (const [key, format] of state.format.formats) {
      const addr = addrFromKey(key);
      if (!addr || !addrInRange(addr, range)) continue;
      if (typeof format.comment === 'string' && format.comment.length > 0) hasComments = true;
      if (typeof format.hyperlink === 'string' && format.hyperlink.length > 0) {
        hasHyperlinks = true;
      }
      if (hasClearableVisualFormat(format)) hasFormats = true;
      if (hasFormats && hasComments && hasHyperlinks) break;
    }
    const hasConditional = conditionalRulesForRange(state, range).length > 0;
    const hasAny = hasContents || hasFormats || hasComments || hasHyperlinks || hasConditional;
    const disabledReason = instance.i18n.strings.ribbon.clearRequiresTarget;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-clear]')) {
      const action = button.dataset.clear;
      const disabled =
        (action === 'all' && !hasAny) ||
        (action === 'contents' && !hasContents) ||
        (action === 'formats' && !hasFormats) ||
        (action === 'comments' && !hasComments) ||
        ((action === 'hyperlinks' || action === 'remove-hyperlinks') && !hasHyperlinks) ||
        (action === 'conditional' && !hasConditional);
      setMenuControlDisabled(button, disabled, disabledReason);
    }
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

const updateFreezeMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateFreezeMenu'] =>
  (menu) => {
    const state = instance.store.getState();
    const hasFreeze = state.layout.freezeRows > 0 || state.layout.freezeCols > 0;
    const strings = instance.i18n.strings;
    const toolbarStrings = strings.toolbar as typeof strings.toolbar & { unfreezePanes?: string };
    const primary = menu.querySelector<HTMLButtonElement>(
      '[data-freeze="selection"], [data-freeze="off"]',
    );
    if (primary) {
      primary.dataset.freeze = hasFreeze ? 'off' : 'selection';
      setMenuControlDisabled(primary, false);
      const text = primary.querySelector<HTMLElement>('.app__menu-item__text');
      if (text) {
        text.textContent = hasFreeze
          ? (toolbarStrings.unfreezePanes ?? strings.toolbar.unfreeze)
          : strings.viewToolbar.freezePanes;
      }
      const icon = primary.querySelector<HTMLElement>('.app__menu-icon');
      icon?.classList.toggle('app__menu-icon--freeze-panes', !hasFreeze);
      icon?.classList.toggle('app__menu-icon--freeze-off', hasFreeze);
    }
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-freeze]')) {
      if (button === primary) continue;
      setMenuControlDisabled(button, false);
    }
  };

const buildUnderlineAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyUnderlineAction'] =>
  async (action) => {
    const strings = instance.i18n.strings;
    const ribbonMenu = strings.ribbonMenu as typeof strings.ribbonMenu & {
      underlineDouble: string;
    };
    const doubleUnderlineLabel = ribbonMenu.underlineDouble;
    if (action === 'single') {
      recordFormatChange(instance.history, instance.store, () => {
        applyFormatPatch(
          instance.store.getState(),
          instance.store,
          instance.store.getState().selection.range,
          { underline: true },
        );
      });
      instance.host.focus();
      return;
    }
    await showInstanceReport(instance, doubleUnderlineLabel, [
      {
        severity: 'warning',
        label: doubleUnderlineLabel,
        detail: strings.workbookObjects.compatibilityDetails.cellFormatting,
      },
    ]);
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

const buildMergeAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyMergeAction'] =>
  async (action) => {
    const range = instance.store.getState().selection.range;
    if (action === 'unmergeCells') {
      applyUnmerge(instance.store, instance.workbook, instance.history, range);
    } else if (action === 'mergeAcross') {
      if (range.c0 === range.c1 || range.r1 - range.r0 + 1 > MAX_MERGE_ACROSS_ROWS) return;
      const state = instance.store.getState();
      if (
        mergeWillLoseData(state, range) &&
        !(await confirmMergeLoseData(instance.i18n.strings, state, range))
      ) {
        return;
      }
      instance.history.begin();
      try {
        for (let row = range.r0; row <= range.r1; row += 1) {
          applyMerge(instance.store, instance.workbook, instance.history, {
            sheet: range.sheet,
            r0: row,
            c0: range.c0,
            r1: row,
            c1: range.c1,
          });
        }
      } finally {
        instance.history.end();
      }
    } else {
      const state = instance.store.getState();
      if (
        mergeWillLoseData(state, range) &&
        !(await confirmMergeLoseData(instance.i18n.strings, state, range))
      ) {
        return;
      }
      const merged = applyMerge(instance.store, instance.workbook, instance.history, range);
      if (merged && action === 'mergeCenter') {
        recordFormatChange(instance.history, instance.store, () => {
          setAlign(instance.store.getState(), instance.store, 'center');
        });
      }
    }
    instance.host.focus();
  };

const updateTextOrientationMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateTextOrientationMenu'] =>
  (menu) => {
    const state = instance.store.getState();
    const active = state.selection.active;
    const rotation = formatWithPending(state, active)?.rotation ?? 0;
    const current =
      rotation === 45
        ? 'ccw'
        : rotation === -45
          ? 'cw'
          : rotation === 90
            ? 'up'
            : rotation === -90
              ? 'down'
              : 'horizontal';
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-text-orientation]')) {
      const action = button.dataset.textOrientation;
      if (action === 'format') continue;
      const activeItem = action === current;
      button.setAttribute('role', 'menuitemradio');
      button.setAttribute('aria-checked', String(activeItem));
      button.classList.toggle('app__menu-item--active', activeItem);
    }
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

const updateReviewCommentsMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateReviewCommentsMenu'] =>
  (menu) => {
    const state = instance.store.getState();
    const hasActiveComment = commentAt(state, state.selection.active) !== null;
    const hasAnyComment = listComments(state, state.data.sheetIndex).length > 0;
    const strings = instance.i18n.strings.ribbonMenu;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-comment-action]')) {
      const action = button.dataset.commentAction;
      const disabled =
        (action === 'delete-active' && !hasActiveComment) ||
        (action === 'delete-all' && !hasAnyComment);
      const reason =
        action === 'delete-active' ? strings.commentDeleteRequiresActive : strings.commentNone;
      setMenuControlDisabled(button, disabled, reason);
    }
  };

const updateWatchMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateWatchMenu'] =>
  (menu) => {
    const state = instance.store.getState();
    const active = state.selection.active;
    const hasWatches = state.watch.watches.length > 0;
    const activeWatched = state.watch.watches.some(
      (watch) =>
        watch.sheet === active.sheet && watch.row === active.row && watch.col === active.col,
    );
    const strings = instance.i18n.strings.ribbonMenu;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-watch-action]')) {
      const action = button.dataset.watchAction;
      const disabled =
        (action === 'delete' && !activeWatched) || (action === 'delete-all' && !hasWatches);
      const reason =
        action === 'delete' ? strings.watchDeleteRequiresActive : strings.watchDeleteAllRequiresAny;
      setMenuControlDisabled(button, disabled, reason);
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

const calcOptionForMode = (mode: 0 | 1 | 2 | null): 'auto' | 'manual' | 'auto-no-table' | null => {
  if (mode === 0) return 'auto';
  if (mode === 1) return 'manual';
  if (mode === 2) return 'auto-no-table';
  return null;
};

const updateCalcOptionsMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateCalcOptionsMenu'] =>
  (menu) => {
    const current = calcOptionForMode(instance.workbook.calcMode());
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[role="menuitemradio"]')) {
      const active = button.dataset.calcOption === current;
      button.setAttribute('aria-checked', String(active));
      button.classList.toggle('app__menu-item--active', active);
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

const updateErrorCheckingMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateErrorCheckingMenu'] =>
  (menu) => {
    const state = instance.store.getState();
    const active = state.selection.active;
    const cell = state.data.cells.get(addrKey(active));
    const activeFormulaError = Boolean(cell?.formula && cellValueIsFormulaError(cell.value));
    const disabledReason = instance.i18n.strings.ribbonMenu.traceErrorRequiresFormulaError;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-formula-audit-action]')) {
      const action = button.dataset.formulaAuditAction;
      const disabled =
        (action === 'trace-error' || action === 'ignore-error') && !activeFormulaError;
      setMenuControlDisabled(button, disabled, disabledReason);
    }
  };

const updateClearArrowsMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateClearArrowsMenu'] =>
  (menu) => {
    const traces = instance.store.getState().traces.items;
    const hasPrecedents = traces.some((trace) => trace.kind === 'precedent');
    const hasDependents = traces.some((trace) => trace.kind === 'dependent');
    const hasAny = hasPrecedents || hasDependents;
    const strings = instance.i18n.strings.ribbonMenu;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-formula-audit-action]')) {
      const action = button.dataset.formulaAuditAction;
      const disabled =
        (action === 'clear-all' && !hasAny) ||
        (action === 'clear-precedents' && !hasPrecedents) ||
        (action === 'clear-dependents' && !hasDependents);
      const reason =
        action === 'clear-precedents'
          ? strings.removePrecedentArrowsRequiresAny
          : action === 'clear-dependents'
            ? strings.removeDependentArrowsRequiresAny
            : strings.removeArrowsRequiresAny;
      setMenuControlDisabled(button, disabled, reason);
    }
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

const updatePasteMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updatePasteMenu'] =>
  (menu) => {
    const hasSnapshot = instance.clipboard?.getSnapshot() != null;
    const disabledReason = instance.i18n.strings.ribbon.pasteRequiresClipboard;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-paste-action]')) {
      const action = button.dataset.pasteAction;
      const disabled = action !== 'all' && !hasSnapshot;
      setMenuControlDisabled(button, disabled, disabledReason);
    }
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

const buildCellFormatAction =
  (
    instance: SpreadsheetInstance,
    opts: Pick<
      DefaultDynamicDropdownsOptions,
      'projectFormatToolbar' | 'refreshCells' | 'renderSheetTabs'
    >,
  ): DynamicDropdownsCtx['applyCellFormatAction'] =>
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
      showRenameSheetDialog: (opts) =>
        showRenameSheetDialog({
          okLabel: strings.hyperlinkDialog.ok,
          cancelLabel: strings.hyperlinkDialog.cancel,
          ...opts,
        }),
      promptDimension: (title, label, initial, max) =>
        showDimensionDialog({
          title,
          label,
          initial,
          max,
          okLabel: strings.hyperlinkDialog.ok,
          cancelLabel: strings.hyperlinkDialog.cancel,
        }),
      renderSheetTabs: opts.renderSheetTabs ?? noop,
      switchSheet: (idx) => {
        mutators.setSheetIndex(instance.store, idx);
      },
      refreshWorkbookCells:
        opts.refreshCells ??
        (() => {
          mutators.replaceCells(
            instance.store,
            instance.workbook.cells(instance.store.getState().data.sheetIndex),
          );
        }),
      sheetTabColorByAction,
      projectFormatToolbar: opts.projectFormatToolbar ?? noop,
      focusSheet: () => instance.host.focus(),
    });
    instance.host.focus();
  };

const updateFormatCellsMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateFormatCellsMenu'] =>
  (menu) => {
    const state = instance.store.getState();
    const range = normalizedSelectionRange(instance);
    const rowsHidden = hiddenInSelection(state.layout, 'row', range.r0, range.r1).length > 0;
    const colsHidden = hiddenInSelection(state.layout, 'col', range.c0, range.c1).length > 0;
    const sheet = state.data.sheetIndex;
    const hiddenSheetCount = state.layout.hiddenSheets.size;
    const visibleSheetCount = instance.workbook.sheetCount - hiddenSheetCount;
    const canMoveSheet = instance.workbook.capabilities.sheetMutate;
    const activeTabColorAction = sheetTabColorActionForColor(
      state.layout.sheetTabColors.get(sheet),
    );
    const reasonForFormatAction = (action: string | undefined): string | undefined => {
      const t = instance.i18n.strings.ribbonMenu;
      if (action === 'show-rows' && !rowsHidden) return t.formatNoHiddenRows;
      if (action === 'show-cols' && !colsHidden) return t.formatNoHiddenCols;
      if (
        (action === 'rename-sheet' ||
          action === 'move-sheet-left' ||
          action === 'move-sheet-right' ||
          action === 'hide-sheet' ||
          action === 'unhide-sheet') &&
        !canMoveSheet
      ) {
        return t.sheetActionUnavailable;
      }
      if (action === 'move-sheet-left' && sheet <= 0) return t.sheetMoveAtBoundary;
      if (action === 'move-sheet-right' && sheet >= instance.workbook.sheetCount - 1) {
        return t.sheetMoveAtBoundary;
      }
      if (action === 'hide-sheet' && visibleSheetCount <= 1) return t.sheetHideRequiresVisibleSheet;
      if (action === 'unhide-sheet' && hiddenSheetCount === 0) {
        return t.sheetUnhideRequiresHiddenSheet;
      }
      return undefined;
    };
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-cell-format]')) {
      const action = button.dataset.cellFormat;
      const disabled =
        (action === 'show-rows' && !rowsHidden) ||
        (action === 'show-cols' && !colsHidden) ||
        (action === 'rename-sheet' && !canMoveSheet) ||
        (action === 'move-sheet-left' && (sheet <= 0 || !canMoveSheet)) ||
        (action === 'move-sheet-right' &&
          (sheet >= instance.workbook.sheetCount - 1 || !canMoveSheet)) ||
        (action === 'hide-sheet' && visibleSheetCount <= 1) ||
        (action === 'unhide-sheet' && hiddenSheetCount === 0);
      setMenuControlDisabled(button, disabled, reasonForFormatAction(action));
      if (action?.startsWith('tab-color-')) {
        const active = action === activeTabColorAction;
        button.setAttribute('role', 'menuitemradio');
        button.setAttribute('aria-checked', String(active));
        button.classList.toggle('app__color-swatch--active', active);
      }
    }
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
      const names = listDefinedNames(instance.workbook);
      if (names.length === 0) {
        await showReport({
          title: instance.i18n.strings.ribbonMenu.useInFormula,
          items: [
            {
              severity: 'info',
              label: instance.i18n.strings.ribbonMenu.noDefinedNames,
              detail: '',
            },
          ],
          ...reportDialogLabels(instance.i18n.strings),
        });
        return;
      }
      const selected = await showDefinedNamePickerDialog({
        title: instance.i18n.strings.ribbonMenu.useInFormula,
        names,
        okLabel: instance.i18n.strings.hyperlinkDialog.ok,
        cancelLabel: instance.i18n.strings.hyperlinkDialog.cancel,
      });
      if (!selected) {
        instance.host.focus();
        return;
      }
      insertDefinedNameFormula(
        instance.store.getState(),
        instance.workbook,
        selected,
        instance.store,
      );
      mutators.replaceCells(
        instance.store,
        instance.workbook.cells(instance.store.getState().data.sheetIndex),
      );
      instance.host.focus();
    }
  };

const updateDefinedNamesMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateDefinedNamesMenu'] =>
  (menu) => {
    const hasDefinedNames = listDefinedNames(instance.workbook).length > 0;
    const canMutateNames = instance.workbook.capabilities.definedNameMutate;
    const strings = instance.i18n.strings.ribbonMenu;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-defined-name-action]')) {
      const action = button.dataset.definedNameAction;
      const disabled =
        action === 'use-formula'
          ? !hasDefinedNames
          : action?.startsWith('create-') === true && !canMutateNames;
      const reason =
        action === 'use-formula' ? strings.noDefinedNames : strings.definedNameMutationUnavailable;
      setMenuControlDisabled(button, disabled, reason);
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
        ...reportDialogLabels(strings),
      });
      return;
    }
    instance.host.focus();
  };

const updateLinksMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateLinksMenu'] =>
  (menu) => {
    const state = instance.store.getState();
    const hasHyperlink = hyperlinkAt(state, state.selection.active) !== null;
    const disabledReason = instance.i18n.strings.ribbonMenu.linkNoHyperlink;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-link-action]')) {
      const action = button.dataset.linkAction;
      const disabled = (action === 'open' || action === 'clear') && !hasHyperlink;
      setMenuControlDisabled(button, disabled, disabledReason);
    }
  };

const buildProtectAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyProtectAction'] =>
  async (action) => {
    if (action === 'protect-sheet') {
      const strings = instance.i18n.strings;
      const protectionStrings = strings.protection as typeof strings.protection & {
        allowAutoFilter: string;
        allowDeleteColumns: string;
        allowDeleteRows: string;
        allowEditObjects: string;
        allowEditScenarios: string;
        allowFormatCells: string;
        allowFormatColumns: string;
        allowFormatRows: string;
        allowInsertColumns: string;
        allowInsertHyperlinks: string;
        allowInsertRows: string;
        allowPivotTables: string;
        allowSelectLockedCells: string;
        allowSelectUnlockedCells: string;
        allowSort: string;
        allowUsersTo: string;
        confirmPassword: string;
        passwordMismatch: string;
      };
      const result = await showProtectSheetDialog({
        strings: {
          title: protectionStrings.protectSheet,
          password: protectionStrings.password,
          passwordPlaceholder: protectionStrings.passwordPlaceholder,
          confirmPassword: protectionStrings.confirmPassword,
          passwordMismatch: protectionStrings.passwordMismatch,
          allowLabel: protectionStrings.allowUsersTo,
          allowSelectLockedCells: protectionStrings.allowSelectLockedCells,
          allowSelectUnlockedCells: protectionStrings.allowSelectUnlockedCells,
          allowFormatCells: protectionStrings.allowFormatCells,
          allowFormatColumns: protectionStrings.allowFormatColumns,
          allowFormatRows: protectionStrings.allowFormatRows,
          allowInsertColumns: protectionStrings.allowInsertColumns,
          allowInsertRows: protectionStrings.allowInsertRows,
          allowInsertHyperlinks: protectionStrings.allowInsertHyperlinks,
          allowDeleteColumns: protectionStrings.allowDeleteColumns,
          allowDeleteRows: protectionStrings.allowDeleteRows,
          allowSort: protectionStrings.allowSort,
          allowAutoFilter: protectionStrings.allowAutoFilter,
          allowPivotTables: protectionStrings.allowPivotTables,
          allowEditObjects: protectionStrings.allowEditObjects,
          allowEditScenarios: protectionStrings.allowEditScenarios,
          ok: strings.pageSetup.ok,
          cancel: strings.pageSetup.cancel,
        },
        initial: protectedSheetPermissions(
          instance.store.getState(),
          instance.store.getState().data.sheetIndex,
        ),
      });
      if (!result) {
        instance.host.focus();
        return;
      }
      (
        instance.setSheetProtected as (
          on: boolean,
          password?: string,
          permissions?: import('../store/store.js').SheetProtectionPermissions,
        ) => void
      )(true, result.password, result.permissions);
      instance.host.focus();
      return;
    }
    if (action === 'unprotect-sheet') {
      const strings = instance.i18n.strings;
      const sheet = instance.store.getState().data.sheetIndex;
      const currentPassword = protectedSheetPassword(instance.store.getState(), sheet);
      const currentPasswordHash = protectedSheetPasswordHash(instance.store.getState(), sheet);
      if (currentPassword || currentPasswordHash) {
        const password = await showUnprotectSheetDialog({
          title: strings.protection.unprotectSheet,
          password: strings.protection.password,
          ok: strings.pageSetup.ok,
          cancel: strings.pageSetup.cancel,
        });
        if (password === null) {
          instance.host.focus();
          return;
        }
        const matches =
          currentPassword !== undefined
            ? password === currentPassword
            : currentPasswordHash
              ? await verifySheetProtectionPasswordHash(password, currentPasswordHash)
              : false;
        if (!matches) {
          await showMessage({
            title: strings.protection.unprotectSheet,
            message: strings.ribbonMenu.workbookIncorrectPassword,
            okLabel: strings.pageSetup.ok,
          });
          instance.host.focus();
          return;
        }
      }
      instance.setSheetProtected(false);
      instance.host.focus();
      return;
    }
    if (action === 'lock-cell' || action === 'unlock-cell') {
      await buildCellFormatAction(instance, {})(action);
      return;
    }
    if (action === 'protect-workbook' || action === 'unprotect-workbook') {
      setWorkbookStructureProtected(instance.store, action === 'protect-workbook');
      instance.host.focus();
      return;
    }
    if (action === 'allow-edit-ranges' || action === 'clear-allowed-edit-ranges') {
      const strings = instance.i18n.strings;
      if (action === 'allow-edit-ranges') {
        const selection = normalizedSelectionRange(instance);
        const sheetName = instance.workbook.sheetName(selection.sheet);
        const pivotStrings = strings.pivotTableDialog as typeof strings.pivotTableDialog & {
          rangePickerSelect: string;
        };
        const result = await showAllowEditRangeDialog({
          strings: {
            title: strings.ribbonMenu.allowEditRangesDialogTitle,
            range: strings.ribbonMenu.allowEditRangesDialogRange,
            invalid: strings.ribbonMenu.allowEditRangesDialogInvalid,
            rangePickerLabel: pivotStrings.rangePickerSelect,
            ok: strings.pageSetup.ok,
            cancel: strings.pageSetup.cancel,
          },
          initialRange: formatA1Range(selection),
          pickRange: () => formatA1Range(normalizedSelectionRange(instance)),
          validateRange: (value) => parseA1Range(value, selection.sheet, sheetName) !== null,
          subscribeToRangeChanges: (listener) => instance.store.subscribe(listener),
        });
        if (result === null) {
          instance.host.focus();
          return;
        }
        const range = parseA1Range(result, selection.sheet, sheetName) ?? selection;
        const rangeText = formatA1Range(range);
        addAllowedEditRange(instance.store, range, { title: rangeText });
        const report = {
          title: strings.ribbonMenu.allowEditRangesDialogTitle,
          items: [
            {
              severity: 'info' as const,
              label: strings.ribbonMenu.allowEditRangesCommand,
              detail: strings.ribbonMenu.allowedEditRangeAddedStatus.replace('{range}', rangeText),
            },
          ],
        };
        await showReport({
          title: report.title,
          items: report.items,
          ...reportDialogLabels(strings),
        });
        instance.host.focus();
        return;
      }
      const report = executeRibbonProtectionAction({
        store: instance.store,
        action: 'clear-allowed-edit-ranges',
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
        ...reportDialogLabels(strings),
      });
    }
  };

const updateProtectMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateProtectMenu'] =>
  (menu) => {
    const state = instance.store.getState();
    const sheetProtected = state.protection.protectedSheets.has(state.data.sheetIndex);
    const workbookProtected = isWorkbookStructureProtected(state);
    const hasAllowedRanges = state.protection.allowedEditRanges.length > 0;
    const strings = instance.i18n.strings.ribbonMenu;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-protect-action]')) {
      const action = button.dataset.protectAction;
      const disabled =
        (action === 'protect-sheet' && sheetProtected) ||
        (action === 'unprotect-sheet' && !sheetProtected) ||
        (action === 'protect-workbook' && workbookProtected) ||
        (action === 'unprotect-workbook' && !workbookProtected) ||
        (action === 'clear-allowed-edit-ranges' && !hasAllowedRanges);
      const reason =
        action === 'protect-sheet'
          ? strings.protectSheetAlreadyProtected
          : action === 'unprotect-sheet'
            ? strings.unprotectSheetRequiresProtected
            : action === 'protect-workbook'
              ? strings.protectWorkbookAlreadyProtected
              : action === 'unprotect-workbook'
                ? strings.unprotectWorkbookRequiresProtected
                : action === 'clear-allowed-edit-ranges'
                  ? strings.allowEditRangesClearRequiresAny
                  : undefined;
      setMenuControlDisabled(button, disabled, reason);
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
    const range = normalizedSelectionRange(instance);
    const previewRows: string[] = [];
    const cells = instance.store.getState().data.cells;
    for (let row = range.r0; row <= Math.min(range.r1, range.r0 + 6); row += 1) {
      const value = cells.get(`${range.sheet}:${row}:${range.c0}`)?.value;
      if (value?.kind === 'text') previewRows.push(value.value);
    }
    const textToColumnsStrings = strings.ribbonMenu as typeof strings.ribbonMenu & {
      textToColumnsDataType: string;
      textToColumnsDelimited: string;
      textToColumnsFixedWidth: string;
      textToColumnsFixedWidthUnavailable: string;
      textToColumnsOther: string;
      textToColumnsPreview: string;
    };
    const result = await showTextToColumnsDialog({
      strings: {
        title: strings.ribbonMenu.textToColumnsDialogTitle,
        dataType: textToColumnsStrings.textToColumnsDataType,
        delimited: textToColumnsStrings.textToColumnsDelimited,
        fixedWidth: textToColumnsStrings.textToColumnsFixedWidth,
        fixedWidthUnavailable: textToColumnsStrings.textToColumnsFixedWidthUnavailable,
        delimiters: strings.ribbonMenu.textToColumnsDialogDelimiters,
        tab: strings.ribbonMenu.textToColumnsTab,
        semicolon: strings.ribbonMenu.textToColumnsSemicolon,
        comma: strings.ribbonMenu.textToColumnsComma,
        space: strings.ribbonMenu.textToColumnsSpace,
        other: textToColumnsStrings.textToColumnsOther,
        treatConsecutive: strings.ribbonMenu.textToColumnsTreatConsecutive,
        preview: textToColumnsStrings.textToColumnsPreview,
        noDelimited: strings.ribbonMenu.textToColumnsNoDelimited,
        ok: strings.hyperlinkDialog.ok,
        cancel: strings.hyperlinkDialog.cancel,
      },
      initialDelimiters: [','],
      previewRows,
    });
    if (result === null) {
      instance.host.focus();
      return;
    }
    const selected = normalizedSelectionRange(instance);
    instance.history.begin();
    try {
      textToColumns(
        instance.store.getState(),
        instance.store,
        instance.workbook,
        selected,
        result.delimiters,
        { collapseConsecutiveDelimiters: result.collapseConsecutiveDelimiters },
      );
      mutators.replaceCells(instance.store, instance.workbook.cells(selected.sheet));
    } finally {
      instance.history.end();
    }
    instance.host.focus();
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
  async () => {
    const strings = instance.i18n.strings;
    await showInstanceReport(instance, strings.ribbonMenu.recommendedCharts, [
      {
        severity: 'info',
        label: strings.ribbon.chart,
        detail: strings.workbookObjects.compatibilityDetails.chartAuthoring,
      },
    ]);
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
    for (const [key, format] of state.format.formats) {
      if (!format.validation) continue;
      const addr = addrFromKey(key);
      if (!addr || !addrInRange(addr, range)) continue;
      const value = instance.workbook.getValue(addr);
      if (cellValueViolatesValidation(value, format.validation)) invalid.add(key);
    }
    mutators.setValidationCircles(instance.store, invalid);
    instance.host.focus();
  };

const updateDataValidationMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateDataValidationMenu'] =>
  (menu) => {
    const state = instance.store.getState();
    const range = normalizedSelectionRange(instance);
    let hasValidation = false;
    for (const [key, format] of state.format.formats) {
      if (!format.validation) continue;
      const addr = addrFromKey(key);
      if (addr && addrInRange(addr, range)) {
        hasValidation = true;
        break;
      }
    }
    const hasValidationCircles = state.errorIndicators.validationCircles.size > 0;
    const strings = instance.i18n.strings.ribbonMenu;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-validation-action]')) {
      const action = button.dataset.validationAction;
      const disabled =
        ((action === 'circle-invalid' || action === 'clear-rules') && !hasValidation) ||
        (action === 'clear-circles' && !hasValidationCircles);
      const reason =
        action === 'clear-circles'
          ? strings.validationClearCirclesRequiresAny
          : strings.validationRequiresRules;
      setMenuControlDisabled(button, disabled, reason);
    }
  };

const buildConditionalMenuAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyConditionalMenuAction'] =>
  async (action, panel) => {
    const strings = instance.i18n.strings;
    const conditionalMenuStrings = strings.conditionalMenu as typeof strings.conditionalMenu & {
      ok: string;
      cancel: string;
      formatWith: string;
      formatPreview: string;
      customFormat: string;
      customFormatTitle: string;
      customFillColor: string;
      customTextColor: string;
      customBold: string;
      customItalic: string;
      customUnderline: string;
      customStrike: string;
      formatLightRed: string;
      formatYellow: string;
      formatGreen: string;
      formatLightRedFill: string;
      formatRedText: string;
      formatRedBorder: string;
      formatRedFill: string;
      formatRedTextFill: string;
      invalidNumber: string;
      invalidText: string;
    };
    const conditionalDialogStrings = {
      ok: conditionalMenuStrings.ok,
      cancel: conditionalMenuStrings.cancel,
      formatWith: conditionalMenuStrings.formatWith,
      formatPreview: conditionalMenuStrings.formatPreview,
      customFormat: conditionalMenuStrings.customFormat,
      customFormatTitle: conditionalMenuStrings.customFormatTitle,
      customFillColor: conditionalMenuStrings.customFillColor,
      customTextColor: conditionalMenuStrings.customTextColor,
      customBold: conditionalMenuStrings.customBold,
      customItalic: conditionalMenuStrings.customItalic,
      customUnderline: conditionalMenuStrings.customUnderline,
      customStrike: conditionalMenuStrings.customStrike,
      formatLightRed: conditionalMenuStrings.formatLightRed,
      formatYellow: conditionalMenuStrings.formatYellow,
      formatGreen: conditionalMenuStrings.formatGreen,
      formatLightRedFill: conditionalMenuStrings.formatLightRedFill,
      formatRedText: conditionalMenuStrings.formatRedText,
      formatRedBorder: conditionalMenuStrings.formatRedBorder,
      formatRedFill: conditionalMenuStrings.formatRedFill,
      formatRedTextFill: conditionalMenuStrings.formatRedTextFill,
      invalidNumber: conditionalMenuStrings.invalidNumber,
      invalidText: conditionalMenuStrings.invalidText,
    };
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
        promptCfNumber: (spec) =>
          showConditionalFormatNumberDialog({
            ...spec,
            strings: conditionalDialogStrings,
          }),
        promptCfText: (spec) =>
          showConditionalFormatTextDialog({
            ...spec,
            strings: conditionalDialogStrings,
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

const currentPageThemeAction = (theme: string | undefined): 'light' | 'dark' | 'contrast' => {
  if (theme === 'dark' || theme === 'ink') return 'dark';
  if (theme === 'contrast') return 'contrast';
  return 'light';
};

const updatePageThemeMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updatePageThemeMenu'] =>
  (menu) => {
    const current = currentPageThemeAction(instance.store.getState().ui.theme);
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-page-theme-action]')) {
      const active = button.dataset.pageThemeAction === current;
      button.setAttribute('role', 'menuitemradio');
      button.setAttribute('aria-checked', String(active));
      button.classList.toggle('app__visual-tile--active', active);
    }
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

const updatePrintAreaMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updatePrintAreaMenu'] =>
  (menu) => {
    const sheet = instance.store.getState().data.sheetIndex;
    const hasPrintArea = !!getPageSetup(instance.store.getState(), sheet).printArea?.trim();
    const disabledReason = instance.i18n.strings.ribbonMenu.printAreaRequiresExisting;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-print-area-action]')) {
      const action = button.dataset.printAreaAction;
      const disabled = (action === 'add' || action === 'clear') && !hasPrintArea;
      setMenuControlDisabled(button, disabled, disabledReason);
    }
  };

const buildPivotTableAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyPivotTableAction'] =>
  async (action) => {
    if (action === 'new-sheet' || action === 'existing-sheet') {
      (instance.openPivotTableDialog as (opts?: { placement?: 'new' | 'existing' }) => void)({
        placement: action === 'new-sheet' ? 'new' : 'existing',
      });
      return;
    }
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
        ...reportDialogLabels(strings),
      });
      return;
    }
    instance.host.focus();
  };

const updateCellInsertMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateCellInsertMenu'] =>
  (menu) => {
    const structureProtected = isWorkbookStructureProtected(instance.store.getState());
    const disabledReason = instance.i18n.strings.ribbonMenu.workbookStructureProtectedBlocked;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-cell-insert]')) {
      const disabled = button.dataset.cellInsert === 'sheet' && structureProtected;
      setMenuControlDisabled(button, disabled, disabledReason);
    }
  };

const updateCellDeleteMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateCellDeleteMenu'] =>
  (menu) => {
    const state = instance.store.getState();
    const structureProtected = isWorkbookStructureProtected(state);
    const canRemoveSheet = instance.workbook.capabilities.sheetMutate;
    const strings = instance.i18n.strings.ribbonMenu;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-cell-delete]')) {
      let disabledReason: string | undefined;
      if (button.dataset.cellDelete === 'sheet') {
        if (structureProtected) disabledReason = strings.workbookStructureProtectedBlocked;
        else if (instance.workbook.sheetCount <= 1) {
          disabledReason = strings.sheetDeleteRequiresAnotherSheet;
        } else if (!canRemoveSheet) disabledReason = strings.sheetMutationUnavailable;
      }
      const disabled =
        button.dataset.cellDelete === 'sheet' &&
        (structureProtected || instance.workbook.sheetCount <= 1 || !canRemoveSheet);
      setMenuControlDisabled(button, disabled, disabledReason);
    }
  };

const buildPageBreakAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyPageBreakAction'] =>
  (action) => {
    const range = normalizedSelectionRange(instance);
    if (action === 'insert') {
      if (range.r0 > 0)
        insertManualPageBreak(instance.store, range.sheet, 'row', range.r0, instance.history);
      if (range.c0 > 0)
        insertManualPageBreak(instance.store, range.sheet, 'col', range.c0, instance.history);
    } else if (action === 'insert-row')
      insertManualPageBreak(instance.store, range.sheet, 'row', range.r0, instance.history);
    else if (action === 'insert-col')
      insertManualPageBreak(instance.store, range.sheet, 'col', range.c0, instance.history);
    else if (action === 'remove') {
      removeManualPageBreak(instance.store, range.sheet, 'row', range.r0, instance.history);
      removeManualPageBreak(instance.store, range.sheet, 'col', range.c0, instance.history);
    } else if (action === 'remove-row')
      removeManualPageBreak(instance.store, range.sheet, 'row', range.r0, instance.history);
    else if (action === 'remove-col')
      removeManualPageBreak(instance.store, range.sheet, 'col', range.c0, instance.history);
    else if (action === 'reset-all')
      resetManualPageBreaks(instance.store, range.sheet, instance.history);
    instance.host.focus();
  };

const updatePageBreaksMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updatePageBreaksMenu'] =>
  (menu) => {
    const range = normalizedSelectionRange(instance);
    const setup = getPageSetup(instance.store.getState(), range.sheet);
    const rowBreaks = new Set(setup.manualPageBreakRows ?? []);
    const colBreaks = new Set(setup.manualPageBreakCols ?? []);
    const hasAnyBreak = rowBreaks.size > 0 || colBreaks.size > 0;
    const canInsertAtSelection = range.r0 > 0 || range.c0 > 0;
    const hasBreakAtSelection = rowBreaks.has(range.r0) || colBreaks.has(range.c0);
    const strings = instance.i18n.strings.ribbonMenu;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-page-break-action]')) {
      const action = button.dataset.pageBreakAction;
      const disabled =
        (action === 'insert' && !canInsertAtSelection) ||
        (action === 'insert-row' && range.r0 <= 0) ||
        (action === 'insert-col' && range.c0 <= 0) ||
        (action === 'remove' && !hasBreakAtSelection) ||
        (action === 'remove-row' && !rowBreaks.has(range.r0)) ||
        (action === 'remove-col' && !colBreaks.has(range.c0)) ||
        (action === 'reset-all' && !hasAnyBreak);
      const reason = action?.startsWith('insert')
        ? strings.pageBreakInsertRequiresSelection
        : action === 'reset-all'
          ? strings.pageBreakResetRequiresAny
          : strings.pageBreakRemoveRequiresBreak;
      setMenuControlDisabled(button, disabled, reason);
    }
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
    const picked = await pickImageFileDataUrl();
    if (!picked) {
      instance.host.focus();
      return;
    }
    setSheetBackgroundImage(instance.store, sheet, picked.src, instance.history);
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
        thenByLabel: strings.sortThenBy,
        noThenByLabel: strings.sortNoThenBy,
        orderLabel: strings.sortDialogOrder,
        headerLabel: strings.sortDialogHeader,
        addLevelLabel: strings.sortAddLevel,
        deleteLevelLabel: strings.sortDeleteLevel,
        copyLevelLabel: strings.sortCopyLevel,
        levelUnavailableLabel: strings.sortLevelUnavailable,
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
    if (action === 'filter-advanced') {
      const strings = instance.i18n.strings
        .ribbonMenu as typeof instance.i18n.strings.ribbonMenu & {
        advancedFilterInvalidRange: string;
        advancedFilterRangePicker: string;
      };
      const range = inferAutoFilterRange(instance.store.getState());
      const sheetName = instance.workbook.sheetName(range.sheet);
      const invalidRange = strings.advancedFilterInvalidRange;
      const validateRange = (value: string): string | null =>
        parseA1Range(value.trim(), range.sheet, sheetName) ? null : invalidRange;
      const validateAddress = (value: string): string | null =>
        parseA1Range(value.trim(), range.sheet, sheetName) ? null : invalidRange;
      const result = await showAdvancedFilterDialog({
        title: strings.advancedFilterDialogTitle,
        listRangeLabel: strings.advancedFilterListRange,
        criteriaRangeLabel: strings.advancedFilterCriteriaRange,
        copyToLabel: strings.advancedFilterCopyTo,
        uniqueOnlyLabel: strings.advancedFilterUniqueOnly,
        initialListRange: formatA1Range(range),
        okLabel: strings.sortDialogApply,
        cancelLabel: strings.sortDialogCancel,
        rangePickerLabel: strings.advancedFilterRangePicker,
        pickRange: () => formatA1Range(normalizedSelectionRange(instance)),
        pickAddress: () => {
          const active = instance.store.getState().selection.active;
          return formatA1Range({
            sheet: active.sheet,
            r0: active.row,
            c0: active.col,
            r1: active.row,
            c1: active.col,
          });
        },
        subscribeToRangeChanges: (listener) => instance.store.subscribe(listener),
        validateRange,
        validateAddress,
      });
      if (!result) {
        instance.host.focus();
        return;
      }
      const listRange = parseA1Range(result.listRange, range.sheet, sheetName);
      const criteriaRange = parseA1Range(result.criteriaRange, range.sheet, sheetName);
      const copyToRange = result.copyTo
        ? parseA1Range(result.copyTo, range.sheet, sheetName)
        : null;
      if (!listRange || !criteriaRange || (result.copyTo && !copyToRange)) {
        instance.host.focus();
        return;
      }
      instance.history.begin();
      try {
        if (copyToRange) {
          const copied = copyAdvancedFilterResult(
            instance.store.getState(),
            instance.store,
            listRange,
            criteriaRange,
            { sheet: copyToRange.sheet, row: copyToRange.r0, col: copyToRange.c0 },
            { uniqueOnly: result.uniqueOnly },
            instance.workbook,
          );
          mutators.replaceCells(instance.store, instance.workbook.cells(listRange.sheet));
          await showInstanceReport(instance, strings.advancedFilterDialogTitle, [
            {
              severity: 'info',
              label: strings.filterAdvanced,
              detail: strings.advancedFilterCopiedStatus.replace('{count}', String(copied)),
            },
          ]);
        } else {
          applyAdvancedFilter(instance.store.getState(), instance.store, listRange, criteriaRange);
        }
      } finally {
        instance.history.end();
      }
      instance.host.focus();
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

const updateSortMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateSortMenu'] =>
  (menu) => {
    const state = instance.store.getState();
    const hasFilterRange = state.ui.filterRange !== null;
    const hasFilterCriteria = state.ui.filterCriteria.length > 0;
    const strings = instance.i18n.strings.ribbonMenu;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-sort]')) {
      const action = button.dataset.sort;
      const disabled =
        (action === 'filter-clear' && !hasFilterRange) ||
        (action === 'filter-reapply' && !hasFilterCriteria);
      const reason =
        action === 'filter-clear'
          ? strings.filterClearRequiresRange
          : action === 'filter-reapply'
            ? strings.filterReapplyRequiresCriteria
            : undefined;
      setMenuControlDisabled(button, disabled, reason);
      if (action === 'filter') {
        button.setAttribute('aria-pressed', String(hasFilterRange));
        button.classList.toggle('app__menu-item--active', hasFilterRange);
      }
    }
  };

const buildTableStyleAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['createTableFromSelection'] =>
  (style = 'medium', color, variant = 'banded') => {
    void showCreateTableDialog(instance, { style, color, variant });
  };

const updateTableStylesMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateTableStylesMenu'] =>
  (menu) => {
    const state = instance.store.getState();
    const active = state.selection.active;
    const activeTable = tableOverlayAt(state, active.sheet, active.row, active.col);
    const activeTableVariant = activeTable
      ? tableVariantFromOptions({
          banded: activeTable.banded,
          firstCol: activeTable.firstCol,
        })
      : null;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-table-style]')) {
      const activeStyle =
        activeTable != null &&
        button.dataset.tableStyle === activeTable.style &&
        button.dataset.tableColor === (activeTable.color ?? DEFAULT_TABLE_COLOR) &&
        button.dataset.tableVariant === activeTableVariant;
      button.setAttribute('role', 'menuitemradio');
      button.setAttribute('aria-checked', String(activeStyle));
      button.classList.toggle('app__menu-item--active', activeStyle);
    }

    const activePivot = findPivotTableAtCell(
      instance.workbook as unknown as Parameters<typeof findPivotTableAtCell>[0],
      active,
    );
    const activePivotStyle = activePivot
      ? pivotTableStyleAssignment(state, activePivot.sheetIndex, activePivot.pivotIndex)?.styleId
      : null;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-pivot-table-style]')) {
      const activeStyle = button.dataset.pivotTableStyle === activePivotStyle;
      button.setAttribute('role', 'menuitemradio');
      button.setAttribute('aria-checked', String(activeStyle));
      button.classList.toggle('app__menu-item--active', activeStyle);
    }
  };

export const showCreateTableDialog = async (
  instance: SpreadsheetInstance,
  options: {
    style?: string;
    color?: string;
    variant?: CustomTableStyle['variant'];
  } = {},
): Promise<void> => {
  const selection = normalizedSelectionRange(instance);
  const sheetName = instance.workbook.sheetName(selection.sheet);
  const pivotDialogStrings = instance.i18n.strings
    .pivotTableDialog as typeof instance.i18n.strings.pivotTableDialog & {
    rangePickerSelect: string;
    createTableTitle: string;
    createTableRangeLabel: string;
    createTableHeadersLabel: string;
    createTableInvalidRange: string;
  };
  const parsedRange = (value: string): Range | null =>
    parseA1Range(value, selection.sheet, sheetName) as Range | null;
  const result = await showFormatAsTableDialog({
    title: pivotDialogStrings.createTableTitle,
    rangeLabel: pivotDialogStrings.createTableRangeLabel,
    headersLabel: pivotDialogStrings.createTableHeadersLabel,
    initialRange: formatA1Range(selection),
    initialHasHeaders: inferTableHasHeaders(instance.workbook, selection),
    okLabel: pivotDialogStrings.ok,
    cancelLabel: pivotDialogStrings.cancel,
    rangePickerLabel: pivotDialogStrings.rangePickerSelect,
    pickRange: () => formatA1Range(normalizedSelectionRange(instance)),
    subscribeToRangeChanges: (listener) => instance.store.subscribe(listener),
    validateRange: (value) =>
      parsedRange(value) ? null : pivotDialogStrings.createTableInvalidRange,
  });
  if (!result) {
    instance.host.focus();
    return;
  }
  const range = parsedRange(result.range);
  if (!range) return;
  recordTablesChange(instance.history, instance.store, () => {
    formatAsTableByStyleId(
      instance.store,
      range,
      options.style ?? 'medium',
      options.color,
      options.variant ?? 'banded',
      {
        showHeader: result.hasHeaders,
      },
    );
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

const updateCellStylesMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateCellStylesMenu'] =>
  (menu) => {
    const state = instance.store.getState();
    const current = formatWithPending(state, state.selection.active)?.cellStyle ?? null;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-cell-style]')) {
      const active = button.dataset.cellStyle === current;
      button.setAttribute('role', 'menuitemradio');
      button.setAttribute('aria-checked', String(active));
      button.classList.toggle('app__menu-item--active', active);
    }
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

const updateCurrencyMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateCurrencyMenu'] =>
  (menu) => {
    const state = instance.store.getState();
    const current = formatWithPending(state, state.selection.active)?.numFmt;
    const activeSymbol = current?.kind === 'currency' ? current.symbol : null;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-currency-preset]')) {
      const active = button.dataset.currencyPreset === activeSymbol;
      button.setAttribute('role', 'menuitemradio');
      button.setAttribute('aria-checked', String(active));
      button.classList.toggle('app__menu-item--active', active);
    }
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
    ...reportDialogLabels(strings),
  });
};

const buildScriptAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyScriptAction'] =>
  async (action) => {
    const strings = instance.i18n.strings;
    const raw =
      action === 'custom'
        ? await showScriptCommandDialog({
            title: strings.ribbonMenu.scriptDialogTitle,
            label: strings.ribbonMenu.scriptDialogCommand,
            options: [
              { value: 'uppercase', label: strings.ribbonMenu.scriptCommandUppercase },
              { value: 'lowercase', label: strings.ribbonMenu.scriptCommandLowercase },
              { value: 'trim', label: strings.ribbonMenu.scriptCommandTrim },
              { value: 'clear', label: strings.ribbonMenu.scriptCommandClear },
            ],
            initial: 'uppercase',
            okLabel: strings.ribbonMenu.scriptDialogRun,
            cancelLabel: strings.hyperlinkDialog.cancel,
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
    const ribbonMenu = strings.ribbonMenu as typeof strings.ribbonMenu & { pictureStock: string };
    const label =
      action === 'stock'
        ? ribbonMenu.pictureStock
        : action === 'online'
          ? strings.ribbonMenu.pictureOnline
          : strings.ribbonMenu.pictureThisDevice;
    const compatibilityDetails = strings.workbookObjects
      .compatibilityDetails as typeof strings.workbookObjects.compatibilityDetails & {
      mediaPickers?: string;
    };
    await showInstanceReport(instance, strings.ribbon.pictures, [
      {
        severity: 'info',
        label,
        detail: compatibilityDetails.mediaPickers ?? compatibilityDetails.chartsDrawings,
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

const updateArrangeMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateArrangeMenu'] =>
  (menu) => {
    const activeId = activeIllustrationId(instance);
    const state = instance.store.getState();
    const sheet = state.data.sheetIndex;
    const visible = state.illustrations.illustrations.filter((item) => item.sheet === sheet);
    const activeIndex = activeId ? visible.findIndex((candidate) => candidate.id === activeId) : -1;
    const hasTarget = activeIndex >= 0;
    const atBack = !hasTarget || activeIndex === 0;
    const atFront = !hasTarget || activeIndex === visible.length - 1;
    const strings = instance.i18n.strings.ribbonMenu;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-arrange-action]')) {
      const action = button.dataset.arrangeAction;
      const disabled =
        (action === 'bring-forward' && atFront) ||
        (action === 'bring-front' && atFront) ||
        (action === 'send-backward' && atBack) ||
        (action === 'send-back' && atBack) ||
        (!hasTarget && action !== 'selection-pane');
      const reason = !hasTarget
        ? strings.arrangeRequiresObject
        : action === 'bring-forward' || action === 'bring-front'
          ? strings.arrangeAtFront
          : action === 'send-backward' || action === 'send-back'
            ? strings.arrangeAtBack
            : undefined;
      setMenuControlDisabled(button, disabled, reason);
    }
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
      const detail =
        result.imported > 0
          ? strings.ribbonMenu.cellStyleMergeImported.replace('{count}', String(result.imported))
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
  const focusSheet = opts.focusSheet ?? ((): void => instance.host.focus());

  const base: DynamicDropdownsCtx = {
    getInst: () => instance,
    updateCalcOptionsMenu: updateCalcOptionsMenu(instance),
    updateCellDeleteMenu: updateCellDeleteMenu(instance),
    updateCellInsertMenu: updateCellInsertMenu(instance),
    updateClearMenu: updateClearMenu(instance),
    updateClearArrowsMenu: updateClearArrowsMenu(instance),
    updateCurrencyMenu: updateCurrencyMenu(instance),
    updateDataValidationMenu: updateDataValidationMenu(instance),
    updateDefinedNamesMenu: updateDefinedNamesMenu(instance),
    updateErrorCheckingMenu: updateErrorCheckingMenu(instance),
    updateFormatCellsMenu: updateFormatCellsMenu(instance),
    updateLinksMenu: updateLinksMenu(instance),
    updatePageBreaksMenu: updatePageBreaksMenu(instance),
    updatePrintAreaMenu: updatePrintAreaMenu(instance),
    updateProtectMenu: updateProtectMenu(instance),
    updatePageThemeMenu: updatePageThemeMenu(instance),
    updateReviewCommentsMenu: updateReviewCommentsMenu(instance),
    updateSortMenu: updateSortMenu(instance),
    updateWatchMenu: updateWatchMenu(instance),
    closeBorderMenu: noop,
    closeFreezeMenu: noop,
    closePrintAreaMenu: noop,
    closeSymbolMenu: noop,
    getConditionalMenu: () => document.getElementById('menu-conditional') as HTMLElement | null,
    focusSheet,

    // Pure / instance-derivable defaults.
    updateArrangeMenu: updateArrangeMenu(instance),
    applyRibbonPasteAction: buildPasteAction(instance),
    updatePasteMenu: updatePasteMenu(instance),
    applyFillSeries: buildFillSeries(instance),
    updateFillMenu: updateFillMenu(instance),
    applyFillDirection: buildFillDirection(instance),
    applyClearAction: buildClearAction(instance),
    applyUnderlineAction: buildUnderlineAction(instance),
    applyMergeAction: buildMergeAction(instance),
    applyFreezeAction: buildFreezeAction(instance),
    updateFreezeMenu: updateFreezeMenu(instance),
    applyTextOrientationAction: buildTextOrientation(instance),
    updateTextOrientationMenu: updateTextOrientationMenu(instance),
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
    applyCellFormatAction: buildCellFormatAction(instance, opts),
    applyPageBreakAction: buildPageBreakAction(instance),
    applySheetBackgroundAction: buildSheetBackgroundAction(instance),
    applyPrintAreaAction: buildPrintAreaAction(instance),
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
    updateTableStylesMenu: updateTableStylesMenu(instance),
    applyPivotTableStyleFromRibbon: buildPivotTableStyleAction(instance),
    applyCellStyleFromRibbon: buildCellStyleAction(instance),
    updateCellStylesMenu: updateCellStylesMenu(instance),
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
