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
  applyConditionalPresetAction,
  autoSum,
  type ConditionalPresetAction,
  clearVisualFormat,
  clearWatchedCells,
  type DynamicDropdownsCtx,
  dispatchHostClipboard,
  fillRange,
  handlePasteAction,
  inferFillSeriesDirection,
  mutators,
  type PasteAction,
  type Range,
  type RibbonFillSeriesMode,
  recordFormatChange,
  recordWatchesChange,
  type SpreadsheetInstance,
  setRotation,
  type ThemeName,
  unwatchCell,
  watchRange,
} from '@libraz/formulon-cell';

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
const asyncNoop = async (): Promise<void> => undefined;

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
  (mode) => {
    if (!mode) return; // host-driven dialog flow is opt-in via overrides
    const range = normalizedSelectionRange(instance);
    const direction = inferFillSeriesDirection(range);
    let src: Range = range;
    if (direction === 'down') src = { ...range, r1: range.r0 };
    else if (direction === 'up') src = { ...range, r0: range.r1 };
    else if (direction === 'right') src = { ...range, c1: range.c0 };
    else src = { ...range, c0: range.c1 };
    if (src.r0 === range.r0 && src.r1 === range.r1 && src.c0 === range.c0 && src.c1 === range.c1) {
      return;
    }
    const dateUnit: RibbonFillSeriesMode | undefined =
      mode === 'days' || mode === 'weekdays' || mode === 'months' || mode === 'years'
        ? mode
        : undefined;
    instance.history.begin();
    try {
      recordFormatChange(instance.history, instance.store, () => {
        fillRange(instance.store.getState(), instance.workbook, src, range, {
          copyOnly: mode === 'copy',
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
    const range = normalizedSelectionRange(instance);
    if (action === 'contents') {
      instance.history.begin();
      try {
        for (let row = range.r0; row <= range.r1; row += 1) {
          for (let col = range.c0; col <= range.c1; col += 1) {
            instance.workbook.setBlank({ sheet: range.sheet, row, col });
          }
        }
      } finally {
        instance.history.end();
      }
      instance.host.focus();
      return;
    }
    if (action === 'formats') {
      recordFormatChange(instance.history, instance.store, () => {
        clearVisualFormat(instance.store.getState(), instance.store);
      });
      instance.host.focus();
      return;
    }
    if (action === 'conditional') {
      mutators.clearConditionalRulesInRange(instance.store, range);
      instance.host.focus();
      return;
    }
    // `comments`, `hyperlinks`, `remove-hyperlinks` need host comment/hyperlink
    // dialogs and history wrappers; the host opts in via `overrides`.
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
  (action) => {
    if (action === 'find') {
      instance.openFindReplace('find');
      return;
    }
    if (action === 'replace') {
      instance.openFindReplace('replace');
      return;
    }
    if (action === 'go-to') {
      instance.openGoTo();
      return;
    }
    if (action === 'go-to-special') {
      instance.openGoToSpecial();
    }
    // `conditional-format`, `formulas`, `constants`, `data-validation`,
    // `comments` walk the workbook to build a selection — host can override
    // via opts when it wires `findMatchingCells` / `listComments`.
  };

const buildFormulaAuditAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyFormulaAuditAction'] =>
  (action) => {
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
    // `clear-precedents`, `clear-dependents`, `error-checking`, `trace-error`
    // and `ignore-error` involve trace-kind filtering / error dialogs the
    // host owns — opt-in via overrides.
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

const buildDataValidationAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyDataValidationAction'] =>
  (action) => {
    if (action === 'manage' || action === 'more' || action === 'open') {
      instance.openDataValidationDialog();
    }
  };

const buildConditionalMenuAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyConditionalMenuAction'] =>
  (action) => {
    if (action === 'manage' || action === 'new-rule') {
      instance.openConditionalDialog();
      return;
    }
    if (action === 'clear-selection' || action === 'clear-sheet' || action === 'clear') {
      const cfAction: ConditionalPresetAction =
        action === 'clear-sheet' ? 'clear-sheet' : 'clear-selection';
      applyConditionalPresetAction(instance.store, cfAction, normalizedSelectionRange(instance));
      instance.host.focus();
      return;
    }
    // `applyConditionalPresetAction` returns false for unknown ids, so the
    // ribbon's full action enum (e.g. `top-bottom-more`) safely falls through
    // without us having to enumerate every preset. The `*-more` /
    // `cell-greater` family open playground-side prompt dialogs and stay
    // host-overridable.
    if (
      applyConditionalPresetAction(
        instance.store,
        action as ConditionalPresetAction,
        normalizedSelectionRange(instance),
      )
    ) {
      instance.host.focus();
    }
  };

const buildUiTheme =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyUiTheme'] =>
  (theme) => {
    instance.setTheme(theme as ThemeName);
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
    applyTextOrientationAction: buildTextOrientation(instance),
    applyAutoSumFormula: buildAutoSumFormula(instance),
    applyFormulaAuditAction: buildFormulaAuditAction(instance),
    applyWatchAction: buildWatchAction(instance),
    applyCalcOptionAction: buildCalcOptionAction(instance),
    applyFindSelectAction: buildFindSelectAction(instance),
    applyDataValidationAction: buildDataValidationAction(instance),
    applyConditionalMenuAction: buildConditionalMenuAction(instance),
    applyUiTheme: buildUiTheme(instance),

    // Dialog / host-glue — host opts in via `overrides`. Defaults are no-op
    // so the click delegator returns true (event consumed) and the open
    // menu still closes, instead of falling through to the legacy fallback.
    applyPivotTableAction: asyncNoop,
    applyDefinedNameAction: asyncNoop,
    applyLinksAction: asyncNoop,
    applyCellInsertAction: asyncNoop,
    applyCellDeleteAction: asyncNoop,
    applyCellFormatAction: asyncNoop,
    applyPageBreakAction: noop,
    applySheetBackgroundAction: asyncNoop,
    applyPrintTitlesAction: noop,
    applySortMenuAction: noop,
    applyReviewCommentAction: noop,
    applyProtectAction: asyncNoop,
    createRecommendedChartFromSelection: asyncNoop,
    createChartFromSelection: noop,
    chartKindFromAction: () => 'column',
    insertPictureFromRibbon: asyncNoop,
    insertShapeFromRibbon: noop,
    insertScreenshotFromRibbon: noop,
    applyScriptAction: asyncNoop,
    applyPdfAction: asyncNoop,
    createTableFromSelection: asyncNoop,
    openTableStyleFooterAction: asyncNoop,
    applyCellStyleFromRibbon: noop,
    openCellStyleFooterAction: asyncNoop,
    applyCurrencyPreset: noop,
    openCurrencyFooterAction: noop,
    splitTextToColumns: asyncNoop,
    splitTextToColumnsCustom: asyncNoop,
    applyAddInAction: asyncNoop,
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
