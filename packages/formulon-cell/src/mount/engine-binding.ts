import type { CellRect } from '@libraz/formulon';
import { setDataTableGuardSink } from '../commands/data-table-guard.js';
import type { History } from '../commands/history.js';
import { interactionControllerFor } from '../commands/interaction-controller.js';
import { localizedDecisionReason } from '../commands/permission-reasons.js';
import { setPivotGuardSink } from '../commands/pivot-guard.js';
import { finishPivotReportUpdate } from '../commands/pivot-report-update.js';
import { setProtectionAlertSink } from '../commands/protection-alert.js';
import { addrKey } from '../engine/address.js';
import { hydrateThreadedCommentsFromEngine } from '../engine/annotations-sync.js';
import { hydrateAutoFilterFromEngine } from '../engine/auto-filter-sync.js';
import { hydrateCalcSettingsFromEngine } from '../engine/calc-settings-sync.js';
import { hydrateCellFormatsFromEngine } from '../engine/cell-format-sync.js';
import { hydrateConditionalRulesFromEngine } from '../engine/cf-sync.js';
import {
  hydrateChartsFromEngine,
  refreshChartsAfterRecalc,
  reloadChartsFromEngine,
} from '../engine/charts-sync.js';
import { editSeedForAddr } from '../engine/edit-seed.js';
import { hydrateCommentsAndHyperlinksFromEngine } from '../engine/format-sync.js';
import { hydrateImagesFromEngine, reloadImagesFromEngine } from '../engine/images-sync.js';
import { hydrateLayoutFromEngine } from '../engine/layout-sync.js';
import { hydrateMergesFromEngine } from '../engine/merges-sync.js';
import { hydratePageSetupFromEngine } from '../engine/print-sync.js';
import {
  hydrateAxisLocksFromEngine,
  hydrateProtectedRangesFromEngine,
  hydrateProtectionFromEngine,
} from '../engine/protection-sync.js';
import { refreshSheetPrFromEngine } from '../engine/sheet-pr-sync.js';
import { hydrateSlicersFromEngine } from '../engine/slicer-sync.js';
import { hydrateSparklinesFromEngine } from '../engine/sparkline-sync.js';
import { hydrateTableOverlaysFromEngine } from '../engine/table-sync.js';
import type { Addr, CellValue } from '../engine/types.js';
import { hydrateValidationsFromEngine } from '../engine/validation-sync.js';
import { hydrateViewStateFromEngine } from '../engine/view-state-sync.js';
import { applySheetStructureEvent } from '../engine/whole-sheet-font-sync.js';
import type {
  ChangeEvent,
  FacadeChange,
  FacadeChangeKind,
  WorkbookHandle,
  WriteRefusal,
} from '../engine/workbook-handle.js';
import type { CellChangeEvent, SpreadsheetEmitter } from '../events.js';
import type { resolveFlags } from '../extensions/features.js';
import type { ExtensionHandle } from '../extensions/index.js';
import { isMacHost } from '../extensions/ui-options.js';
import type { FormulaRegistry } from '../formula.js';
import type { Strings } from '../i18n/strings.js';
import { attachAutoFillOptions } from '../interact/auto-fill-options.js';
import { attachClipboard, runClipboardShortcutFrom } from '../interact/clipboard.js';
import { attachContextMenu } from '../interact/context-menu.js';
import type {
  ContextMenuInteractionController,
  ContextMenuOptions,
} from '../interact/context-menu-options.js';
import type { DialogEntry } from '../interact/dialog-entry.js';
import { InlineEditor } from '../interact/editor.js';
import { attachFindReplace } from '../interact/find-replace.js';
import type { FormatDialogHandle } from '../interact/format-dialog.js';
import type { HyperlinkDialogHandle } from '../interact/hyperlink-dialog.js';
import { followHyperlink, type HyperlinkFollowDeps } from '../interact/hyperlink-follow.js';
import { attachKeyboard } from '../interact/keyboard.js';
import type { NamedRangeDialogHandle } from '../interact/named-range-dialog.js';
import { attachPageBandEditor } from '../interact/page-band-editor.js';
import { attachPasteOptions } from '../interact/paste-options.js';
import { attachPasteSpecial } from '../interact/paste-special.js';
import { attachPointer } from '../interact/pointer.js';
import { attachQuickAnalysis } from '../interact/quick-analysis.js';
import type { RangeInsertTarget } from '../interact/range-insert.js';
import { attachSplitDrag } from '../interact/split-drag.js';
import { attachTabNavigationSession } from '../interact/tab-navigation-session.js';
import { attachTableTotals } from '../interact/table-totals.js';
import {
  attachValidationAlert,
  attachValidationList,
  attachValidationPrompt,
} from '../interact/validation.js';
import type { GridRenderer } from '../render/grid.js';
import type { SpreadsheetStore } from '../store/store.js';
import { mutators } from '../store/store.js';
import { resolveTheme } from '../theme/resolve.js';
import { showCommandRefusal } from '../toolbar/dialogs/command-refusal.js';
import { confirmMoveOverwrite } from '../toolbar/dialogs/move-overwrite-confirm.js';
import { hydrateActiveSheetFromEngine, hydrateWorkbookMetadataFromEngine } from './hydration.js';
import type { SheetTabsController } from './sheet-tabs-controller.js';

type FeatureFlags = ReturnType<typeof resolveFlags>;

type SliceHydrator = (
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
  range?: CellRect,
) => void;

/** The store slices each kind of facade write leaves stale, re-read from the engine for the sheet
 *  the write touched (the active sheet for a workbook-scoped write). */
const FACADE_HYDRATORS: Readonly<Record<FacadeChangeKind, readonly SliceHydrator[]>> = {
  format: [
    hydrateCellFormatsFromEngine,
    hydrateCommentsAndHyperlinksFromEngine,
    hydrateAxisLocksFromEngine,
  ],
  comments: [hydrateCommentsAndHyperlinksFromEngine],
  threadedComments: [hydrateThreadedCommentsFromEngine],
  // An engine insert/delete moves every range-anchored object, as the command path re-reads them.
  layout: [
    (wb, store, sheet) => mutators.replaceCells(store, wb.cells(sheet)),
    hydrateLayoutFromEngine,
    hydrateMergesFromEngine,
    hydrateCellFormatsFromEngine,
    hydrateCommentsAndHyperlinksFromEngine,
    hydrateThreadedCommentsFromEngine,
    hydrateConditionalRulesFromEngine,
    hydrateValidationsFromEngine,
    (wb, store) => hydrateTableOverlaysFromEngine(wb, store),
    hydrateAutoFilterFromEngine,
  ],
  view: [
    (wb, store) => hydrateViewStateFromEngine(wb, store, { followActiveSheet: false }),
    hydrateLayoutFromEngine,
    refreshSheetPrFromEngine,
  ],
  names: [],
  table: [(wb, store) => hydrateTableOverlaysFromEngine(wb, store), hydrateAutoFilterFromEngine],
  autoFilter: [hydrateAutoFilterFromEngine],
  drawing: [
    hydrateImagesFromEngine,
    hydrateChartsFromEngine,
    (wb, store) => hydrateSlicersFromEngine(wb, store),
  ],
  pivot: [(wb, store, sheet) => mutators.replaceCells(store, wb.cells(sheet))],
  sparklineGroups: [hydrateSparklinesFromEngine],
  print: [hydratePageSetupFromEngine],
  cf: [hydrateConditionalRulesFromEngine],
  // The calc mode and iterative settings live in the workbook alone; a new state object makes
  // the chrome re-read them.
  calc: [
    (wb, store) => hydrateCalcSettingsFromEngine(wb, store),
    (_wb, store) => store.setState((s) => ({ ...s })),
  ],
  validation: [hydrateValidationsFromEngine],
  protection: [
    (wb, store) => hydrateProtectionFromEngine(wb, store),
    hydrateProtectedRangesFromEngine,
  ],
};

/** Re-read the drawings a sheet rename, move, removal or copy changed. The
 *  store's sheet indexes follow the edit first, so each drawing is found on
 *  its new sheet under its id. */
export function reloadDrawingsAfterSheetEdit(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  event: Extract<
    ChangeEvent,
    { kind: 'sheet-rename' | 'sheet-remove' | 'sheet-move' | 'sheet-copy' }
  >,
): void {
  if (event.kind !== 'sheet-rename') applySheetStructureEvent(wb, event);
  reloadChartsFromEngine(wb, store);
  if (event.kind !== 'sheet-rename') reloadImagesFromEngine(wb, store);
}

export interface EngineBinding {
  editor: InlineEditor;
  pasteSpecialDialog: ReturnType<typeof attachPasteSpecial> | null;
  findReplace: ReturnType<typeof attachFindReplace> | null;
  validation: ReturnType<typeof attachValidationList> | null;
  validationPrompt: ReturnType<typeof attachValidationPrompt> | null;
  validationAlert: ReturnType<typeof attachValidationAlert> | null;
  quickAnalysis: ReturnType<typeof attachQuickAnalysis> | null;
  clipboardH: ReturnType<typeof attachClipboard> | null;
  contextMenu: ExtensionHandle | null;
  /** Follow a hyperlink target: select an in-document location or open an external URL. */
  followHyperlink: (target: string) => boolean;
  /** Make `sheet` the active sheet as a tab click does; false when it is hidden or out of range. */
  activateSheet: (sheet: number) => boolean;
  unbind: () => void;
}

interface AttachEngineBindingInput {
  contextMenuOptions?: ContextMenuOptions;
  emitter: SpreadsheetEmitter;
  flags: FeatureFlags;
  formulaRegistry: FormulaRegistry;
  getCommentDialog: () => { open(options?: { entry?: DialogEntry }): void } | null;
  getFormatDialog: () => FormatDialogHandle | null;
  getFormatPainter: () => { isActive(): boolean } | null;
  getFormulaBarEditor?: () => RangeInsertTarget | null;
  /** UI/data-format locale, the same one the renderer formats cells with. */
  getLocale: () => string;
  getGoToDialog: () => { open(mode?: 'go-to' | 'special'): void } | null;
  getHyperlinkDialog: () => Pick<HyperlinkDialogHandle, 'open'> | null;
  getNamedRangeDialog: () => Pick<NamedRangeDialogHandle, 'open'> | null;
  getPivotTableDialog: () => { open(): void } | null;
  getWorkbookObjects?: () => {
    openPivotFieldList(sheetIndex: number, pivotIndex: number): boolean;
    isPivotFieldListOpen(): boolean;
    hidePivotFieldList(): void;
  } | null;
  getSessionCharts: () => unknown | null;
  getSheetTabs: () => SheetTabsController | null;
  grid: HTMLElement;
  history: History;
  host: HTMLElement;
  renderer: GridRenderer;
  store: SpreadsheetStore;
  strings: Strings;
  tag: HTMLInputElement;
  updateChrome: () => void;
  wb: WorkbookHandle;
}

interface BufferedAtomicValue {
  readonly addr: Addr;
  readonly value: CellValue;
  readonly formula: string | null;
  readonly publicPayload: CellChangeEvent;
}

interface PendingAtomicBatch {
  readonly size: number;
  readonly slots: Array<BufferedAtomicValue | undefined>;
  count: number;
}

export const WB_REGISTRY_IDS = [
  'clipboard',
  'pasteSpecial',
  'quickAnalysis',
  'contextMenu',
  'findReplace',
  'validation',
] as const;

/** The dialog a refused store write belongs to, which titles its notice. */
function writeRefusalTitle(kind: WriteRefusal['kind'], strings: Strings): string {
  switch (kind) {
    // Tab colour is the only sheetPr field with a dialog.
    case 'sheetPr':
      return strings.sheetTabs.tabColor;
  }
}

export function attachEngineBinding(input: AttachEngineBindingInput): EngineBinding {
  const {
    contextMenuOptions,
    emitter,
    flags,
    formulaRegistry,
    getCommentDialog,
    getFormatDialog,
    getFormatPainter,
    getFormulaBarEditor,
    getGoToDialog,
    getHyperlinkDialog,
    getLocale,
    getNamedRangeDialog,
    getPivotTableDialog,
    getWorkbookObjects,
    getSessionCharts,
    getSheetTabs,
    grid,
    history,
    host,
    renderer,
    store,
    strings,
    tag,
    updateChrome,
    wb,
  } = input;

  const detachTabNavigationSession = attachTabNavigationSession(store);
  const pendingAtomic = new Map<number, PendingAtomicBatch>();

  let cellRefreshes = 0;
  const refreshCells = (): void => {
    cellRefreshes += 1;
    mutators.refreshActiveSheetCells(store, wb);
  };

  // A recalc changes cells besides the written one (dependents, spill ranges).
  // The UI commits reload the grid's cells themselves; a write made straight
  // through the workbook API does not, so reload once after the recalc unless
  // something already did.
  // `unbind` clears `bound`, which cancels a reload already queued: the host
  // may dispose the workbook in the same tick as the write.
  let bound = true;
  let recalcRefreshPending = false;
  const scheduleRecalcRefresh = (): void => {
    if (recalcRefreshPending) return;
    recalcRefreshPending = true;
    const seen = cellRefreshes;
    queueMicrotask(() => {
      recalcRefreshPending = false;
      if (bound && cellRefreshes === seen) refreshCells();
    });
  };

  const validationAlert = flags.validation
    ? attachValidationAlert({
        host,
        labels: {
          ok: strings.formatDialog.ok,
          stop: strings.formatDialog.validationErrorStop,
          warning: strings.formatDialog.validationErrorWarning,
          information: strings.formatDialog.validationErrorInfo,
          protectedCell: strings.protection.protectedCellEdit,
          reason: (refusal) => localizedDecisionReason(refusal, strings),
        },
      })
    : null;

  setPivotGuardSink(wb, (alert) => validationAlert?.show(alert), {
    edit: strings.pivotTableDialog.cannotChange,
    paste: strings.pivotTableDialog.cannotChangePaste,
  });
  setDataTableGuardSink(wb, (alert) => validationAlert?.show(alert), strings.dataTableDialog);
  setProtectionAlertSink(wb, (alert) => validationAlert?.show(alert), {
    filterRefused: strings.protection.filterRefused,
  });

  const activateSheet = (sheet: number): boolean => {
    if (
      store.getState().layout.hiddenSheets.has(sheet) ||
      store.getState().layout.veryHiddenSheets.has(sheet)
    )
      return false;
    const tabs = getSheetTabs();
    if (tabs) {
      tabs.switchSheet(sheet);
      return store.getState().data.sheetIndex === sheet;
    }
    mutators.setSheetIndex(store, sheet);
    if (store.getState().data.sheetIndex !== sheet) return false;
    hydrateActiveSheetFromEngine(wb, store);
    updateChrome();
    renderer.invalidate();
    return true;
  };
  const hyperlinkFollow: Pick<HyperlinkFollowDeps, 'activateSheet' | 'onInvalidTarget'> = {
    activateSheet,
    onInvalidTarget: () =>
      validationAlert?.show({ severity: 'stop', message: strings.goToDialog.invalidReference }),
  };

  const editor = new InlineEditor({
    host,
    grid,
    store,
    wb,
    getLabels: () => ({
      autocomplete: strings.autocomplete,
      argHelper: strings.argHelper,
    }),
    getCustomFunctions: () => formulaRegistry.list(),
    getCustomFunctionArgs: (name) => formulaRegistry.get(name)?.meta.args,
    onValidation: (outcome) => validationAlert?.show(outcome),
    onAfterCommit: refreshCells,
  });
  const detachPtr = attachPointer(
    grid,
    store,
    wb,
    refreshCells,
    history,
    () =>
      editor.isActive() && editor.isFormulaEdit()
        ? {
            isFormulaEdit: () => editor.isFormulaEdit(),
            editedCell: () => editor.editedCell(),
            insertRefAtCaret: (ref, options) => editor.insertRefAtCaret(ref, options),
          }
        : (getFormulaBarEditor?.() ?? null),
    () => ({
      locale: getLocale(),
      theme: resolveTheme(host),
      workbookDisplayFont: wb.workbookDisplayFont,
      workbook: wb,
    }),
    hyperlinkFollow,
    () => confirmMoveOverwrite(strings, wb),
    () => strings.resizeTip,
  );
  // Header / footer slots are only reachable in Page Layout view, but the
  // listener is cheap and stateless until one is clicked.
  const pageBandEditor = attachPageBandEditor({ grid, store, history });
  const detachSplitDrag = attachSplitDrag({ grid, store, getWb: () => wb });
  const autoFillOptions = attachAutoFillOptions({
    host: grid,
    store,
    wb,
    strings,
    history,
    onAfterCommit: refreshCells,
  });
  const pasteOptions = attachPasteOptions({
    host,
    grid,
    store,
    wb,
    strings,
    history,
    onAfterCommit: refreshCells,
  });
  // Clipboard must be attached before the keyboard router so the router can
  // forward Mod+C/X/V to the clipboard handle (browsers won't dispatch
  // copy/paste events on our non-editable, user-select:none host).
  const clipboardH = flags.clipboard
    ? attachClipboard({
        host,
        history,
        store,
        wb,
        onAfterCommit: refreshCells,
        onPasteOptions: pasteOptions.show,
        getStrings: () => strings,
      })
    : null;
  const pasteSpecialDialog =
    flags.pasteSpecial && clipboardH
      ? attachPasteSpecial({
          host,
          store,
          wb,
          strings,
          history,
          getSnapshot: () => clipboardH.getSnapshot(),
          onAfterCommit: refreshCells,
        })
      : null;
  const detachKey = flags.shortcuts
    ? attachKeyboard({
        host,
        store,
        wb,
        history,
        onBeginEdit: (seed, options) => editor.begin(seed, options),
        onClearActive: () => {
          refreshCells();
          updateChrome();
        },
        onAfterHistory: refreshCells,
        onGoTo: () => {
          const goToDialog = getGoToDialog();
          if (goToDialog) {
            goToDialog.open('go-to');
            return;
          }
          tag.focus();
          tag.select();
        },
        onSwitchSheet: (delta) => getSheetTabs()?.switchRelative(delta),
        onEditComment: () =>
          getCommentDialog()?.open({ entry: { origin: 'keyboard', commandId: 'insertComment' } }),
        onClipboardShortcut: clipboardH
          ? (kind) =>
              runClipboardShortcutFrom(clipboardH, kind, { origin: 'keyboard', commandId: kind })
          : undefined,
        onPasteSpecial: pasteSpecialDialog
          ? () =>
              pasteSpecialDialog.open(undefined, { origin: 'keyboard', commandId: 'pasteSpecial' })
          : undefined,
      })
    : (): void => {};
  const detachContextMenu = flags.contextMenu
    ? attachContextMenu({
        host,
        grid,
        store,
        wb,
        strings,
        options: contextMenuOptions,
        interactionController: interactionControllerFor(store) as ContextMenuInteractionController,
        history,
        onAfterCommit: refreshCells,
        onClipboardShortcut: clipboardH
          ? (kind) =>
              runClipboardShortcutFrom(clipboardH, kind, { origin: 'contextMenu', commandId: kind })
          : undefined,
        onFormatDialog: () =>
          getFormatDialog()?.open(undefined, {
            entry: { origin: 'contextMenu', commandId: 'formatCells' },
          }),
        // The dialogs attach and detach with their features, so availability is read as the menu opens.
        dialogAttached: (id) => {
          switch (id) {
            case 'formatCells':
              return getFormatDialog() !== null;
            case 'pasteSpecial':
              return pasteSpecialDialog !== null;
            case 'insertHyperlink':
              return getHyperlinkDialog() !== null;
            case 'insertComment':
              return getCommentDialog() !== null;
            case 'defineName':
              return getNamedRangeDialog() !== null;
          }
        },
        onPasteSpecial: () =>
          pasteSpecialDialog?.open(undefined, { origin: 'contextMenu', commandId: 'pasteSpecial' }),
        onInsertHyperlink: () =>
          getHyperlinkDialog()?.open({ origin: 'contextMenu', commandId: 'insertHyperlink' }),
        onPickFromList: flags.validation ? () => validation?.openPickList() : undefined,
        onOpenHyperlink: (url) => {
          followHyperlink({ store, wb, ...hyperlinkFollow }, url);
        },
        onEditComment: () =>
          getCommentDialog()?.open({
            entry: { origin: 'contextMenu', commandId: 'insertComment' },
          }),
        onDefineName: () =>
          getNamedRangeDialog()?.open({ origin: 'contextMenu', commandId: 'defineName' }),
        getLocale,
        onOpenPivotFieldList: getWorkbookObjects
          ? (sheetIndex, pivotIndex) => {
              getWorkbookObjects()?.openPivotFieldList(sheetIndex, pivotIndex);
            }
          : undefined,
        isPivotFieldListOpen: getWorkbookObjects
          ? () => getWorkbookObjects()?.isPivotFieldListOpen() ?? false
          : undefined,
        onClosePivotFieldList: getWorkbookObjects
          ? () => getWorkbookObjects()?.hidePivotFieldList()
          : undefined,
        getClipboardSnapshot: clipboardH ? () => clipboardH.getSnapshot() : undefined,
        onToggleWatch: flags.watchWindow
          ? (addr) => {
              const watches = store.getState().watch.watches;
              const isOn = watches.some(
                (w) => w.sheet === addr.sheet && w.row === addr.row && w.col === addr.col,
              );
              if (isOn) mutators.removeWatch(store, addr);
              else {
                mutators.addWatch(store, addr);
                mutators.setWatchPanelOpen(store, true);
              }
            }
          : undefined,
        isWatched: flags.watchWindow
          ? (addr) =>
              store
                .getState()
                .watch.watches.some(
                  (w) => w.sheet === addr.sheet && w.row === addr.row && w.col === addr.col,
                )
          : undefined,
      })
    : (): void => {};
  const contextMenu = flags.contextMenu
    ? ({
        dispose: () => {
          detachContextMenu();
        },
      } satisfies ExtensionHandle)
    : null;
  const findReplace = flags.findReplace
    ? attachFindReplace({
        host,
        store,
        wb,
        strings,
        history,
        getLocale,
        onActivateSheet: activateSheet,
        onAfterCommit: refreshCells,
      })
    : null;
  const validation = flags.validation
    ? attachValidationList({
        grid,
        store,
        wb,
        onAfterCommit: refreshCells,
      })
    : null;
  const validationPrompt = flags.validation ? attachValidationPrompt({ grid, store }) : null;
  const tableTotals = attachTableTotals({
    grid,
    store,
    history,
    wb,
    strings,
    onAfterCommit: refreshCells,
    onMoreFunctions: () => beginEditActiveCell(),
  });
  const quickAnalysis = flags.quickAnalysis
    ? attachQuickAnalysis({
        host,
        store,
        wb,
        strings,
        history,
        onAfterCommit: refreshCells,
        invalidate: () => renderer.invalidate(),
        onOpenPivotTable: flags.pivotTableDialog ? () => getPivotTableDialog()?.open() : undefined,
        canOpenPivotTable: () => !!getPivotTableDialog(),
        canCreateChart: () =>
          flags.charts && !!getSessionCharts() && wb.capabilities.charts === true,
      })
    : null;

  const beginEditActiveCell = (caretClientX?: number): void => {
    const s = store.getState();
    const a = s.selection.active;
    const seed = editSeedForAddr(s, wb, a);
    editor.begin(seed, {
      mode: 'edit',
      ...(caretClientX === undefined ? {} : { caretClientX }),
    });
  };
  const onDblClick = (e: MouseEvent): void => {
    if (e.button !== 0) return;
    // Pointer-level double-click affordances (autofit / fill) run earlier on
    // this same node and consume their event with preventDefault().
    if (e.defaultPrevented) return;
    if (editor.isActive()) return;
    if (getFormatPainter()?.isActive()) return;
    // Formula-bar reference picking preserves its focus. Do not start a second
    // inline edit from the persisted, still-old formula while that draft is live.
    if (getFormulaBarEditor?.()?.isFormulaEdit()) return;
    // The Mac desktop puts the caret where the cell was double-clicked.
    const mac = isMacHost(host);
    beginEditActiveCell(mac ? e.clientX : undefined);
    e.preventDefault();
  };
  grid.addEventListener('dblclick', onDblClick);

  const publishAuthoritativeCells = (events: readonly BufferedAtomicValue[]): void => {
    const addresses = events.map((event) => event.addr);
    const formulas = wb.cellFormulas(addresses);
    const current = events.map((event) => ({
      key: addrKey(event.addr),
      value: { ...wb.getValue(event.addr) },
      formula: formulas.get(addrKey(event.addr)) ?? null,
    }));
    store.setState((s) => {
      const cells = new Map(s.data.cells);
      for (const entry of current)
        cells.set(entry.key, { value: entry.value, formula: entry.formula });
      return { ...s, data: { ...s.data, cells } };
    });
  };

  const publishOrdinaryValue = (e: Extract<ChangeEvent, { kind: 'value' }>): void => {
    const formula = wb.cellFormula(e.addr);
    const cell = { value: e.next, formula };
    store.setState((s) => {
      const cells = new Map(s.data.cells);
      cells.set(addrKey(e.addr), cell);
      return { ...s, data: { ...s.data, cells } };
    });
    emitter.emit('cellChange', { addr: e.addr, value: e.next, formula });
  };

  const publishAtomicValue = (e: Extract<ChangeEvent, { kind: 'value' }>): void => {
    const meta = e.atomicBatch;
    if (
      !meta ||
      !Number.isSafeInteger(meta.id) ||
      meta.id < 1 ||
      !Number.isSafeInteger(meta.index) ||
      !Number.isSafeInteger(meta.size) ||
      meta.size < 1 ||
      meta.index < 0 ||
      meta.index >= meta.size
    ) {
      if (meta) pendingAtomic.delete(meta.id);
      publishOrdinaryValue(e);
      return;
    }

    const existing = pendingAtomic.get(meta.id);
    const batch = existing ?? { size: meta.size, slots: [], count: 0 };
    if (batch.size !== meta.size || batch.slots[meta.index]) {
      pendingAtomic.delete(meta.id);
      publishOrdinaryValue(e);
      return;
    }
    if (!existing) {
      batch.slots.length = meta.size;
      pendingAtomic.set(meta.id, batch);
    }
    batch.slots[meta.index] = {
      addr: e.addr,
      value: e.next,
      formula: meta.formula,
      publicPayload: { addr: e.addr, value: e.next, formula: meta.formula },
    };
    batch.count += 1;
    if (batch.count !== batch.size) return;

    pendingAtomic.delete(meta.id);
    const complete = batch.slots;
    if (complete.some((event) => event === undefined)) {
      publishOrdinaryValue(e);
      return;
    }
    const buffered = complete as BufferedAtomicValue[];
    publishAuthoritativeCells(buffered);
    for (const event of buffered) emitter.emit('cellChange', event.publicPayload);
  };

  const unsubWb = wb.subscribe((e: ChangeEvent) => {
    if (e.kind === 'value') {
      publishAtomicValue(e);
    } else if (e.kind === 'recalc') {
      refreshChartsAfterRecalc(wb, store);
      scheduleRecalcRefresh();
      emitter.emit('recalc', { dirty: e.dirty });
    } else if (
      e.kind === 'sheet-add' ||
      e.kind === 'sheet-rename' ||
      e.kind === 'sheet-remove' ||
      e.kind === 'sheet-move' ||
      e.kind === 'sheet-copy'
    ) {
      getSheetTabs()?.update();
      if (e.kind !== 'sheet-add') reloadDrawingsAfterSheetEdit(wb, store, e);
    }
  });

  // A store change the engine refused has been reverted; tell the user.
  const unsubRefusals = wb.subscribeWriteRefusals((refusal) => {
    void showCommandRefusal({
      host,
      strings,
      title: writeRefusalTitle(refusal.kind, strings),
      reason: 'engine-failed',
    });
  });

  // A write through the public facade changes the engine only; bring the store slice it touched
  // back in line so the grid shows it.
  const unsubFacade = wb.subscribeFacadeChanges((change: FacadeChange) => {
    const active = store.getState().data.sheetIndex;
    const sheet = change.sheet ?? active;
    if (sheet < 0 || sheet >= wb.sheetCount) return;
    const hydrate = (): void => {
      // A report on the sheet on show also refits its columns; another sheet only reloads.
      if (change.kind === 'pivot' && sheet === active) {
        finishPivotReportUpdate(store, wb, { theme: resolveTheme(host) });
        return;
      }
      for (const fn of FACADE_HYDRATORS[change.kind]) fn(wb, store, sheet, change.range);
    };
    wb.withEngineSyncMuted(hydrate);
    if (change.kind === 'names' || change.kind === 'protection') {
      hydrateWorkbookMetadataFromEngine(wb, store);
    }
    renderer.invalidate();
  });

  return {
    editor,
    pasteSpecialDialog,
    findReplace,
    validation,
    validationPrompt,
    validationAlert,
    quickAnalysis,
    clipboardH,
    contextMenu,
    followHyperlink: (target) => followHyperlink({ store, wb, ...hyperlinkFollow }, target),
    activateSheet,
    unbind: () => {
      bound = false;
      detachPtr();
      detachSplitDrag();
      pageBandEditor.detach();
      autoFillOptions.detach();
      pasteOptions.detach();
      detachKey();
      detachTabNavigationSession();
      clipboardH?.detach();
      detachContextMenu();
      findReplace?.detach();
      pasteSpecialDialog?.detach();
      validation?.detach();
      validationPrompt?.detach();
      validationAlert?.detach();
      tableTotals.detach();
      setPivotGuardSink(wb, null);
      setDataTableGuardSink(wb, null);
      setProtectionAlertSink(wb, null);
      quickAnalysis?.detach();
      grid.removeEventListener('dblclick', onDblClick);
      unsubWb();
      unsubFacade();
      unsubRefusals();
      pendingAtomic.clear();
      if (editor.isActive()) editor.cancel();
    },
  };
}
