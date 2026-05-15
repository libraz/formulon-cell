import { CellRegistry } from './cells.js';
import { History } from './commands/history.js';
import { printSheet } from './commands/print.js';
import {
  isSheetProtected,
  setProtectedSheet,
  toggleProtectedSheet,
} from './commands/protection.js';
import {
  clearTraceArrows,
  traceDependents as traceDependentArrows,
  tracePrecedents as tracePrecedentArrows,
} from './commands/traces.js';
import { WorkbookHandle } from './engine/workbook-handle.js';
import { SpreadsheetEmitter } from './events.js';
import {
  dedupeById,
  type Extension,
  type ExtensionContext,
  type ExtensionHandle,
  flattenExtensions,
  resolveFlags,
  sortByPriority,
} from './extensions/index.js';
import { FormulaRegistry } from './formula.js';
import { createI18nController } from './i18n/controller.js';
import type { Strings } from './i18n/strings.js';
import { attachCellStylesGallery } from './interact/cell-styles-gallery.js';
import { attachCfRulesDialog } from './interact/cf-rules-dialog.js';
import { attachExternalLinksDialog } from './interact/external-links-dialog.js';
import { attachFilterDropdown, type FilterDropdownHandle } from './interact/filter-dropdown.js';
import { createMountChrome } from './mount/chrome.js';
import { attachChromeSync, type ChromeSyncController } from './mount/chrome-sync.js';
import {
  attachEngineBinding,
  type EngineBinding,
  WB_REGISTRY_IDS,
} from './mount/engine-binding.js';
import { attachFormulaBarController } from './mount/formula-bar.js';
import { prepareMountHost, releaseMountHost } from './mount/host.js';
import {
  createAutocompleteStub,
  createHostFeatureController,
  createHostFeatureState,
  HOST_FEATURE_USES_STRINGS,
  HOST_TOGGLEABLE_IDS,
  WB_TOGGLEABLE_IDS,
} from './mount/host-features.js';
import { createHostShortcutHandler } from './mount/host-shortcuts.js';
import {
  dispatchWorkbookObjectSummaries,
  hydrateActiveSheetFromEngine,
  hydrateWorkbookMetadataFromEngine,
} from './mount/hydration.js';
import {
  attachSheetTabsController,
  type SheetTabsController,
} from './mount/sheet-tabs-controller.js';
import type { MountOptions, SpreadsheetInstance } from './mount/types.js';
import { GridRenderer, getErrorTriangleHits } from './render/grid.js';
import { createSpreadsheetStore, mutators } from './store/store.js';
import { resolveTheme } from './theme/resolve.js';

export type { MountOptions, SpreadsheetInstance } from './mount/types.js';

function mountErrorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function renderMountError(host: HTMLElement, error: unknown): void {
  const panel = document.createElement('div');
  panel.className = 'fc-mount-error';
  panel.setAttribute('role', 'alert');

  const title = document.createElement('strong');
  title.textContent = 'Spreadsheet engine unavailable';

  const help = document.createElement('p');
  help.textContent =
    'The formulon WASM engine could not start. Serve the page with COOP: same-origin and COEP: require-corp so SharedArrayBuffer is available.';

  const detail = document.createElement('code');
  detail.textContent = mountErrorMessage(error);

  panel.append(title, help, detail);
  host.replaceChildren(panel);
}

/**
 * Mount a spreadsheet onto a DOM host. Returns an instance with imperative
 * controls. The host element is taken over — its existing children are
 * cleared. Idempotent dispose.
 */
export const Spreadsheet = {
  async mount(host: HTMLElement, opts: MountOptions = {}): Promise<SpreadsheetInstance> {
    if (!host) throw new Error('Spreadsheet.mount: host element required');

    // Reactive strings — extensions read `strings` (a `let` re-assigned by
    // the i18n subscription) so any future setStrings hook lands on a fresh
    // snapshot. v0.1 built-ins still snapshot at attach time; v0.2 will
    // wire setStrings hooks for live label updates.
    const i18n = createI18nController({ locale: opts.locale, overlay: opts.strings });
    let strings: Strings = i18n.strings;
    let flags = resolveFlags(opts.features);
    const emitter = new SpreadsheetEmitter();
    const formulaRegistry = new FormulaRegistry();
    if (opts.functions) {
      for (const f of opts.functions) {
        formulaRegistry.register(f.name, f.impl, f.meta ?? {});
      }
    }

    const instanceId = prepareMountHost(host, strings, opts.theme);
    host.dataset.fcEngineState = 'loading';

    let sheetTabsController: SheetTabsController | null = null;

    // Track ownership before seeding — only owned (default-created) workbooks
    // should be touched by `seed`. Pre-loaded workbooks are the consumer's
    // data and must not be overwritten by the demo helper.
    let ownsWb = !opts.workbook;
    let wb: WorkbookHandle;
    try {
      wb = opts.workbook ?? (await WorkbookHandle.createDefault());
      if (opts.seed && ownsWb) opts.seed(wb);
    } catch (err) {
      host.dataset.fcEngineState = 'error';
      try {
        opts.onError?.(err);
      } catch (hookErr) {
        console.error('formulon-cell: mount error handler threw', hookErr);
      }
      if (opts.renderError !== false) renderMountError(host, err);
      throw err;
    }

    const {
      formulabar,
      tag,
      fx,
      fxCancel,
      fxAccept,
      fxInput,
      viewbar,
      grid,
      canvas,
      a11y,
      statusbar,
      firstSheet,
      lastSheet,
      sheetTabs,
      addSheetBtn,
      sheetMenu,
      watchDock,
      refreshFormulaBarLabels,
      setChromeAttached,
    } = createMountChrome({
      host,
      getStrings: () => strings,
      flags,
      onSheetTabContextMenu: (idx, tab, x, y) => {
        sheetTabsController?.switchSheet(idx);
        sheetTabsController?.showMenu(idx, tab, x, y);
      },
    });

    const store = createSpreadsheetStore();
    if (opts.theme) mutators.setTheme(store, opts.theme);

    // Unified undo/redo. Attach BEFORE seed-cell hydration so the seed itself
    // doesn't pollute the stack — but seed runs above on the wb. Clear the
    // stack after attach to drop any pre-attach entries (none expected, but
    // cheap insurance).
    const history = new History();
    wb.attachHistory(history);
    history.clear();

    hydrateActiveSheetFromEngine(wb, store);
    hydrateWorkbookMetadataFromEngine(wb, store);
    dispatchPassthroughSummary();

    function dispatchPassthroughSummary(): void {
      // Surface preserved OOXML objects (charts/drawings/pivot parts) and
      // Spreadsheet Tables as host events so chrome (status bar, toast) can show a
      // read-only/editing-limited badge. Pivot layouts are rendered when the
      // engine exposes projection, but the object definition is still not
      // authorable from the UI.
      dispatchWorkbookObjectSummaries(host, wb);
    }

    function hydrateActiveSheet(): void {
      hydrateActiveSheetFromEngine(wb, store);
    }

    const cellRegistry = new CellRegistry();
    const renderer = new GridRenderer({
      host: grid,
      canvas,
      getState: () => store.getState(),
      getTheme: () => resolveTheme(host),
      onViewportSize: (rowCount, colCount) => mutators.setViewportSize(store, rowCount, colCount),
      getWb: () => wb,
      getLocale: () => i18n.locale,
      getDisplay: (addr, value, formula, format) =>
        cellRegistry.resolveDisplay({ addr, value, formula, format }),
    });
    const unsubCellRegistry = cellRegistry.subscribe(() => renderer.invalidate());
    renderer.resize();
    sheetTabsController = attachSheetTabsController({
      addSheetBtn,
      firstSheet,
      getStrings: () => strings,
      getWb: () => wb,
      history,
      host,
      hydrateActiveSheet,
      invalidate: () => renderer.invalidate(),
      lastSheet,
      refreshStatusBar: () => featureState.statusBar?.refresh(),
      sheetMenu,
      sheetTabs,
      store,
    });
    sheetTabsController.update();

    // Always-on host features — not toggleable via `MountOptions.features`.
    // Held in `let` so the locale-change subscription can rebuild them with
    // fresh strings; detach+reattach is the v0.2 fallback for dialogs that
    // don't expose a `setStrings` hook.
    let externalLinksDialog = attachExternalLinksDialog({
      host,
      getWb: () => wb,
      strings,
    });
    let cfRulesDialog = attachCfRulesDialog({
      host,
      getWb: () => wb,
      getActiveSheet: () => store.getState().data.sheetIndex,
      onChanged: () => renderer.invalidate(),
      strings,
    });
    const cellStylesGallery = attachCellStylesGallery({
      host,
      store,
      history,
      getWb: () => wb,
    });
    // Filter dropdown — opens when the pointer dispatches `fc:openfilter`
    // from a clicked column-filter chevron. Rebuilt on locale change so its
    // captured strings stay fresh; no public toggle.
    let filterDropdown: FilterDropdownHandle = attachFilterDropdown({ store, strings });
    interface OpenFilterDetail {
      range: import('./engine/types.js').Range;
      col: number;
      anchor: { x: number; y: number; h: number; clientX: number; clientY: number };
    }
    const onOpenFilter = (e: Event): void => {
      const detail = (e as CustomEvent<OpenFilterDetail>).detail;
      if (!detail) return;
      // The dropdown is positioned with `position: fixed`, so it expects
      // viewport-relative coords. The pointer payload's `x/y` are host-relative;
      // use `clientX/clientY` instead. `- 4` matches the chevron offset.
      filterDropdown.open(detail.range, detail.col, {
        x: detail.anchor.clientX,
        y: detail.anchor.clientY - 4,
        h: detail.anchor.h,
      });
    };
    host.addEventListener('fc:openfilter', onOpenFilter);

    const featureRegistry = new Map<string, ExtensionHandle>();
    const wrapHandle = (raw: unknown, detach: () => void): ExtensionHandle => {
      const h = (
        raw && typeof raw === 'object' ? (raw as Record<string, unknown>) : {}
      ) as ExtensionHandle;
      h.dispose = detach;
      return h;
    };

    const autocompleteStub = createAutocompleteStub();
    const featureState = createHostFeatureState(autocompleteStub);

    const syncBindingFeatures = (current: EngineBinding): void => {
      for (const id of WB_REGISTRY_IDS) featureRegistry.delete(id);
      if (current.clipboardH) {
        featureRegistry.set(
          'clipboard',
          wrapHandle(current.clipboardH, () => current.clipboardH?.detach()),
        );
      }
      if (current.pasteSpecialDialog) {
        featureRegistry.set(
          'pasteSpecial',
          wrapHandle(current.pasteSpecialDialog, () => current.pasteSpecialDialog?.detach()),
        );
      }
      if (current.quickAnalysis) {
        featureRegistry.set(
          'quickAnalysis',
          wrapHandle(current.quickAnalysis, () => current.quickAnalysis?.detach()),
        );
      }
      if (current.contextMenu) featureRegistry.set('contextMenu', current.contextMenu);
      if (current.findReplace) {
        featureRegistry.set(
          'findReplace',
          wrapHandle(current.findReplace, () => current.findReplace?.detach()),
        );
      }
      if (current.validation) {
        featureRegistry.set(
          'validation',
          wrapHandle(current.validation, () => current.validation?.detach()),
        );
      }
    };

    let chromeSync: ChromeSyncController | null = null;
    const updateChrome = (): void => chromeSync?.updateChrome();

    const bindEngine = (currentWb: WorkbookHandle): EngineBinding =>
      attachEngineBinding({
        emitter,
        flags,
        getCommentDialog: () => featureState.commentDialog,
        getFormatDialog: () => featureState.formatDialog,
        getFormatPainter: () => featureState.formatPainter,
        getGoToDialog: () => featureState.goToDialog,
        getHyperlinkDialog: () => featureState.hyperlinkDialog,
        getPivotTableDialog: () => featureState.pivotTableDialog,
        getSessionCharts: () => featureState.sessionCharts,
        getSheetTabs: () => sheetTabsController,
        grid,
        history,
        host,
        renderer,
        store,
        strings,
        tag,
        updateChrome,
        wb: currentWb,
      });

    let binding = bindEngine(wb);
    syncBindingFeatures(binding);

    const onHostKey = createHostShortcutHandler({
      findReplace: () => binding.findReplace,
      formatDialog: () => featureState.formatDialog,
      formatPainter: () => featureState.formatPainter,
      goToDialog: () => featureState.goToDialog,
      history,
      hostTag: tag,
      hyperlinkDialog: () => featureState.hyperlinkDialog,
      invalidate: () => renderer.invalidate(),
      namedRangeDialog: () => featureState.namedRangeDialog,
      pasteSpecialDialog: () => binding.pasteSpecialDialog,
      quickAnalysis: () => binding.quickAnalysis,
      store,
      wb: () => wb,
    });

    // Error / validation triangle clicks. Bound on the canvas by the
    // `errorIndicators` attacher above. We use `click` (not `pointerdown`)
    // so the existing pointer handler gets to set the active cell first —
    // that way the menu and the cell select agree on the addr the user
    // just clicked.
    const onCanvasClick = (e: MouseEvent): void => {
      if (!featureState.errorMenu) return;
      if (e.button !== 0) return;
      const rect = canvas.getBoundingClientRect();
      const lx = e.clientX - rect.left;
      const ly = e.clientY - rect.top;
      // Pad by 2px on each side so the 6px corner triangle is comfortable to
      // hit on touch / coarse-pointer devices.
      const pad = 2;
      for (const hit of getErrorTriangleHits()) {
        const r = hit.rect;
        if (lx < r.x - pad || lx > r.x + r.w + pad || ly < r.y - pad || ly > r.y + r.h + pad) {
          continue;
        }
        e.stopPropagation();
        e.preventDefault();
        featureState.errorMenu.open(hit.addr, e.clientX, e.clientY, hit.kind);
        return;
      }
    };

    const formulaBar = attachFormulaBarController({
      cancelBindingEditor: () => {
        if (binding.editor.isActive()) binding.editor.cancel();
      },
      formulabar,
      fxAccept,
      fxCancel,
      fxInput,
      getArgHelper: () => featureState.fxArgHelper,
      getAutocomplete: () => featureState.fxAutocomplete,
      host,
      store,
      updateChrome,
      wb: () => wb,
    });

    chromeSync = attachChromeSync({
      a11y,
      emitter,
      fxInput,
      getFormulaEditing: () => formulaBar.isEditing(),
      getSheetTabs: () => sheetTabsController,
      getWb: () => wb,
      host,
      invalidate: () => renderer.invalidate(),
      store,
      tag,
    });

    // Resize observer — we follow the host, not the window.
    const ro = new ResizeObserver(() => renderer.resize());
    ro.observe(grid);

    let disposed = false;

    // User extensions — additive on top of built-ins. Run after built-ins
    // and the engine binding so they can read other features via
    // `ctx.resolve()`.
    const userHandles = new Map<string, ExtensionHandle>();
    const refreshCells = (): void => {
      mutators.replaceCells(store, wb.cells(store.getState().data.sheetIndex));
    };
    const wbListeners = new Set<(next: WorkbookHandle) => void>();
    const ctx: ExtensionContext = {
      host,
      formulabar,
      viewbar,
      grid,
      statusbar,
      canvas,
      a11y,
      store,
      history,
      i18n,
      getWb: () => wb,
      refreshCells,
      invalidate: () => renderer.invalidate(),
      resolve: <T extends ExtensionHandle = ExtensionHandle>(id: string): T | undefined =>
        (featureRegistry.get(id) ?? userHandles.get(id)) as T | undefined,
      onWorkbookChange: (fn) => {
        wbListeners.add(fn);
        return () => {
          wbListeners.delete(fn);
        };
      },
    };

    const mountExtension = (ext: Extension): void => {
      if (userHandles.has(ext.id) || featureRegistry.has(ext.id)) {
        // last-wins via remove + re-add; users explicitly opt in
        userHandles.get(ext.id)?.dispose();
        userHandles.delete(ext.id);
      }
      const handle = ext.setup(ctx);
      if (handle) userHandles.set(ext.id, handle);
    };
    // Combined view exposed on `instance.features` — built-ins + user.
    const featuresView: Record<string, ExtensionHandle | undefined> = {};
    const refreshFeaturesView = (): void => {
      for (const k of Object.keys(featuresView)) delete featuresView[k];
      for (const [k, v] of featureRegistry) featuresView[k] = v;
      for (const [k, v] of userHandles) featuresView[k] = v;
    };
    const hostFeatures = createHostFeatureController({
      autocompleteStub,
      canvas,
      emitter,
      featureRegistry,
      flags: () => flags,
      formulaRegistry,
      fx,
      fxInput,
      getFormulaBar: () => formulaBar,
      getOnCanvasClick: () => onCanvasClick,
      getOnHostKey: () => onHostKey,
      getSheetTabs: () => sheetTabsController,
      grid,
      history,
      host,
      i18nLocale: () => i18n.locale,
      refreshFeaturesView,
      renderer,
      setChromeAttached,
      state: featureState,
      statusbar,
      store,
      strings: () => strings,
      viewbar,
      watchDock,
      wb: () => wb,
      wrapHandle,
    });
    const attachHostFeature = hostFeatures.attach;
    const detachHostFeature = hostFeatures.detach;

    // Initial host-feature attach — runs after every helper closure
    // (`onCanvasClick`, `onHostKey`, `syncFxRefs`, `commitFx`) is in
    // scope so the attacher bodies can resolve them at call time.
    for (const id of HOST_TOGGLEABLE_IDS) {
      if (flags[id as keyof typeof flags]) attachHostFeature(id);
    }

    if (opts.extensions) {
      const sorted = sortByPriority(dedupeById(flattenExtensions(opts.extensions)));
      for (const ext of sorted) mountExtension(ext);
    }
    refreshFeaturesView();

    // Locale change → push fresh strings everywhere. Built-ins that ship a
    // `setStrings` hook live-update labels in place; the rest are rebuilt by
    // detaching and re-attaching with the new dictionary in their closure.
    const unsubI18n = i18n.subscribe((next) => {
      strings = next;
      host.setAttribute('aria-label', strings.a11y.spreadsheet);
      tag.setAttribute('aria-label', strings.a11y.nameBox);
      refreshFormulaBarLabels();
      featureState.fxAutocomplete.setLabels(next.autocomplete);
      featureState.fxArgHelper?.setLabels(next.argHelper);
      sheetTabsController?.update();
      renderer.invalidate();

      // Always-on dialogs: rebuild — none of these expose setStrings yet.
      externalLinksDialog.detach();
      externalLinksDialog = attachExternalLinksDialog({ host, getWb: () => wb, strings });
      cfRulesDialog.detach();
      cfRulesDialog = attachCfRulesDialog({
        host,
        getWb: () => wb,
        getActiveSheet: () => store.getState().data.sheetIndex,
        onChanged: () => renderer.invalidate(),
        strings,
      });
      filterDropdown.detach();
      filterDropdown = attachFilterDropdown({ store, strings });

      // Toggleable host features: prefer setStrings when the handle exposes
      // it, otherwise fall back to detach+reattach.
      for (const id of HOST_TOGGLEABLE_IDS) {
        const handle = featureRegistry.get(id);
        if (!handle) continue;
        if (typeof handle.setStrings === 'function') {
          handle.setStrings(next);
        } else if (HOST_FEATURE_USES_STRINGS.has(id)) {
          detachHostFeature(id);
          attachHostFeature(id);
        }
      }

      // Engine-bound attaches (clipboard, paste-special, context-menu,
      // find-replace, validation) live inside `binding`. Rebuild it.
      binding.unbind();
      binding = bindEngine(wb);
      syncBindingFeatures(binding);
      featureState.viewToolbar?.bindWorkbook(wb);
      featureState.workbookObjects?.bindWorkbook(wb);
      featureState.pivotTableDialog?.bindWorkbook(wb);

      // User extensions opt-in via setStrings.
      for (const handle of userHandles.values()) handle.setStrings?.(next);

      emitter.emit('localeChange', { locale: i18n.locale, strings: next });
    });

    host.dataset.fcEngineState = wb.isStub ? 'ready-stub' : 'ready';

    return {
      host,
      get workbook() {
        return wb;
      },
      store,
      history,
      i18n,
      features: featuresView,
      get formatPainter() {
        return featureState.formatPainter ?? undefined;
      },
      formula: formulaRegistry,
      cells: cellRegistry,
      use(input) {
        const sorted = sortByPriority(dedupeById(flattenExtensions([input])));
        for (const ext of sorted) mountExtension(ext);
        refreshFeaturesView();
      },
      remove(id) {
        const handle = userHandles.get(id);
        if (!handle) return false;
        handle.dispose();
        userHandles.delete(id);
        refreshFeaturesView();
        return true;
      },
      setFeatures(next) {
        const prevFlags = flags;
        const nextFlags = resolveFlags(next);
        const shouldRebuildViewToolbarObjects =
          prevFlags.viewToolbar &&
          nextFlags.viewToolbar &&
          prevFlags.workbookObjects !== nextFlags.workbookObjects;
        flags = nextFlags;
        // Diff host-level features and dispatch attach/detach.
        for (const id of HOST_TOGGLEABLE_IDS) {
          const k = id as keyof typeof prevFlags;
          const was = prevFlags[k];
          const now = nextFlags[k];
          if (was === now) continue;
          if (was && !now) detachHostFeature(id);
          else if (!was && now) attachHostFeature(id);
        }
        // Wb-side rebuild only when a wb-bound feature flipped — keeps the
        // editor / pointer / undo state intact when only host-level flags
        // change.
        const wbChanged = WB_TOGGLEABLE_IDS.some(
          (id) =>
            prevFlags[id as keyof typeof prevFlags] !== nextFlags[id as keyof typeof nextFlags],
        );
        if (shouldRebuildViewToolbarObjects && featureState.viewToolbar) {
          detachHostFeature('viewToolbar');
          attachHostFeature('viewToolbar');
        }
        if (wbChanged) {
          binding.unbind();
          binding = bindEngine(wb);
          syncBindingFeatures(binding);
          featureState.viewToolbar?.bindWorkbook(wb);
          featureState.workbookObjects?.bindWorkbook(wb);
        }
        refreshFeaturesView();
      },
      setExtensions(next) {
        // Dispose all currently-mounted user extensions, then re-mount the
        // new list. Built-ins are untouched — use `setFeatures` for those.
        for (const handle of userHandles.values()) handle.dispose();
        userHandles.clear();
        if (next?.length) {
          const sorted = sortByPriority(dedupeById(flattenExtensions(next)));
          for (const ext of sorted) mountExtension(ext);
        }
        refreshFeaturesView();
      },
      openConditionalDialog() {
        featureState.conditionalDialog?.open();
      },
      openIterativeDialog() {
        featureState.iterativeDialog?.open();
      },
      openExternalLinksDialog() {
        externalLinksDialog.open();
      },
      openCfRulesDialog() {
        cfRulesDialog.open();
      },
      openCellStylesGallery() {
        cellStylesGallery.open();
      },
      openFunctionArguments(seedName?: string) {
        featureState.fxDialog?.open(seedName);
      },
      openHyperlinkDialog() {
        featureState.hyperlinkDialog?.open();
      },
      openCommentDialog() {
        featureState.commentDialog?.open();
      },
      openFindReplace() {
        binding.findReplace?.open();
      },
      closeFindReplace() {
        binding.findReplace?.close();
      },
      openPasteSpecial() {
        binding.pasteSpecialDialog?.open();
      },
      openNamedRangeDialog() {
        featureState.namedRangeDialog?.open();
      },
      openPageSetup() {
        featureState.pageSetupDialog?.open();
      },
      print() {
        // The print command is wired through the same flag as the dialog —
        // when the feature is off, both call sites are no-ops. Skip if the
        // dialog never attached so consumers can rely on the gate.
        if (!featureState.pageSetupDialog) return;
        printSheet(wb, store, store.getState().data.sheetIndex, host);
      },
      recalc() {
        wb.recalc();
        mutators.replaceCells(store, wb.cells(store.getState().data.sheetIndex));
        renderer.invalidate();
      },
      openFormatDialog() {
        featureState.formatDialog?.open();
      },
      openGoToSpecial() {
        featureState.goToDialog?.open();
      },
      openWatchWindow() {
        featureState.watchPanel?.open();
      },
      closeWatchWindow() {
        featureState.watchPanel?.close();
      },
      toggleWatchWindow() {
        featureState.watchPanel?.toggle();
      },
      openQuickAnalysis() {
        const userQuick = userHandles.get('quickAnalysis') as
          | (ExtensionHandle & { open?: () => void })
          | undefined;
        if (userQuick?.open) {
          userQuick.open();
          return;
        }
        binding.quickAnalysis?.open();
      },
      openWorkbookObjects() {
        const userObjects = userHandles.get('workbookObjects') as
          | (ExtensionHandle & { open?: () => void })
          | undefined;
        if (userObjects?.open) {
          userObjects.open();
          return;
        }
        featureState.workbookObjects?.open();
      },
      openPivotTableDialog() {
        const userPivot = userHandles.get('pivotTableDialog') as
          | (ExtensionHandle & { open?: () => void })
          | undefined;
        if (userPivot?.open) {
          userPivot.open();
          return;
        }
        featureState.pivotTableDialog?.open();
      },
      addSlicer(input) {
        if (!featureState.slicer) {
          throw new Error('addSlicer: features.slicer is disabled');
        }
        return featureState.slicer.addSlicer(input);
      },
      removeSlicer(id) {
        featureState.slicer?.removeSlicer(id);
      },
      toggleSheetProtection() {
        toggleProtectedSheet(store, store.getState().data.sheetIndex, { workbook: wb });
        renderer.invalidate();
      },
      setSheetProtected(on: boolean, password?: string) {
        setProtectedSheet(store, store.getState().data.sheetIndex, on, { workbook: wb, password });
        renderer.invalidate();
      },
      isSheetProtected() {
        return isSheetProtected(store.getState(), store.getState().data.sheetIndex);
      },
      tracePrecedents() {
        tracePrecedentArrows(store, wb);
        renderer.invalidate();
      },
      traceDependents() {
        traceDependentArrows(store, wb);
        renderer.invalidate();
      },
      clearTraces() {
        clearTraceArrows(store);
        renderer.invalidate();
      },
      setTheme(t) {
        host.dataset.fcTheme = t;
        mutators.setTheme(store, t);
        renderer.invalidate();
        emitter.emit('themeChange', { theme: t });
      },
      undo() {
        const ok = history.undo();
        if (ok) {
          // Force a recalc — undo's per-cell replays may restore values without
          // triggering recalc (setNumber/setText skip it), leaving formula cells
          // stale. One end-of-batch recalc fixes them all.
          wb.recalc();
          mutators.replaceCells(store, wb.cells(store.getState().data.sheetIndex));
        }
        return ok;
      },
      redo() {
        const ok = history.redo();
        if (ok) {
          wb.recalc();
          mutators.replaceCells(store, wb.cells(store.getState().data.sheetIndex));
        }
        return ok;
      },
      async setWorkbook(next) {
        if (next === wb) return;
        binding.unbind();
        if (ownsWb) wb.dispose();
        wb = next;
        ownsWb = true; // we now own the next handle and will dispose it
        wb.attachHistory(history);
        wb.clearViewportHint();
        history.clear();
        const nextSheet = Math.min(
          store.getState().data.sheetIndex,
          Math.max(0, wb.sheetCount - 1),
        );
        mutators.setSheetIndex(store, nextSheet);
        hydrateActiveSheet();
        hydrateWorkbookMetadataFromEngine(wb, store);
        dispatchPassthroughSummary();
        binding = bindEngine(wb);
        syncBindingFeatures(binding);
        featureState.viewToolbar?.bindWorkbook(wb);
        featureState.workbookObjects?.bindWorkbook(wb);
        featureState.namedRangeDialog?.bindWorkbook(wb);
        featureState.pivotTableDialog?.bindWorkbook(wb);
        featureState.statusBar?.refresh();
        sheetTabsController?.update();
        // Notify user extensions so they can rebind their wb references.
        for (const handle of userHandles.values()) handle.rebindWorkbook?.(wb);
        for (const fn of wbListeners) fn(wb);
        updateChrome();
        renderer.invalidate();
        emitter.emit('workbookChange', { workbook: wb });
      },
      on: (name, fn) => emitter.on(name, fn),
      off: (name, fn) => emitter.off(name, fn),
      dispose() {
        if (disposed) return;
        disposed = true;
        emitter.dispose();
        ro.disconnect();
        binding.unbind();
        for (const handle of userHandles.values()) handle.dispose();
        userHandles.clear();
        for (const id of HOST_TOGGLEABLE_IDS) detachHostFeature(id);
        formulaBar.detach();
        sheetTabsController?.detach();
        chromeSync?.detach();
        host.removeEventListener('fc:openfilter', onOpenFilter);
        filterDropdown.detach();
        unsubCellRegistry();
        unsubI18n();
        i18n.dispose();
        renderer.dispose();
        if (ownsWb) wb.dispose();
        releaseMountHost(host, instanceId);
      },
    };
  },
};
