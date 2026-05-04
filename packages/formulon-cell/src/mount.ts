import { CellRegistry } from './cells.js';
import { writeInputValidated } from './commands/coerce-input.js';
import { fillRange } from './commands/fill.js';
import { toggleBold, toggleItalic, toggleStrike, toggleUnderline } from './commands/format.js';
import { History, recordFormatChange } from './commands/history.js';
import { printSheet } from './commands/print.js';
import { extractRefs, rotateRefAt } from './commands/refs.js';
import { flushFormatToEngine, hydrateCellFormatsFromEngine } from './engine/cell-format-sync.js';
import { formatCellForEdit } from './engine/edit-seed.js';
import { hydrateCommentsAndHyperlinksFromEngine } from './engine/format-sync.js';
import { hydrateLayoutFromEngine } from './engine/layout-sync.js';
import { hydrateMergesFromEngine } from './engine/merges-sync.js';
import { summarizePassthroughs, summarizeTables } from './engine/passthrough-sync.js';
import { flushProtectionToEngine, hydrateProtectionFromEngine } from './engine/protection-sync.js';
import { findDependents, findPrecedents } from './engine/refs-graph.js';
import type { CellValue } from './engine/types.js';
import { hydrateValidationsFromEngine } from './engine/validation-sync.js';
import { type ChangeEvent, WorkbookHandle } from './engine/workbook-handle.js';
import {
  SpreadsheetEmitter,
  type SpreadsheetEventHandler,
  type SpreadsheetEventName,
  selectionEquals,
} from './events.js';
import {
  dedupeById,
  type Extension,
  type ExtensionContext,
  type ExtensionHandle,
  type ExtensionInput,
  type FeatureFlags,
  flattenExtensions,
  resolveFlags,
  sortByPriority,
  type ThemeName,
} from './extensions/index.js';
import type { CustomFunction, CustomFunctionMeta } from './formula.js';
import { FormulaRegistry } from './formula.js';
import { createI18nController, type I18nController } from './i18n/controller.js';
import type { DeepPartial, Locale, Strings } from './i18n/strings.js';
import { attachArgHelper } from './interact/arg-helper.js';
import { attachAutocomplete } from './interact/autocomplete.js';
import { attachCellStylesGallery } from './interact/cell-styles-gallery.js';
import { attachClipboard } from './interact/clipboard.js';
import { attachConditionalDialog } from './interact/conditional-dialog.js';
import { attachContextMenu } from './interact/context-menu.js';
import { InlineEditor } from './interact/editor.js';
import { attachErrorMenu, type ErrorMenuHandle } from './interact/error-menu.js';
import { attachExternalLinksDialog } from './interact/external-links-dialog.js';
import { attachFindReplace } from './interact/find-replace.js';
import { attachFormatDialog } from './interact/format-dialog.js';
import { attachFormatPainter, type FormatPainterHandle } from './interact/format-painter.js';
import { attachFxDialog, type FxDialogHandle } from './interact/fx-dialog.js';
import { attachGoToDialog } from './interact/goto-dialog.js';
import { attachHover } from './interact/hover.js';
import { attachHyperlinkDialog } from './interact/hyperlink-dialog.js';
import { attachIterativeDialog } from './interact/iterative-dialog.js';
import { attachKeyboard } from './interact/keyboard.js';
import { attachNamedRangeDialog } from './interact/named-range-dialog.js';
import { attachPageSetupDialog } from './interact/page-setup-dialog.js';
import { attachPasteSpecial } from './interact/paste-special.js';
import { attachPointer } from './interact/pointer.js';
import { attachSlicer, type SlicerHandle } from './interact/slicer.js';
import { attachStatusBar } from './interact/status-bar.js';
import { attachValidationList } from './interact/validation.js';
import { attachWatchPanel } from './interact/watch-panel.js';
import { attachWheel } from './interact/wheel.js';
import { GridRenderer, getErrorTriangleHits } from './render/grid.js';
import {
  createSpreadsheetStore,
  mutators,
  type SlicerSpec,
  type SpreadsheetStore,
} from './store/store.js';
import { resolveTheme } from './theme/resolve.js';

export interface MountOptions {
  /** Pre-loaded workbook (e.g. from xlsx bytes). If omitted, creates a fresh
   *  default workbook. */
  workbook?: WorkbookHandle;
  /** Theme to apply on mount. Switchable later via instance.setTheme. */
  theme?: ThemeName;
  /** Optional initial-cell seeding. Useful for the playground & docs. */
  seed?: (wb: WorkbookHandle) => void;
  /** UI locale for built-in dialogs and menus. Defaults to 'ja'. Swap at
   *  runtime via `instance.i18n.setLocale`. */
  locale?: Locale | (string & {});
  /** Per-string overrides applied on top of the chosen locale. Deep-merged.
   *  For runtime overlays use `instance.i18n.extend`. */
  strings?: DeepPartial<Strings>;
  /** Toggle individual built-in features. Defaults to "all on" (Excel-style
   *  full chrome). Pass a preset (`presets.minimal()` etc.) or your own
   *  `{ findReplace: false, ... }`. Cross-references between features are
   *  handled defensively — disabling format-dialog hides the menu item that
   *  would otherwise open it. */
  features?: FeatureFlags;
  /** Custom extensions added on top of the built-ins. Run after built-ins
   *  in priority order. v0.1 extensions are *additive* — you cannot replace
   *  a built-in via this slot; toggle the built-in via `features` and add
   *  your replacement here. Nested arrays are flattened. */
  extensions?: ExtensionInput[];
  /** Optional initial set of custom functions registered against
   *  `instance.formula`. Names are upper-cased; impls accept `CellValue`s
   *  and return `CellValue | number | string | boolean | null`. */
  functions?: readonly {
    name: string;
    impl: CustomFunction['impl'];
    meta?: CustomFunctionMeta;
  }[];
}

export interface SpreadsheetInstance {
  readonly host: HTMLElement;
  readonly workbook: WorkbookHandle;
  readonly store: SpreadsheetStore;
  /** Unified undo/redo for cell, format, and layout changes. Each user-level
   *  action pushes one entry; transactions (paste, fill drag) are batched. */
  readonly history: History;
  /** Reactive locale + strings registry. `setLocale` swaps the active
   *  dictionary in place; `extend` overlays partial overrides; `register`
   *  adds a brand-new locale. */
  readonly i18n: I18nController;
  /** Snapshot of every loaded extension keyed by id. Built-in feature ids
   *  match the keys on `MountOptions.features`. */
  readonly features: Readonly<Record<string, ExtensionHandle | undefined>>;
  /** Format Painter controls — surfaced so chrome (toolbar buttons) can
   *  arm/disarm and reflect the active state. `undefined` if disabled via
   *  `features.formatPainter: false`. */
  readonly formatPainter: FormatPainterHandle | undefined;
  /** Host-side custom function registry. `register(name, impl, meta?)`
   *  surfaces a name in the autocomplete and exposes
   *  `evaluate(name, args)` for app code that wires derived cells via
   *  `inst.on('cellChange', …)`. Engine-level user-function support
   *  ships when formulon exposes the embind callback hook. */
  readonly formula: FormulaRegistry;
  /** Cell renderer registry — `cells.registerFormatter({match, format})`
   *  substitutes the displayed string for matching cells without
   *  bypassing the canvas paint pipeline. The editor slot
   *  (`registerEditor`) is reserved for v0.2. */
  readonly cells: CellRegistry;
  /** Mount a custom extension after the spreadsheet is already up. */
  use(ext: ExtensionInput): void;
  /** Tear down a previously-mounted extension by id. */
  remove(id: string): boolean;
  /** Live-toggle built-in features. Diffs against the current flag set
   *  and only attaches/detaches what actually changed; keeps editor /
   *  selection / undo state intact when wb-bound features don't flip. */
  setFeatures(next: FeatureFlags): void;
  /** Replace the entire user-extension list. Existing extensions are
   *  disposed; the new list is mounted in priority order. Built-ins are
   *  untouched — use `setFeatures` for those. */
  setExtensions(next: ExtensionInput[] | undefined): void;
  /** Open the conditional-formatting rule manager dialog. No-op when the
   *  feature is disabled. */
  openConditionalDialog(): void;
  /** Open the read-only named-range listing dialog. */
  openNamedRangeDialog(): void;
  /** Open the cell format dialog (Excel ⌘1). */
  openFormatDialog(): void;
  /** Open the Go To Special (F5) dialog. No-op when the feature is disabled. */
  openGoToSpecial(): void;
  /** Open the iterative-calculation settings dialog (Excel File → Options
   *  → Formulas). */
  openIterativeDialog(): void;
  /** Open the read-only "Edit Links" inspector listing every external
   *  workbook reference carried by the active workbook. The list is empty
   *  for fresh workbooks and any package that had no
   *  `<externalReferences>` block. */
  openExternalLinksDialog(): void;
  /** Open the named cell-styles gallery (Excel Home → Cell Styles). */
  openCellStylesGallery(): void;
  /** Open the Function Arguments dialog. Pass `seedName` (case-insensitive)
   *  to skip the picker and jump straight to argument entry for that
   *  function. No-op when the `fxDialog` feature is disabled. */
  openFunctionArguments(seedName?: string): void;
  /** Open the Page Setup dialog (orientation, paper size, margins,
   *  header/footer, print titles). No-op when `features.pageSetup` is off. */
  openPageSetup(): void;
  /** Build a print document for the active sheet and open the browser's
   *  native print dialog. Use the dialog's "Save as PDF" action to export.
   *  No-op when `features.pageSetup` is off. */
  print(): void;
  /** Show the Watch Window panel. No-op when `features.watchWindow` is off. */
  openWatchWindow(): void;
  /** Hide the Watch Window panel. No-op when the feature is off. */
  closeWatchWindow(): void;
  /** Toggle Watch Window visibility. No-op when the feature is off. */
  toggleWatchWindow(): void;
  /** Open a floating slicer panel for the given table column. Returns the
   *  freshly-built `SlicerSpec` (including its auto-assigned id). Throws
   *  when the table or column can't be resolved against the workbook, or
   *  when `features.slicer` is off. */
  addSlicer(input: {
    tableName: string;
    column: string;
    selected?: readonly string[];
    x?: number;
    y?: number;
  }): SlicerSpec;
  /** Remove a slicer by id. No-op when the id isn't tracked. */
  removeSlicer(id: string): void;
  /** Toggle sheet-protection on the currently active sheet. Equivalent to
   *  Excel's Review → Protect Sheet button. Locked cells on protected
   *  sheets gate writes through the command layer. */
  toggleSheetProtection(): void;
  /** Set sheet-protection explicitly. `password` is currently stored
   *  verbatim and NOT enforced — v1 ships without password validation. */
  setSheetProtected(on: boolean, password?: string): void;
  /** True when the active sheet is currently protected. */
  isSheetProtected(): boolean;
  /** Append precedent arrows for the active cell. Same-sheet only. Repeated
   *  calls deduplicate against existing arrows. */
  tracePrecedents(): void;
  /** Append dependent arrows for the active cell. Same-sheet only. */
  traceDependents(): void;
  /** Remove all currently visible trace arrows. */
  clearTraces(): void;
  setTheme(t: ThemeName): void;
  /** Pop the most recent undoable action and revert it. Returns false when
   *  the stack is empty. */
  undo(): boolean;
  /** Re-apply the most recently undone action. Returns false when nothing
   *  to redo. */
  redo(): boolean;
  /** Apply a fresh workbook (e.g. after `loadBytes`). Disposes old one. */
  setWorkbook(next: WorkbookHandle): Promise<void>;
  /** Subscribe to a named lifecycle event. Returns an unsubscribe fn.
   *  Events:
   *  - `cellChange` — a cell value mutated (engine-side); fires once per
   *    cell, after recalc has settled.
   *  - `selectionChange` — active/anchor/range moved.
   *  - `workbookChange` — `setWorkbook` swapped the engine.
   *  - `localeChange` — `i18n.setLocale` switched active dictionary.
   *  - `themeChange` — `setTheme` mutated the host theme.
   *  - `recalc` — engine reported a recalc batch (with dirty-cell keys).
   *  Handlers are isolated: thrown errors are caught and logged, not
   *  propagated to siblings. */
  on<K extends SpreadsheetEventName>(name: K, fn: SpreadsheetEventHandler<K>): () => void;
  /** Imperative unsubscribe. Most code should use the disposer returned
   *  from `on()` instead. */
  off<K extends SpreadsheetEventName>(name: K, fn: SpreadsheetEventHandler<K>): void;
  dispose(): void;
}

// Monotonic counter feeding `data-fc-inst-id` on each mount. Used by the
// dispose path to detect that a later mount has commandeered the same
// host (StrictMode double-mount, hot-reload, etc.).
let MOUNT_COUNTER = 0;

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

    host.classList.add('fc-host');
    host.setAttribute('tabindex', '0');
    // Canvas-rendered grids can't expose ARIA grid/row/cell descendants, so
    // role="grid" would lie about the structure (axe flags it as
    // aria-required-children). Use role="region" with a roledescription so
    // screen readers still announce the surface as a spreadsheet, and let
    // the aria-live mirror inside carry per-cell announcements.
    host.setAttribute('role', 'region');
    host.setAttribute('aria-roledescription', 'spreadsheet');
    host.setAttribute('aria-label', strings.a11y.spreadsheet);
    host.dataset.fcTheme = opts.theme ?? 'paper';
    host.replaceChildren();
    // Stamp a per-mount instance id on the host. Used by `dispose()` so it
    // only clears children if it still owns the host — guards against
    // React 18+ StrictMode double-mount, where a previous instance's
    // deferred dispose would otherwise wipe the host that a newer mount
    // already populated.
    const instanceId = `fc-${++MOUNT_COUNTER}`;
    host.dataset.fcInstId = instanceId;

    // Build chrome: formulabar (top), grid surface, statusbar (bottom).
    const formulabar = document.createElement('div');
    formulabar.className = 'fc-host__formulabar';
    // Name box — typing "A1" / "B5" jumps the active cell. Doubles as the
    // address indicator when not focused.
    const tag = document.createElement('input');
    tag.type = 'text';
    tag.className = 'fc-host__formulabar-tag';
    tag.spellcheck = false;
    tag.autocomplete = 'off';
    tag.setAttribute('aria-label', strings.a11y.nameBox);
    tag.value = 'A1';
    // The fx button opens the Function Arguments dialog. Falls back to a
    // visually-identical decorative span when the feature is disabled so the
    // chrome layout stays the same.
    const fx = document.createElement('button');
    fx.type = 'button';
    fx.className = 'fc-host__formulabar-fx';
    fx.textContent = 'ƒx';
    fx.tabIndex = -1;
    fx.setAttribute('aria-label', strings.fxDialog?.fxButtonLabel ?? 'Insert function');
    const fxInput = document.createElement('textarea');
    fxInput.className = 'fc-host__formulabar-input';
    fxInput.spellcheck = false;
    fxInput.autocomplete = 'off';
    fxInput.rows = 1;
    fxInput.wrap = 'soft';
    fxInput.setAttribute('aria-label', strings.a11y.formulaBar);
    // Excel-style expand/collapse handle. Toggles `data-fc-expanded` on the
    // formulabar so CSS can switch the textarea between 1-row and multi-row.
    const fxExpand = document.createElement('button');
    fxExpand.type = 'button';
    fxExpand.className = 'fc-host__formulabar-expand';
    fxExpand.setAttribute('aria-label', 'Expand formula bar');
    fxExpand.setAttribute('aria-expanded', 'false');
    fxExpand.tabIndex = -1;
    fxExpand.textContent = '⌄';
    fxExpand.addEventListener('click', () => {
      const expanded = formulabar.dataset.fcExpanded === '1';
      if (expanded) {
        delete formulabar.dataset.fcExpanded;
        fxExpand.setAttribute('aria-expanded', 'false');
        fxExpand.textContent = '⌄';
        fxInput.rows = 1;
      } else {
        formulabar.dataset.fcExpanded = '1';
        fxExpand.setAttribute('aria-expanded', 'true');
        fxExpand.textContent = '⌃';
        fxInput.rows = 4;
      }
    });
    formulabar.append(tag, fx, fxInput, fxExpand);

    const grid = document.createElement('div');
    grid.className = 'fc-host__grid';
    const canvas = document.createElement('canvas');
    canvas.className = 'fc-host__canvas';
    grid.appendChild(canvas);

    const a11y = document.createElement('div');
    a11y.className = 'fc-host__a11y';
    a11y.setAttribute('aria-live', 'polite');
    grid.appendChild(a11y);

    const statusbar = document.createElement('div');
    statusbar.className = 'fc-host__statusbar';
    // Watch Window dock — only attached when `features.watchWindow` is on.
    const watchDock = document.createElement('div');
    watchDock.dataset.fcWatch = 'dock';
    watchDock.className = 'fc-host__watchdock';

    // Gate chrome elements on their flags. The grid is always present and
    // anchors slot ordering; chrome slots come and go via `setChromeAttached`
    // so disabled flags reclaim vertical space (no empty bars left behind).
    const setChromeAttached = (
      slot: 'formulabar' | 'statusbar' | 'watchDock',
      on: boolean,
    ): void => {
      const el = slot === 'formulabar' ? formulabar : slot === 'statusbar' ? statusbar : watchDock;
      if (on) {
        if (el.parentElement === host) return;
        if (slot === 'formulabar') {
          host.insertBefore(el, grid);
        } else if (slot === 'statusbar') {
          if (watchDock.parentElement === host) host.insertBefore(el, watchDock);
          else host.appendChild(el);
        } else {
          host.appendChild(el);
        }
      } else if (el.parentElement === host) {
        host.removeChild(el);
      }
    };

    host.appendChild(grid);
    setChromeAttached('formulabar', flags.formulaBar);
    setChromeAttached('statusbar', flags.statusBar);
    setChromeAttached('watchDock', flags.watchWindow);

    // Track ownership before seeding — only owned (default-created) workbooks
    // should be touched by `seed`. Pre-loaded workbooks are the consumer's
    // data and must not be overwritten by the demo helper.
    let ownsWb = !opts.workbook;
    let wb: WorkbookHandle = opts.workbook ?? (await WorkbookHandle.createDefault());
    if (opts.seed && ownsWb) opts.seed(wb);

    const store = createSpreadsheetStore();
    if (opts.theme) mutators.setTheme(store, opts.theme);

    // Unified undo/redo. Attach BEFORE seed-cell hydration so the seed itself
    // doesn't pollute the stack — but seed runs above on the wb. Clear the
    // stack after attach to drop any pre-attach entries (none expected, but
    // cheap insurance).
    const history = new History();
    wb.attachHistory(history);
    history.clear();

    // Hydrate cells from engine
    mutators.replaceCells(store, wb.cells(store.getState().data.sheetIndex));
    hydrateLayoutFromEngine(wb, store, store.getState().data.sheetIndex);
    hydrateCommentsAndHyperlinksFromEngine(wb, store, store.getState().data.sheetIndex);
    hydrateMergesFromEngine(wb, store, store.getState().data.sheetIndex);
    hydrateValidationsFromEngine(wb, store, store.getState().data.sheetIndex);
    hydrateCellFormatsFromEngine(wb, store, store.getState().data.sheetIndex);
    hydrateProtectionFromEngine(wb, store);
    dispatchPassthroughSummary();

    function dispatchPassthroughSummary(): void {
      // Surface non-rendered OOXML parts (charts/drawings/pivots) and Excel
      //  Tables as host events so chrome (status bar, toast) can show a
      //  read-only badge. The core itself does not paint these; users open
      //  the .xlsx in Excel for full editing.
      const passthroughs = summarizePassthroughs(wb);
      const tables = summarizeTables(wb);
      host.dispatchEvent(new CustomEvent('fc:passthroughs', { detail: passthroughs }));
      host.dispatchEvent(new CustomEvent('fc:tables', { detail: tables }));
    }

    const cellRegistry = new CellRegistry();
    const renderer = new GridRenderer({
      host: grid,
      canvas,
      getState: () => store.getState(),
      getTheme: () => resolveTheme(host),
      getWb: () => wb,
      getDisplay: (addr, value, formula, format) =>
        cellRegistry.resolveDisplay({ addr, value, formula, format }),
    });
    cellRegistry.subscribe(() => renderer.invalidate());
    renderer.resize();

    // Always-on host features — not toggleable via `MountOptions.features`.
    const iterativeDialog = attachIterativeDialog({ host, getWb: () => wb, strings });
    const externalLinksDialog = attachExternalLinksDialog({
      host,
      getWb: () => wb,
      strings,
    });
    const cellStylesGallery = attachCellStylesGallery({
      host,
      store,
      history,
      getWb: () => wb,
    });

    // Toggleable host-level features. Each binding starts `null` and is
    // populated by its attacher when the corresponding flag is on. Cross-
    // feature references (e.g. `onHostKey` checking `formatPainter`) read
    // the let-bound vars directly so they always see the current value.
    let formatDialog: ReturnType<typeof attachFormatDialog> | null = null;
    let formatPainter: FormatPainterHandle | null = null;
    let hover: ReturnType<typeof attachHover> | null = null;
    let conditionalDialog: ReturnType<typeof attachConditionalDialog> | null = null;
    let goToDialog: ReturnType<typeof attachGoToDialog> | null = null;
    let fxDialog: FxDialogHandle | null = null;
    let fxClickHandler: (() => void) | null = null;
    let namedRangeDialog: ReturnType<typeof attachNamedRangeDialog> | null = null;
    let pageSetupDialog: ReturnType<typeof attachPageSetupDialog> | null = null;
    let hyperlinkDialog: ReturnType<typeof attachHyperlinkDialog> | null = null;
    let statusBar: ReturnType<typeof attachStatusBar> | null = null;
    let watchPanel: ReturnType<typeof attachWatchPanel> | null = null;
    let unsubWatchRecalc: () => void = (): void => {};
    let unsubWatchWb: () => void = (): void => {};
    let slicer: SlicerHandle | null = null;
    let unsubSlicerRecalc: () => void = (): void => {};
    let unsubSlicerWb: () => void = (): void => {};
    let errorMenu: ErrorMenuHandle | null = null;
    let detachWheel: () => void = (): void => {};
    type AutocompleteHandle = ReturnType<typeof attachAutocomplete>;
    const autocompleteStub: AutocompleteHandle = {
      isOpen: () => false,
      move: (_n: number) => {},
      acceptHighlighted: () => false,
      close: () => {},
      refresh: () => {},
      detach: () => {},
    };
    let fxAutocomplete: AutocompleteHandle = autocompleteStub;
    let fxArgHelper: ReturnType<typeof attachArgHelper> | null = null;
    let hostShortcutsAttached = false;
    let canvasErrorClickAttached = false;

    // Built-in feature registry — keyed by feature id, populated by
    // attachers, drained by `dispose()`. Mirrored into `featuresView`
    // below for the public `instance.features` surface.
    const featureRegistry = new Map<string, ExtensionHandle>();
    const wrapHandle = (raw: unknown, detach: () => void): ExtensionHandle => {
      const h = (
        raw && typeof raw === 'object' ? (raw as Record<string, unknown>) : {}
      ) as ExtensionHandle;
      h.dispose = detach;
      return h;
    };

    // Host-level toggleable feature ids — `setFeatures(next)` walks this
    // list and dispatches to the attach/detach helpers. `shortcuts` is
    // listed for the host-level meta-key listener; the wb-side keyboard
    // attach is rebuilt by `bindEngine` when shortcuts flip.
    const HOST_TOGGLEABLE_IDS = [
      'formatDialog',
      'formatPainter',
      'hoverComment',
      'conditional',
      'gotoSpecial',
      'fxDialog',
      'namedRanges',
      'pageSetup',
      'hyperlink',
      'statusBar',
      'watchWindow',
      'slicer',
      'errorIndicators',
      'autocomplete',
      'wheel',
      'shortcuts',
      'formulaBar',
    ] as const;

    // Wb-side toggleable feature ids — these live inside `bindEngine`.
    // Toggling any of them triggers a binding rebuild from `setFeatures`.
    const WB_TOGGLEABLE_IDS = [
      'shortcuts',
      'clipboard',
      'pasteSpecial',
      'contextMenu',
      'findReplace',
      'validation',
    ] as const;

    const attachHostFeature = (id: string): void => {
      switch (id) {
        case 'formatDialog':
          if (formatDialog) return;
          formatDialog = attachFormatDialog({ host, store, strings, history, getWb: () => wb });
          featureRegistry.set(
            'formatDialog',
            wrapHandle(formatDialog, () => formatDialog?.detach()),
          );
          break;
        case 'formatPainter':
          if (formatPainter) return;
          formatPainter = attachFormatPainter({ host, store, history });
          featureRegistry.set(
            'formatPainter',
            wrapHandle(formatPainter, () => formatPainter?.detach()),
          );
          break;
        case 'hoverComment':
          if (hover) return;
          hover = attachHover({ grid, store });
          featureRegistry.set(
            'hoverComment',
            wrapHandle(hover, () => hover?.detach()),
          );
          break;
        case 'conditional':
          if (conditionalDialog) return;
          conditionalDialog = attachConditionalDialog({ host, store, strings });
          featureRegistry.set(
            'conditional',
            wrapHandle(conditionalDialog, () => conditionalDialog?.detach()),
          );
          break;
        case 'gotoSpecial':
          if (goToDialog) return;
          goToDialog = attachGoToDialog({ host, store, strings, getWb: () => wb });
          featureRegistry.set(
            'gotoSpecial',
            wrapHandle(goToDialog, () => goToDialog?.detach()),
          );
          break;
        case 'fxDialog':
          if (fxDialog) return;
          fxDialog = attachFxDialog({
            host,
            store,
            strings,
            onInsert: (formula) => {
              fxInput.value = formula;
              fxInput.focus();
              commitFx('none');
            },
          });
          fxClickHandler = (): void => fxDialog?.open();
          fx.addEventListener('click', fxClickHandler);
          fx.disabled = false;
          fx.style.cursor = '';
          featureRegistry.set(
            'fxDialog',
            wrapHandle(fxDialog, () => fxDialog?.detach()),
          );
          break;
        case 'namedRanges':
          if (namedRangeDialog) return;
          namedRangeDialog = attachNamedRangeDialog({ host, wb, strings });
          featureRegistry.set(
            'namedRanges',
            wrapHandle(namedRangeDialog, () => namedRangeDialog?.detach()),
          );
          break;
        case 'pageSetup':
          if (pageSetupDialog) return;
          pageSetupDialog = attachPageSetupDialog({ host, store, strings, history });
          featureRegistry.set(
            'pageSetup',
            wrapHandle(pageSetupDialog, () => pageSetupDialog?.detach()),
          );
          break;
        case 'hyperlink':
          if (hyperlinkDialog) return;
          hyperlinkDialog = attachHyperlinkDialog({
            host,
            store,
            strings,
            history,
            getWb: () => wb,
          });
          featureRegistry.set(
            'hyperlink',
            wrapHandle(hyperlinkDialog, () => hyperlinkDialog?.detach()),
          );
          break;
        case 'statusBar':
          if (statusBar) return;
          setChromeAttached('statusbar', true);
          statusBar = attachStatusBar({
            statusbar,
            store,
            strings,
            getEngineLabel: () => (wb.isStub ? 'stub' : `formulon ${wb.version}`),
            getCalcMode: () => wb.calcMode(),
            onCycleCalcMode: () => {
              const cur = wb.calcMode();
              if (cur === null) return;
              const next = ((cur + 1) % 3) as 0 | 1 | 2;
              wb.setCalcMode(next);
              statusBar?.refresh();
            },
            onRecalc: () => {
              wb.recalc();
              mutators.replaceCells(store, wb.cells(store.getState().data.sheetIndex));
              renderer.invalidate();
            },
          });
          featureRegistry.set(
            'statusBar',
            wrapHandle(statusBar, () => statusBar?.detach()),
          );
          break;
        case 'watchWindow':
          if (watchPanel) return;
          setChromeAttached('watchDock', true);
          watchPanel = attachWatchPanel({ host: watchDock, store, getWb: () => wb, strings });
          featureRegistry.set(
            'watchWindow',
            wrapHandle(watchPanel, () => watchPanel?.detach()),
          );
          unsubWatchRecalc = emitter.on('recalc', () => watchPanel?.refresh());
          unsubWatchWb = emitter.on('workbookChange', () => watchPanel?.refresh());
          break;
        case 'slicer':
          if (slicer) return;
          slicer = attachSlicer({ host, store, getWb: () => wb, history, strings });
          featureRegistry.set(
            'slicer',
            wrapHandle(slicer, () => slicer?.detach()),
          );
          unsubSlicerRecalc = emitter.on('recalc', () => slicer?.refresh());
          unsubSlicerWb = emitter.on('workbookChange', () => slicer?.refresh());
          break;
        case 'errorIndicators':
          if (errorMenu) return;
          errorMenu = attachErrorMenu({
            host,
            store,
            getWb: () => wb,
            strings,
            onEditCell: (addr) => {
              mutators.setActive(store, addr);
              const cell = store.getState().data.cells.get(`${addr.sheet}:${addr.row}:${addr.col}`);
              fxInput.value = formatCellForEdit(cell, wb, addr);
              fxInput.focus();
              fxInput.setSelectionRange(fxInput.value.length, fxInput.value.length);
            },
          });
          if (!canvasErrorClickAttached) {
            canvas.addEventListener('click', onCanvasClick);
            canvasErrorClickAttached = true;
          }
          featureRegistry.set(
            'errorIndicators',
            wrapHandle(errorMenu, () => errorMenu?.detach()),
          );
          break;
        case 'autocomplete':
          if (fxAutocomplete !== autocompleteStub) return;
          fxAutocomplete = attachAutocomplete({
            input: fxInput,
            onAfterInsert: () => syncFxRefs(),
            getTables: () => wb.getTables(),
            getCustomFunctions: () => formulaRegistry.list(),
            getFunctionNames: () => wb.functionNames(),
          });
          fxArgHelper = attachArgHelper({ input: fxInput });
          break;
        case 'wheel':
          // Idempotent — `detachWheel` is replaced wholesale.
          detachWheel();
          detachWheel = attachWheel({ grid, store, wb });
          break;
        case 'shortcuts':
          if (hostShortcutsAttached) return;
          host.addEventListener('keydown', onHostKey);
          hostShortcutsAttached = true;
          break;
        case 'formulaBar':
          setChromeAttached('formulabar', true);
          break;
      }
      refreshFeaturesView();
    };

    const detachHostFeature = (id: string): void => {
      switch (id) {
        case 'formatDialog':
          formatDialog?.detach();
          formatDialog = null;
          featureRegistry.delete('formatDialog');
          break;
        case 'formatPainter':
          formatPainter?.detach();
          formatPainter = null;
          featureRegistry.delete('formatPainter');
          break;
        case 'hoverComment':
          hover?.detach();
          hover = null;
          featureRegistry.delete('hoverComment');
          break;
        case 'conditional':
          conditionalDialog?.detach();
          conditionalDialog = null;
          featureRegistry.delete('conditional');
          break;
        case 'gotoSpecial':
          goToDialog?.detach();
          goToDialog = null;
          featureRegistry.delete('gotoSpecial');
          break;
        case 'fxDialog':
          if (fxClickHandler) fx.removeEventListener('click', fxClickHandler);
          fxClickHandler = null;
          fxDialog?.detach();
          fxDialog = null;
          fx.disabled = true;
          fx.style.cursor = 'default';
          featureRegistry.delete('fxDialog');
          break;
        case 'namedRanges':
          namedRangeDialog?.detach();
          namedRangeDialog = null;
          featureRegistry.delete('namedRanges');
          break;
        case 'pageSetup':
          pageSetupDialog?.detach();
          pageSetupDialog = null;
          featureRegistry.delete('pageSetup');
          break;
        case 'hyperlink':
          hyperlinkDialog?.detach();
          hyperlinkDialog = null;
          featureRegistry.delete('hyperlink');
          break;
        case 'statusBar':
          statusBar?.detach();
          statusBar = null;
          featureRegistry.delete('statusBar');
          setChromeAttached('statusbar', false);
          break;
        case 'watchWindow':
          unsubWatchRecalc();
          unsubWatchWb();
          unsubWatchRecalc = (): void => {};
          unsubWatchWb = (): void => {};
          watchPanel?.detach();
          watchPanel = null;
          featureRegistry.delete('watchWindow');
          setChromeAttached('watchDock', false);
          break;
        case 'slicer':
          unsubSlicerRecalc();
          unsubSlicerWb();
          unsubSlicerRecalc = (): void => {};
          unsubSlicerWb = (): void => {};
          slicer?.detach();
          slicer = null;
          featureRegistry.delete('slicer');
          break;
        case 'errorIndicators':
          if (canvasErrorClickAttached) canvas.removeEventListener('click', onCanvasClick);
          canvasErrorClickAttached = false;
          errorMenu?.detach();
          errorMenu = null;
          featureRegistry.delete('errorIndicators');
          break;
        case 'autocomplete':
          fxAutocomplete.detach();
          fxAutocomplete = autocompleteStub;
          fxArgHelper?.detach();
          fxArgHelper = null;
          break;
        case 'wheel':
          detachWheel();
          detachWheel = (): void => {};
          break;
        case 'shortcuts':
          if (hostShortcutsAttached) host.removeEventListener('keydown', onHostKey);
          hostShortcutsAttached = false;
          break;
        case 'formulaBar':
          setChromeAttached('formulabar', false);
          break;
      }
      refreshFeaturesView();
    };

    // wb-dependent layer — re-built whenever setWorkbook swaps the engine.
    interface EngineBinding {
      editor: InlineEditor;
      pasteSpecialDialog: ReturnType<typeof attachPasteSpecial> | null;
      findReplace: ReturnType<typeof attachFindReplace> | null;
      validation: ReturnType<typeof attachValidationList> | null;
      clipboardH: ReturnType<typeof attachClipboard> | null;
      unbind: () => void;
    }

    const bindEngine = (currentWb: WorkbookHandle): EngineBinding => {
      const editor = new InlineEditor({
        host,
        grid,
        store,
        wb: currentWb,
        onAfterCommit: () => {
          mutators.replaceCells(store, currentWb.cells(store.getState().data.sheetIndex));
        },
      });
      const detachPtr = attachPointer(
        grid,
        store,
        currentWb,
        () => mutators.replaceCells(store, currentWb.cells(store.getState().data.sheetIndex)),
        history,
        () =>
          editor.isActive() && editor.isFormulaEdit()
            ? {
                isFormulaEdit: () => editor.isFormulaEdit(),
                insertRefAtCaret: (ref) => editor.insertRefAtCaret(ref),
              }
            : null,
      );
      const detachKey = flags.shortcuts
        ? attachKeyboard({
            host,
            store,
            wb: currentWb,
            history,
            onBeginEdit: (seed) => editor.begin(seed),
            onClearActive: () => {
              mutators.replaceCells(store, currentWb.cells(store.getState().data.sheetIndex));
              updateChrome();
            },
            onAfterHistory: () =>
              mutators.replaceCells(store, currentWb.cells(store.getState().data.sheetIndex)),
            onGoTo: () => {
              // F5 / Ctrl+G — open Go To Special when the feature is on,
              // otherwise fall back to focusing the Name Box for direct
              // navigation typing.
              if (goToDialog) {
                goToDialog.open();
                return;
              }
              tag.focus();
              tag.select();
            },
          })
        : (): void => {};
      const clipboardH = flags.clipboard
        ? attachClipboard({
            host,
            store,
            wb: currentWb,
            onAfterCommit: () =>
              mutators.replaceCells(store, currentWb.cells(store.getState().data.sheetIndex)),
          })
        : null;
      const pasteSpecialDialog =
        flags.pasteSpecial && clipboardH
          ? attachPasteSpecial({
              host,
              store,
              wb: currentWb,
              strings,
              history,
              getSnapshot: () => clipboardH.getSnapshot(),
              onAfterCommit: () =>
                mutators.replaceCells(store, currentWb.cells(store.getState().data.sheetIndex)),
            })
          : null;
      const detachContextMenu = flags.contextMenu
        ? attachContextMenu({
            host,
            store,
            wb: currentWb,
            strings,
            history,
            onAfterCommit: () =>
              mutators.replaceCells(store, currentWb.cells(store.getState().data.sheetIndex)),
            onFormatDialog: () => formatDialog?.open(),
            onPasteSpecial: () => pasteSpecialDialog?.open(),
            onInsertHyperlink: () => hyperlinkDialog?.open(),
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
      const findReplace = flags.findReplace
        ? attachFindReplace({
            host,
            store,
            wb: currentWb,
            strings,
            onAfterCommit: () =>
              mutators.replaceCells(store, currentWb.cells(store.getState().data.sheetIndex)),
          })
        : null;
      const validation = flags.validation
        ? attachValidationList({
            grid,
            store,
            wb: currentWb,
            onAfterCommit: () =>
              mutators.replaceCells(store, currentWb.cells(store.getState().data.sheetIndex)),
          })
        : null;

      const onDblClick = (e: MouseEvent): void => {
        if (e.button !== 0) return;
        if (editor.isActive()) return;
        if (formatPainter?.isActive()) return;
        const s = store.getState();
        const a = s.selection.active;
        const seed =
          currentWb.cellFormula(a) ??
          formatCellForEdit(s.data.cells.get(`${a.sheet}:${a.row}:${a.col}`), currentWb, a);
        editor.begin(seed);
        e.preventDefault();
      };
      grid.addEventListener('dblclick', onDblClick);

      const unsubWb = currentWb.subscribe((e: ChangeEvent) => {
        if (e.kind === 'value') {
          const formula = currentWb.cellFormula(e.addr);
          const cell = { value: e.next, formula };
          store.setState((s) => {
            const cells = new Map(s.data.cells);
            cells.set(`${e.addr.sheet}:${e.addr.row}:${e.addr.col}`, cell);
            return { ...s, data: { ...s.data, cells } };
          });
          emitter.emit('cellChange', { addr: e.addr, value: e.next, formula });
        } else if (e.kind === 'recalc') {
          emitter.emit('recalc', { dirty: e.dirty });
        }
      });

      return {
        editor,
        pasteSpecialDialog,
        findReplace,
        validation,
        clipboardH,
        unbind: () => {
          detachPtr();
          detachKey();
          clipboardH?.detach();
          detachContextMenu();
          findReplace?.detach();
          pasteSpecialDialog?.detach();
          validation?.detach();
          grid.removeEventListener('dblclick', onDblClick);
          unsubWb();
          if (editor.isActive()) editor.cancel();
        },
      };
    };

    let binding = bindEngine(wb);

    // Top-level shortcuts that need to beat the browser default — Cmd+F opens
    // Find/Replace, Cmd+A selects all cells, Cmd+1 opens Format Cells. Bound on
    // the host so they only trigger while the spreadsheet has focus
    // (formula-bar / find / dialog inputs keep their browser-native behavior).
    // Skipped entirely when `flags.shortcuts === false`.
    const onHostKey = (e: KeyboardEvent): void => {
      const meta = e.ctrlKey || e.metaKey;
      // F9 / Ctrl+Alt+F9 — full recalc. Mirrors Excel: F9 alone in
      // manual mode kicks a recalc, Ctrl+Alt+F9 forces re-evaluation
      // even on cells the engine considers clean. The cell engine
      // doesn't distinguish these — both call wb.recalc() — so the
      // shortcut serves the same intent under either binding.
      if (e.key === 'F9') {
        e.preventDefault();
        wb.recalc();
        mutators.replaceCells(store, wb.cells(store.getState().data.sheetIndex));
        renderer.invalidate();
        return;
      }
      if (!meta) return;
      const k = e.key.toLowerCase();
      if (e.shiftKey && k === 'c') {
        // Cmd/Ctrl+Shift+C — copy formatting (one-shot).
        if (!formatPainter) return;
        e.preventDefault();
        formatPainter.activate(false);
        return;
      }
      if (e.shiftKey && k === 'v') {
        // Cmd/Ctrl+Shift+V — open Paste Special.
        if (!binding.pasteSpecialDialog) return;
        e.preventDefault();
        binding.pasteSpecialDialog.open();
        return;
      }
      if (e.altKey && k === 'v') {
        // Excel alt-binding: Ctrl+Alt+V (Win) / Cmd+Option+V (Mac).
        if (!binding.pasteSpecialDialog) return;
        e.preventDefault();
        binding.pasteSpecialDialog.open();
        return;
      }
      if (k === 'f') {
        if (!binding.findReplace) return;
        e.preventDefault();
        binding.findReplace.open();
      } else if (k === 'k') {
        // Ctrl/Cmd+K — Insert Hyperlink dialog (Excel/Sheets parity).
        if (!hyperlinkDialog) return;
        e.preventDefault();
        hyperlinkDialog.open();
      } else if (k === 'a') {
        e.preventDefault();
        mutators.selectAll(store);
      } else if (e.key === '1') {
        if (!formatDialog) return;
        e.preventDefault();
        formatDialog.open();
      } else if (e.key === '`') {
        // Ctrl+` — toggle show-formulas mode.
        e.preventDefault();
        mutators.setShowFormulas(store, !store.getState().ui.showFormulas);
      } else if (e.altKey && k === 'r') {
        // Ctrl/Cmd+Alt+R — toggle R1C1 reference style. Mirrors Excel's
        //  File → Options → Formulas → "Use R1C1 reference style" checkbox
        //  but exposed as a shortcut for power users.
        e.preventDefault();
        mutators.setR1C1(store, !store.getState().ui.r1c1);
      } else if (e.key === ';') {
        // Ctrl+; — insert today's date as Excel serial.
        e.preventDefault();
        const now = new Date();
        const utcMs = Date.UTC(now.getFullYear(), now.getMonth(), now.getDate());
        const serial = utcMs / 86_400_000 + 25569;
        const a = store.getState().selection.active;
        wb.setNumber(a, Math.floor(serial));
        mutators.replaceCells(store, wb.cells(store.getState().data.sheetIndex));
      } else if (e.shiftKey && e.key === ':') {
        // Ctrl+Shift+: — insert current time fraction.
        e.preventDefault();
        const now = new Date();
        const frac =
          (now.getUTCHours() * 3600 + now.getUTCMinutes() * 60 + now.getUTCSeconds()) / 86400;
        const a = store.getState().selection.active;
        wb.setNumber(a, frac);
        mutators.replaceCells(store, wb.cells(store.getState().data.sheetIndex));
      } else if (k === 'd') {
        // Ctrl+D — fill down from the top row of the selection to the rest.
        e.preventDefault();
        const r = store.getState().selection.range;
        if (r.r1 > r.r0) {
          fillRange(
            store.getState(),
            wb,
            { sheet: r.sheet, r0: r.r0, c0: r.c0, r1: r.r0, c1: r.c1 },
            r,
          );
          mutators.replaceCells(store, wb.cells(store.getState().data.sheetIndex));
        }
      } else if (k === 'r') {
        // Ctrl+R — fill right from the left column of the selection.
        e.preventDefault();
        const r = store.getState().selection.range;
        if (r.c1 > r.c0) {
          fillRange(
            store.getState(),
            wb,
            { sheet: r.sheet, r0: r.r0, c0: r.c0, r1: r.r1, c1: r.c0 },
            r,
          );
          mutators.replaceCells(store, wb.cells(store.getState().data.sheetIndex));
        }
      } else if (k === 'b') {
        e.preventDefault();
        recordFormatChange(history, store, () => {
          toggleBold(store.getState(), store);
        });
        flushFormatToEngine(wb, store, store.getState().data.sheetIndex);
      } else if (k === 'i') {
        e.preventDefault();
        recordFormatChange(history, store, () => {
          toggleItalic(store.getState(), store);
        });
        flushFormatToEngine(wb, store, store.getState().data.sheetIndex);
      } else if (k === 'u') {
        e.preventDefault();
        recordFormatChange(history, store, () => {
          toggleUnderline(store.getState(), store);
        });
        flushFormatToEngine(wb, store, store.getState().data.sheetIndex);
      } else if (e.key === '5') {
        e.preventDefault();
        recordFormatChange(history, store, () => {
          toggleStrike(store.getState(), store);
        });
        flushFormatToEngine(wb, store, store.getState().data.sheetIndex);
      }
    };
    // Error / validation triangle clicks. Bound on the canvas by the
    // `errorIndicators` attacher above. We use `click` (not `pointerdown`)
    // so the existing pointer handler gets to set the active cell first —
    // that way the menu and the cell select agree on the addr the user
    // just clicked.
    const onCanvasClick = (e: MouseEvent): void => {
      if (!errorMenu) return;
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
        errorMenu.open(hit.addr, e.clientX, e.clientY, hit.kind);
        return;
      }
    };

    // Formula bar editing — typing in the formula bar edits the active cell.
    let fxEditing = false;
    let fxBaseline = '';
    const syncFxRefs = (): void => {
      const refs = extractRefs(fxInput.value).map((r) => ({
        r0: r.r0,
        c0: r.c0,
        r1: r.r1,
        c1: r.c1,
        colorIndex: r.colorIndex,
      }));
      mutators.setEditorRefs(store, refs);
    };
    const clearFxRefs = (): void => mutators.setEditorRefs(store, []);
    const onFxFocus = (): void => {
      if (binding.editor.isActive()) binding.editor.cancel();
      fxEditing = true;
      fxBaseline = fxInput.value;
      syncFxRefs();
    };
    const onFxInput = (): void => {
      if (fxEditing) syncFxRefs();
      fxAutocomplete.refresh();
      fxArgHelper?.refresh();
    };
    const onFxKeyUp = (): void => {
      // Caret moves on arrow / Home / End / click — refresh the arg tooltip
      //  alone (autocomplete already keys off `input`).
      if (fxEditing) fxArgHelper?.refresh();
    };
    const onFxKey = (e: KeyboardEvent): void => {
      // The formula bar lives inside `host`, so its key events bubble to the
      // grid's keyboard handler. Stop propagation on the keys we handle so the
      // grid handler doesn't interpret Enter/Tab as begin-edit / move-active.
      if (fxAutocomplete.isOpen()) {
        if (e.key === 'ArrowDown') {
          e.preventDefault();
          e.stopPropagation();
          fxAutocomplete.move(1);
          return;
        }
        if (e.key === 'ArrowUp') {
          e.preventDefault();
          e.stopPropagation();
          fxAutocomplete.move(-1);
          return;
        }
        if ((e.key === 'Enter' || e.key === 'Tab') && fxAutocomplete.acceptHighlighted()) {
          e.preventDefault();
          e.stopPropagation();
          return;
        }
        if (e.key === 'Escape') {
          e.preventDefault();
          e.stopPropagation();
          fxAutocomplete.close();
          return;
        }
      }
      if (e.key === 'Enter') {
        // Excel: Alt+Enter inserts a newline (multi-line cell content); plain
        // Enter commits and advances. Shift+Enter mirrors Alt+Enter for users
        // expecting browser-textarea behavior.
        if (e.altKey || e.shiftKey) {
          e.stopPropagation();
          return;
        }
        e.preventDefault();
        e.stopPropagation();
        commitFx('down');
      } else if (e.key === 'Tab') {
        e.preventDefault();
        e.stopPropagation();
        commitFx(e.shiftKey ? 'none' : 'right');
      } else if (e.key === 'Escape') {
        e.preventDefault();
        e.stopPropagation();
        fxInput.value = fxBaseline;
        fxEditing = false;
        host.focus();
        updateChrome();
      } else if (e.key === 'F4') {
        e.preventDefault();
        e.stopPropagation();
        const caret = fxInput.selectionStart ?? fxInput.value.length;
        const r = rotateRefAt(fxInput.value, caret);
        if (r.text !== fxInput.value) {
          fxInput.value = r.text;
          fxInput.setSelectionRange(r.caret, r.caret);
          syncFxRefs();
        }
      }
    };
    const onFxBlur = (): void => {
      clearFxRefs();
      fxAutocomplete.close();
      if (!fxEditing) return;
      // Only commit if value actually changed.
      if (fxInput.value !== fxBaseline) commitFx('none');
      else fxEditing = false;
    };
    function commitFx(advance: 'down' | 'right' | 'none'): void {
      const s = store.getState();
      const a = s.selection.active;
      try {
        const fmt = s.format.formats.get(`${a.sheet}:${a.row}:${a.col}`);
        const outcome = writeInputValidated(wb, a, fxInput.value, fmt?.validation);
        if (!outcome.ok) {
          console.warn(`formulon-cell: validation ${outcome.severity}: ${outcome.message}`);
          if (outcome.severity === 'stop') {
            // Keep editing — don't advance, don't commit baseline change.
            fxInput.focus();
            return;
          }
        }
      } catch (err) {
        console.warn('formulon-cell: writeInput failed', err);
      }
      mutators.replaceCells(store, wb.cells(store.getState().data.sheetIndex));
      fxEditing = false;
      clearFxRefs();
      if (advance === 'down') {
        mutators.setActive(store, { ...a, row: a.row + 1 });
      } else if (advance === 'right') {
        mutators.setActive(store, { ...a, col: a.col + 1 });
      }
      host.focus();
    }
    fxInput.addEventListener('focus', onFxFocus);
    fxInput.addEventListener('input', onFxInput);
    fxInput.addEventListener('keyup', onFxKeyUp);
    fxInput.addEventListener('keydown', onFxKey);
    fxInput.addEventListener('blur', onFxBlur);

    // Name box — Enter jumps to a cell ref, Escape reverts.
    const onTagFocus = (): void => tag.select();
    const onTagKey = (e: KeyboardEvent): void => {
      // Same caveat as onFxKey — stop propagation so the grid's keyboard
      // handler doesn't catch our Enter/Escape.
      if (e.key === 'Enter') {
        e.preventDefault();
        e.stopPropagation();
        const sheetIdx = store.getState().data.sheetIndex;
        // Try range first (A1:B5), fall back to single ref, then defined name.
        const range = parseRangeRef(tag.value);
        if (range) {
          store.setState((s) => ({
            ...s,
            selection: {
              active: { sheet: sheetIdx, row: range.r0, col: range.c0 },
              anchor: { sheet: sheetIdx, row: range.r0, col: range.c0 },
              range: { sheet: sheetIdx, ...range },
            },
          }));
          host.focus();
          return;
        }
        const parsed = parseCellRef(tag.value);
        if (parsed) {
          mutators.setActive(store, {
            sheet: sheetIdx,
            row: parsed.row,
            col: parsed.col,
          });
          host.focus();
          return;
        }
        // Defined-name lookup (engine-side, RO).
        const dn = lookupDefinedName(wb, tag.value.trim());
        if (dn) {
          const sub = parseRangeRef(dn) ?? parseCellRef(dn);
          if (sub) {
            if ('r0' in sub) {
              store.setState((s) => ({
                ...s,
                selection: {
                  active: { sheet: sheetIdx, row: sub.r0, col: sub.c0 },
                  anchor: { sheet: sheetIdx, row: sub.r0, col: sub.c0 },
                  range: { sheet: sheetIdx, ...sub },
                },
              }));
            } else {
              mutators.setActive(store, { sheet: sheetIdx, row: sub.row, col: sub.col });
            }
            host.focus();
            return;
          }
        }
      } else if (e.key === 'Escape') {
        e.preventDefault();
        e.stopPropagation();
        host.focus();
        updateChrome();
      }
    };
    const onTagBlur = (): void => {
      // Revert to current selection when leaving without committing.
      updateChrome();
    };
    tag.addEventListener('focus', onTagFocus);
    tag.addEventListener('keydown', onTagKey);
    tag.addEventListener('blur', onTagBlur);

    // Re-paint and update chrome on every store change.
    let lastSheetIdx = store.getState().data.sheetIndex;
    let lastSelection = store.getState().selection;
    const unsub = store.subscribe(() => {
      const s = store.getState();
      if (s.data.sheetIndex !== lastSheetIdx) {
        wb.clearViewportHint();
        lastSheetIdx = s.data.sheetIndex;
      }
      if (!selectionEquals(lastSelection, s.selection)) {
        lastSelection = s.selection;
        emitter.emit('selectionChange', {
          active: s.selection.active,
          anchor: s.selection.anchor,
          range: s.selection.range,
        });
      }
      renderer.invalidate();
      updateChrome();
    });

    function updateChrome(): void {
      const s = store.getState();
      const a = s.selection.active;
      const colLetter = ((): string => {
        let n = a.col;
        let out = '';
        do {
          out = String.fromCharCode(65 + (n % 26)) + out;
          n = Math.floor(n / 26) - 1;
        } while (n >= 0);
        return out;
      })();
      const ref = s.ui.r1c1 ? `R${a.row + 1}C${a.col + 1}` : `${colLetter}${a.row + 1}`;
      // Don't stomp the user's in-progress name-box typing.
      if (document.activeElement !== tag) tag.value = ref;
      const cell = s.data.cells.get(`${a.sheet}:${a.row}:${a.col}`);
      const formula = cell?.formula ?? '';
      let display = '';
      if (formula) display = formula;
      else if (cell) {
        const v = cell.value;
        switch (v.kind) {
          case 'number':
            display = String(v.value);
            break;
          case 'bool':
            display = v.value ? 'TRUE' : 'FALSE';
            break;
          case 'text':
            display = v.value;
            break;
          case 'error':
            display = v.text;
            break;
          default: {
            // Lambda values fall through here (no `lambda` kind on
            // CellValue yet) — render their body as `=LAMBDA(...)` so the
            // formula bar shows something editable instead of blank.
            const lambda = wb.getLambdaText(a);
            display = lambda ? `=${lambda}` : '';
            break;
          }
        }
      }
      // Don't stomp on the user's in-progress formula bar typing.
      if (!fxEditing) fxInput.value = display;
      a11y.textContent = `${ref} ${display}`;
    }
    updateChrome();

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
    if (opts.extensions) {
      const sorted = sortByPriority(dedupeById(flattenExtensions(opts.extensions)));
      for (const ext of sorted) mountExtension(ext);
    }

    // Combined view exposed on `instance.features` — built-ins + user.
    const featuresView: Record<string, ExtensionHandle | undefined> = {};
    const refreshFeaturesView = (): void => {
      for (const k of Object.keys(featuresView)) delete featuresView[k];
      for (const [k, v] of featureRegistry) featuresView[k] = v;
      for (const [k, v] of userHandles) featuresView[k] = v;
    };
    refreshFeaturesView();

    // Initial host-feature attach — runs after every helper closure
    // (`onCanvasClick`, `onHostKey`, `syncFxRefs`, `commitFx`) is in
    // scope so the attacher bodies can resolve them at call time.
    for (const id of HOST_TOGGLEABLE_IDS) {
      if (flags[id as keyof typeof flags]) attachHostFeature(id);
    }

    // Locale change → snapshot strings into the local `let` so future
    // attaches see the latest. Built-in attaches in v0.1 capture strings at
    // attach time; full live updates land in v0.2 once each `attach*` ships
    // a `setStrings` hook.
    const unsubI18n = i18n.subscribe((next) => {
      strings = next;
      // Notify user extensions that opt into live updates.
      for (const handle of userHandles.values()) handle.setStrings?.(next);
      for (const handle of featureRegistry.values()) handle.setStrings?.(next);
      emitter.emit('localeChange', { locale: i18n.locale, strings: next });
    });

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
        return formatPainter ?? undefined;
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
        const nextFlags = resolveFlags(next);
        // Diff host-level features and dispatch attach/detach.
        for (const id of HOST_TOGGLEABLE_IDS) {
          const k = id as keyof typeof flags;
          const was = flags[k];
          const now = nextFlags[k];
          if (was === now) continue;
          if (was && !now) detachHostFeature(id);
          else if (!was && now) attachHostFeature(id);
        }
        // Wb-side rebuild only when a wb-bound feature flipped — keeps the
        // editor / pointer / undo state intact when only host-level flags
        // change.
        const wbChanged = WB_TOGGLEABLE_IDS.some(
          (id) => flags[id as keyof typeof flags] !== nextFlags[id as keyof typeof nextFlags],
        );
        flags = nextFlags;
        if (wbChanged) {
          binding.unbind();
          binding = bindEngine(wb);
        }
        refreshFeaturesView();
      },
      setExtensions(next) {
        // Dispose all currently-mounted user extensions, then re-mount the
        // new list. Built-ins are untouched — use `setFeatures` for those.
        for (const handle of userHandles.values()) handle.dispose();
        userHandles.clear();
        if (next && next.length) {
          const sorted = sortByPriority(dedupeById(flattenExtensions(next)));
          for (const ext of sorted) mountExtension(ext);
        }
        refreshFeaturesView();
      },
      openConditionalDialog() {
        conditionalDialog?.open();
      },
      openIterativeDialog() {
        iterativeDialog.open();
      },
      openExternalLinksDialog() {
        externalLinksDialog.open();
      },
      openCellStylesGallery() {
        cellStylesGallery.open();
      },
      openFunctionArguments(seedName?: string) {
        fxDialog?.open(seedName);
      },
      openNamedRangeDialog() {
        namedRangeDialog?.open();
      },
      openPageSetup() {
        pageSetupDialog?.open();
      },
      print() {
        // The print command is wired through the same flag as the dialog —
        // when the feature is off, both call sites are no-ops. Skip if the
        // dialog never attached so consumers can rely on the gate.
        if (!pageSetupDialog) return;
        printSheet(wb, store, store.getState().data.sheetIndex, host);
      },
      openFormatDialog() {
        formatDialog?.open();
      },
      openGoToSpecial() {
        goToDialog?.open();
      },
      openWatchWindow() {
        watchPanel?.open();
      },
      closeWatchWindow() {
        watchPanel?.close();
      },
      toggleWatchWindow() {
        watchPanel?.toggle();
      },
      addSlicer(input) {
        if (!slicer) {
          throw new Error('addSlicer: features.slicer is disabled');
        }
        return slicer.addSlicer(input);
      },
      removeSlicer(id) {
        slicer?.removeSlicer(id);
      },
      toggleSheetProtection() {
        const sheet = store.getState().data.sheetIndex;
        const on = !store.getState().protection.protectedSheets.has(sheet);
        mutators.setSheetProtected(store, sheet, on);
        flushProtectionToEngine(wb, sheet, on);
        renderer.invalidate();
      },
      setSheetProtected(on: boolean, password?: string) {
        const sheet = store.getState().data.sheetIndex;
        mutators.setSheetProtected(
          store,
          sheet,
          on,
          password !== undefined ? { password } : undefined,
        );
        flushProtectionToEngine(wb, sheet, on, password);
        renderer.invalidate();
      },
      isSheetProtected() {
        const sheet = store.getState().data.sheetIndex;
        return store.getState().protection.protectedSheets.has(sheet);
      },
      tracePrecedents() {
        const a = store.getState().selection.active;
        for (const from of findPrecedents(wb, a)) {
          mutators.addTrace(store, { kind: 'precedent', from, to: a });
        }
        renderer.invalidate();
      },
      traceDependents() {
        const a = store.getState().selection.active;
        for (const to of findDependents(wb, a)) {
          mutators.addTrace(store, { kind: 'dependent', from: a, to });
        }
        renderer.invalidate();
      },
      clearTraces() {
        mutators.clearTraces(store);
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
        mutators.replaceCells(store, wb.cells(store.getState().data.sheetIndex));
        hydrateLayoutFromEngine(wb, store, store.getState().data.sheetIndex);
        hydrateCommentsAndHyperlinksFromEngine(wb, store, store.getState().data.sheetIndex);
        hydrateMergesFromEngine(wb, store, store.getState().data.sheetIndex);
        hydrateValidationsFromEngine(wb, store, store.getState().data.sheetIndex);
        hydrateCellFormatsFromEngine(wb, store, store.getState().data.sheetIndex);
        dispatchPassthroughSummary();
        binding = bindEngine(wb);
        namedRangeDialog?.bindWorkbook(wb);
        statusBar?.refresh();
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
        formatDialog?.detach();
        formatPainter?.detach();
        hover?.detach();
        conditionalDialog?.detach();
        goToDialog?.detach();
        iterativeDialog.detach();
        externalLinksDialog.detach();
        cellStylesGallery.detach();
        fxDialog?.detach();
        namedRangeDialog?.detach();
        pageSetupDialog?.detach();
        hyperlinkDialog?.detach();
        statusBar?.detach();
        watchPanel?.detach();
        unsubWatchRecalc();
        unsubWatchWb();
        slicer?.detach();
        unsubSlicerRecalc();
        unsubSlicerWb();
        errorMenu?.detach();
        if (errorMenu) canvas.removeEventListener('click', onCanvasClick);
        host.removeEventListener('keydown', onHostKey);
        detachWheel();
        fxAutocomplete.detach();
        fxArgHelper?.detach();
        fxInput.removeEventListener('focus', onFxFocus);
        fxInput.removeEventListener('input', onFxInput);
        fxInput.removeEventListener('keyup', onFxKeyUp);
        fxInput.removeEventListener('keydown', onFxKey);
        fxInput.removeEventListener('blur', onFxBlur);
        tag.removeEventListener('focus', onTagFocus);
        tag.removeEventListener('keydown', onTagKey);
        tag.removeEventListener('blur', onTagBlur);
        unsub();
        unsubI18n();
        i18n.dispose();
        renderer.dispose();
        if (ownsWb) wb.dispose();
        // Only touch the host if a later mount hasn't claimed it. See
        // `instanceId` stamp above for the StrictMode race this guards.
        if (host.dataset.fcInstId === instanceId) {
          host.replaceChildren();
          host.classList.remove('fc-host');
          delete host.dataset.fcInstId;
        }
      },
    };
  },
};

/** Case-insensitive defined-name lookup. Returns the formula text stripped
 *  of any leading `=`, sheet qualifier, and `$` anchors so it can be parsed
 *  by `parseRangeRef` / `parseCellRef`. */
function lookupDefinedName(wb: WorkbookHandle, query: string): string | null {
  if (!query) return null;
  const q = query.toLowerCase();
  for (const dn of wb.definedNames()) {
    if (dn.name.toLowerCase() !== q) continue;
    const eq = dn.formula.replace(/^=/, '');
    const bang = eq.lastIndexOf('!');
    return (bang >= 0 ? eq.slice(bang + 1) : eq).replace(/\$/g, '');
  }
  return null;
}

function parseCellRef(raw: string): { row: number; col: number } | null {
  const trimmed = raw.trim().toUpperCase();
  // R1C1 form: e.g. "R5C2"
  const r1c1 = trimmed.match(/^R([1-9][0-9]*)C([1-9][0-9]*)$/);
  if (r1c1) {
    const row = Number.parseInt(r1c1[1] ?? '', 10) - 1;
    const col = Number.parseInt(r1c1[2] ?? '', 10) - 1;
    if (row < 0 || col < 0) return null;
    if (col > 16383 || row > 1048575) return null;
    return { row, col };
  }
  const m = trimmed.match(/^\$?([A-Z]+)\$?([1-9][0-9]*)$/);
  if (!m) return null;
  const letters = m[1] ?? '';
  const rowStr = m[2] ?? '';
  let col = 0;
  for (let i = 0; i < letters.length; i += 1) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  col -= 1;
  const row = Number.parseInt(rowStr, 10) - 1;
  if (col < 0 || row < 0) return null;
  if (col > 16383 || row > 1048575) return null;
  return { row, col };
}

/** Parse A1:B5 style range. Returns null when the input doesn't match. */
function parseRangeRef(raw: string): { r0: number; c0: number; r1: number; c1: number } | null {
  const parts = raw.trim().toUpperCase().split(':');
  if (parts.length !== 2) return null;
  const a = parseCellRef(parts[0] ?? '');
  const b = parseCellRef(parts[1] ?? '');
  if (!a || !b) return null;
  return {
    r0: Math.min(a.row, b.row),
    c0: Math.min(a.col, b.col),
    r1: Math.max(a.row, b.row),
    c1: Math.max(a.col, b.col),
  };
}
