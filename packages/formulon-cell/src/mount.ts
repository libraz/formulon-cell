import { CellRegistry } from './cells.js';
import { writeInputValidated } from './commands/coerce-input.js';
import { fillRange } from './commands/fill.js';
import { toggleBold, toggleItalic, toggleStrike, toggleUnderline } from './commands/format.js';
import { History, recordFormatChange } from './commands/history.js';
import { extractRefs, rotateRefAt } from './commands/refs.js';
import { flushFormatToEngine, hydrateCellFormatsFromEngine } from './engine/cell-format-sync.js';
import { hydrateCommentsAndHyperlinksFromEngine } from './engine/format-sync.js';
import { hydrateLayoutFromEngine } from './engine/layout-sync.js';
import { hydrateMergesFromEngine } from './engine/merges-sync.js';
import { summarizePassthroughs, summarizeTables } from './engine/passthrough-sync.js';
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
  type Extension,
  type ExtensionContext,
  type ExtensionHandle,
  type ExtensionInput,
  type FeatureFlags,
  type ThemeName,
  dedupeById,
  flattenExtensions,
  resolveFlags,
  sortByPriority,
} from './extensions/index.js';
import type { CustomFunction, CustomFunctionMeta } from './formula.js';
import { FormulaRegistry } from './formula.js';
import { type I18nController, createI18nController } from './i18n/controller.js';
import type { DeepPartial, Locale, Strings } from './i18n/strings.js';
import { attachArgHelper } from './interact/arg-helper.js';
import { attachAutocomplete } from './interact/autocomplete.js';
import { attachCellStylesGallery } from './interact/cell-styles-gallery.js';
import { attachClipboard } from './interact/clipboard.js';
import { attachConditionalDialog } from './interact/conditional-dialog.js';
import { attachContextMenu } from './interact/context-menu.js';
import { InlineEditor } from './interact/editor.js';
import { attachFindReplace } from './interact/find-replace.js';
import { attachFormatDialog } from './interact/format-dialog.js';
import { type FormatPainterHandle, attachFormatPainter } from './interact/format-painter.js';
import { attachHover } from './interact/hover.js';
import { attachHyperlinkDialog } from './interact/hyperlink-dialog.js';
import { attachIterativeDialog } from './interact/iterative-dialog.js';
import { attachKeyboard } from './interact/keyboard.js';
import { attachNamedRangeDialog } from './interact/named-range-dialog.js';
import { attachPasteSpecial } from './interact/paste-special.js';
import { attachPointer } from './interact/pointer.js';
import { attachStatusBar } from './interact/status-bar.js';
import { attachValidationList } from './interact/validation.js';
import { attachWheel } from './interact/wheel.js';
import { GridRenderer } from './render/grid.js';
import { type SpreadsheetStore, createSpreadsheetStore, mutators } from './store/store.js';
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
  /** Open the conditional-formatting rule manager dialog. No-op when the
   *  feature is disabled. */
  openConditionalDialog(): void;
  /** Open the read-only named-range listing dialog. */
  openNamedRangeDialog(): void;
  /** Open the cell format dialog (Excel ⌘1). */
  openFormatDialog(): void;
  /** Open the iterative-calculation settings dialog (Excel File → Options
   *  → Formulas). */
  openIterativeDialog(): void;
  /** Open the named cell-styles gallery (Excel Home → Cell Styles). */
  openCellStylesGallery(): void;
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
    const flags = resolveFlags(opts.features);
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
    const fx = document.createElement('span');
    fx.className = 'fc-host__formulabar-fx';
    fx.textContent = 'ƒx';
    fx.setAttribute('aria-hidden', 'true');
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
    host.append(formulabar, grid, statusbar);

    let wb: WorkbookHandle = opts.workbook ?? (await WorkbookHandle.createDefault());
    if (opts.seed) opts.seed(wb);
    let ownsWb = !opts.workbook;

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

    // Built-in feature instances. Each one is gated by `flags.<id>` — when
    // disabled the binding is `null` and consumers / cross-feature
    // references must use `?.` to no-op gracefully.
    const formatDialog = flags.formatDialog
      ? attachFormatDialog({ host, store, strings, history, getWb: () => wb })
      : null;
    const formatPainter = flags.formatPainter
      ? attachFormatPainter({ host, store, history })
      : null;
    const hover = flags.hoverComment ? attachHover({ grid, store }) : null;
    const conditionalDialog = flags.conditional
      ? attachConditionalDialog({ host, store, strings })
      : null;
    // Iterative-calc settings dialog isn't on the feature menu yet — keep
    // it always-on so the public openIterativeDialog() never returns silent.
    const iterativeDialog = attachIterativeDialog({ host, getWb: () => wb, strings });
    const cellStylesGallery = attachCellStylesGallery({
      host,
      store,
      history,
      getWb: () => wb,
    });
    const namedRangeDialog = flags.namedRanges
      ? attachNamedRangeDialog({ host, wb, strings })
      : null;
    const hyperlinkDialog = flags.hyperlink
      ? attachHyperlinkDialog({ host, store, strings, history, getWb: () => wb })
      : null;
    const statusBar = flags.statusBar
      ? attachStatusBar({
          statusbar,
          store,
          strings,
          getEngineLabel: () => (wb.isStub ? 'stub' : `formulon ${wb.version}`),
        })
      : null;

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
          formatCellForEdit(s.data.cells.get(`${a.sheet}:${a.row}:${a.col}`));
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
    if (flags.shortcuts) host.addEventListener('keydown', onHostKey);

    const detachWheel = flags.wheel ? attachWheel({ grid, store, wb }) : (): void => {};

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
    // Function autocomplete inside the formula bar. When disabled, hand back
    // a stub so the rest of the fxKey handling can call its methods safely.
    const fxAutocomplete = flags.autocomplete
      ? attachAutocomplete({
          input: fxInput,
          onAfterInsert: () => syncFxRefs(),
          getTables: () => wb.getTables(),
          getCustomFunctions: () => formulaRegistry.list(),
        })
      : {
          isOpen: () => false,
          move: (_n: number) => {},
          acceptHighlighted: () => false,
          close: () => {},
          refresh: () => {},
          detach: () => {},
        };
    const fxArgHelper = flags.autocomplete ? attachArgHelper({ input: fxInput }) : null;
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
          default:
            display = '';
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

    // Built-in feature registry. Each entry exposes the `dispose` hook so
    // `instance.remove(id)` can tear it down later. Order doesn't matter
    // here — we use it for lookup, not iteration.
    const featureRegistry = new Map<string, ExtensionHandle>();
    const registerBuiltIn = (id: string, raw: unknown, detach: () => void): void => {
      const handle = (
        raw && typeof raw === 'object' ? (raw as Record<string, unknown>) : {}
      ) as ExtensionHandle;
      handle.dispose = detach;
      featureRegistry.set(id, handle);
    };
    if (formatDialog) registerBuiltIn('formatDialog', formatDialog, formatDialog.detach);
    if (formatPainter) registerBuiltIn('formatPainter', formatPainter, formatPainter.detach);
    if (hover) registerBuiltIn('hoverComment', hover, hover.detach);
    if (conditionalDialog)
      registerBuiltIn('conditional', conditionalDialog, conditionalDialog.detach);
    if (namedRangeDialog) registerBuiltIn('namedRanges', namedRangeDialog, namedRangeDialog.detach);
    if (hyperlinkDialog) registerBuiltIn('hyperlink', hyperlinkDialog, hyperlinkDialog.detach);
    if (statusBar) registerBuiltIn('statusBar', statusBar, statusBar.detach);

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
      formatPainter: formatPainter ?? undefined,
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
      openConditionalDialog() {
        conditionalDialog?.open();
      },
      openIterativeDialog() {
        iterativeDialog.open();
      },
      openCellStylesGallery() {
        cellStylesGallery.open();
      },
      openNamedRangeDialog() {
        namedRangeDialog?.open();
      },
      openFormatDialog() {
        formatDialog?.open();
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
        iterativeDialog.detach();
        cellStylesGallery.detach();
        namedRangeDialog?.detach();
        hyperlinkDialog?.detach();
        statusBar?.detach();
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

function formatCellForEdit(cell: { value: CellValue; formula: string | null } | undefined): string {
  if (!cell) return '';
  if (cell.formula) return cell.formula;
  const v = cell.value;
  switch (v.kind) {
    case 'number':
      return String(v.value);
    case 'bool':
      return v.value ? 'TRUE' : 'FALSE';
    case 'text':
      return v.value;
    case 'error':
      return v.text;
    default:
      return '';
  }
}

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
