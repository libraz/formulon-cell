import { writeInput } from './commands/coerce-input.js';
import { fillRange } from './commands/fill.js';
import { toggleBold, toggleItalic, toggleStrike, toggleUnderline } from './commands/format.js';
import { History, recordFormatChange } from './commands/history.js';
import { extractRefs, rotateRefAt } from './commands/refs.js';
import { flushFormatToEngine, hydrateCellFormatsFromEngine } from './engine/cell-format-sync.js';
import { hydrateCommentsAndHyperlinksFromEngine } from './engine/format-sync.js';
import { hydrateLayoutFromEngine } from './engine/layout-sync.js';
import { hydrateMergesFromEngine } from './engine/merges-sync.js';
import type { CellValue } from './engine/types.js';
import { hydrateValidationsFromEngine } from './engine/validation-sync.js';
import { type ChangeEvent, WorkbookHandle } from './engine/workbook-handle.js';
import {
  type DeepPartial,
  type Locale,
  type Strings,
  defaultStrings,
  dictionaries,
  mergeStrings,
} from './i18n/strings.js';
import { attachAutocomplete } from './interact/autocomplete.js';
import { attachClipboard } from './interact/clipboard.js';
import { attachConditionalDialog } from './interact/conditional-dialog.js';
import { attachContextMenu } from './interact/context-menu.js';
import { InlineEditor } from './interact/editor.js';
import { attachFindReplace } from './interact/find-replace.js';
import { attachFormatDialog } from './interact/format-dialog.js';
import { type FormatPainterHandle, attachFormatPainter } from './interact/format-painter.js';
import { attachHover } from './interact/hover.js';
import { attachHyperlinkDialog } from './interact/hyperlink-dialog.js';
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
  theme?: 'paper' | 'ink';
  /** Optional initial-cell seeding. Useful for the playground & docs. */
  seed?: (wb: WorkbookHandle) => void;
  /** UI locale for built-in dialogs and menus. Defaults to 'ja'. */
  locale?: Locale;
  /** Per-string overrides applied on top of the chosen locale. Deep-merged. */
  strings?: DeepPartial<Strings>;
}

export interface SpreadsheetInstance {
  readonly host: HTMLElement;
  readonly workbook: WorkbookHandle;
  readonly store: SpreadsheetStore;
  /** Unified undo/redo for cell, format, and layout changes. Each user-level
   *  action pushes one entry; transactions (paste, fill drag) are batched. */
  readonly history: History;
  /** Format Painter controls — surfaced so chrome (toolbar buttons) can
   *  arm/disarm and reflect the active state. */
  readonly formatPainter: FormatPainterHandle;
  /** Open the conditional-formatting rule manager dialog. */
  openConditionalDialog(): void;
  /** Open the read-only named-range listing dialog. */
  openNamedRangeDialog(): void;
  /** Open the cell format dialog (Excel ⌘1). */
  openFormatDialog(): void;
  setTheme(t: 'paper' | 'ink'): void;
  /** Pop the most recent undoable action and revert it. Returns false when
   *  the stack is empty. */
  undo(): boolean;
  /** Re-apply the most recently undone action. Returns false when nothing
   *  to redo. */
  redo(): boolean;
  /** Apply a fresh workbook (e.g. after `loadBytes`). Disposes old one. */
  setWorkbook(next: WorkbookHandle): Promise<void>;
  dispose(): void;
}

/**
 * Mount a spreadsheet onto a DOM host. Returns an instance with imperative
 * controls. The host element is taken over — its existing children are
 * cleared. Idempotent dispose.
 */
export const Spreadsheet = {
  async mount(host: HTMLElement, opts: MountOptions = {}): Promise<SpreadsheetInstance> {
    if (!host) throw new Error('Spreadsheet.mount: host element required');

    const baseStrings = opts.locale ? dictionaries[opts.locale] : defaultStrings;
    const strings = mergeStrings(baseStrings, opts.strings);

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

    const renderer = new GridRenderer({
      host: grid,
      canvas,
      getState: () => store.getState(),
      getTheme: () => resolveTheme(host),
      getWb: () => wb,
    });
    renderer.resize();

    // The format dialog needs wb so it can flush data-validation entries to
    // the engine on OK; pass a getter so setWorkbook swaps land transparently.
    const formatDialog = attachFormatDialog({
      host,
      store,
      strings,
      history,
      getWb: () => wb,
    });
    const formatPainter = attachFormatPainter({ host, store, history });
    const hover = attachHover({ grid, store });
    const conditionalDialog = attachConditionalDialog({ host, store, strings });
    const namedRangeDialog = attachNamedRangeDialog({ host, wb, strings });
    const hyperlinkDialog = attachHyperlinkDialog({
      host,
      store,
      strings,
      history,
      getWb: () => wb,
    });
    const statusBar = attachStatusBar({
      statusbar,
      store,
      strings,
      getEngineLabel: () => (wb.isStub ? 'stub' : `formulon ${wb.version}`),
    });

    // wb-dependent layer — re-built whenever setWorkbook swaps the engine.
    interface EngineBinding {
      editor: InlineEditor;
      pasteSpecialDialog: ReturnType<typeof attachPasteSpecial>;
      findReplace: ReturnType<typeof attachFindReplace>;
      validation: ReturnType<typeof attachValidationList>;
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
      const detachKey = attachKeyboard({
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
      });
      const clipboardH = attachClipboard({
        host,
        store,
        wb: currentWb,
        onAfterCommit: () =>
          mutators.replaceCells(store, currentWb.cells(store.getState().data.sheetIndex)),
      });
      const pasteSpecialDialog = attachPasteSpecial({
        host,
        store,
        wb: currentWb,
        strings,
        history,
        getSnapshot: () => clipboardH.getSnapshot(),
        onAfterCommit: () =>
          mutators.replaceCells(store, currentWb.cells(store.getState().data.sheetIndex)),
      });
      const detachContextMenu = attachContextMenu({
        host,
        store,
        wb: currentWb,
        strings,
        history,
        onAfterCommit: () =>
          mutators.replaceCells(store, currentWb.cells(store.getState().data.sheetIndex)),
        onFormatDialog: () => formatDialog.open(),
        onPasteSpecial: () => pasteSpecialDialog.open(),
        onInsertHyperlink: () => hyperlinkDialog.open(),
      });
      const findReplace = attachFindReplace({
        host,
        store,
        wb: currentWb,
        strings,
        onAfterCommit: () =>
          mutators.replaceCells(store, currentWb.cells(store.getState().data.sheetIndex)),
      });
      const validation = attachValidationList({
        grid,
        store,
        wb: currentWb,
        onAfterCommit: () =>
          mutators.replaceCells(store, currentWb.cells(store.getState().data.sheetIndex)),
      });

      const onDblClick = (e: MouseEvent): void => {
        if (e.button !== 0) return;
        if (editor.isActive()) return;
        if (formatPainter.isActive()) return;
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
          const cell = { value: e.next, formula: currentWb.cellFormula(e.addr) };
          store.setState((s) => {
            const cells = new Map(s.data.cells);
            cells.set(`${e.addr.sheet}:${e.addr.row}:${e.addr.col}`, cell);
            return { ...s, data: { ...s.data, cells } };
          });
        }
      });

      return {
        editor,
        pasteSpecialDialog,
        findReplace,
        validation,
        unbind: () => {
          detachPtr();
          detachKey();
          clipboardH.detach();
          detachContextMenu();
          findReplace.detach();
          pasteSpecialDialog.detach();
          validation.detach();
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
    const onHostKey = (e: KeyboardEvent): void => {
      const meta = e.ctrlKey || e.metaKey;
      if (!meta) return;
      const k = e.key.toLowerCase();
      if (e.shiftKey && k === 'c') {
        // Cmd/Ctrl+Shift+C — copy formatting (one-shot).
        e.preventDefault();
        formatPainter.activate(false);
        return;
      }
      if (e.shiftKey && k === 'v') {
        // Cmd/Ctrl+Shift+V — open Paste Special.
        e.preventDefault();
        binding.pasteSpecialDialog.open();
        return;
      }
      if (e.altKey && k === 'v') {
        // Excel alt-binding: Ctrl+Alt+V (Win) / Cmd+Option+V (Mac).
        e.preventDefault();
        binding.pasteSpecialDialog.open();
        return;
      }
      if (k === 'f') {
        e.preventDefault();
        binding.findReplace.open();
      } else if (k === 'k') {
        // Ctrl/Cmd+K — Insert Hyperlink dialog (Excel/Sheets parity).
        e.preventDefault();
        hyperlinkDialog.open();
      } else if (k === 'a') {
        e.preventDefault();
        mutators.selectAll(store);
      } else if (e.key === '1') {
        e.preventDefault();
        formatDialog.open();
      } else if (e.key === '`') {
        // Ctrl+` — toggle show-formulas mode.
        e.preventDefault();
        mutators.setShowFormulas(store, !store.getState().ui.showFormulas);
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
    host.addEventListener('keydown', onHostKey);

    const detachWheel = attachWheel({ grid, store, wb });

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
    const fxAutocomplete = attachAutocomplete({
      input: fxInput,
      onAfterInsert: () => syncFxRefs(),
    });
    const onFxFocus = (): void => {
      if (binding.editor.isActive()) binding.editor.cancel();
      fxEditing = true;
      fxBaseline = fxInput.value;
      syncFxRefs();
    };
    const onFxInput = (): void => {
      if (fxEditing) syncFxRefs();
      fxAutocomplete.refresh();
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
        writeInput(wb, a, fxInput.value);
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
    const unsub = store.subscribe(() => {
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

    return {
      host,
      get workbook() {
        return wb;
      },
      store,
      history,
      formatPainter,
      openConditionalDialog() {
        conditionalDialog.open();
      },
      openNamedRangeDialog() {
        namedRangeDialog.open();
      },
      openFormatDialog() {
        formatDialog.open();
      },
      setTheme(t) {
        host.dataset.fcTheme = t;
        mutators.setTheme(store, t);
        renderer.invalidate();
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
        history.clear();
        mutators.replaceCells(store, wb.cells(store.getState().data.sheetIndex));
        hydrateLayoutFromEngine(wb, store, store.getState().data.sheetIndex);
        hydrateCommentsAndHyperlinksFromEngine(wb, store, store.getState().data.sheetIndex);
        hydrateMergesFromEngine(wb, store, store.getState().data.sheetIndex);
        hydrateValidationsFromEngine(wb, store, store.getState().data.sheetIndex);
        hydrateCellFormatsFromEngine(wb, store, store.getState().data.sheetIndex);
        binding = bindEngine(wb);
        namedRangeDialog.bindWorkbook(wb);
        statusBar.refresh();
        updateChrome();
        renderer.invalidate();
      },
      dispose() {
        if (disposed) return;
        disposed = true;
        ro.disconnect();
        binding.unbind();
        formatDialog.detach();
        formatPainter.detach();
        hover.detach();
        conditionalDialog.detach();
        namedRangeDialog.detach();
        hyperlinkDialog.detach();
        statusBar.detach();
        host.removeEventListener('keydown', onHostKey);
        detachWheel();
        fxAutocomplete.detach();
        fxInput.removeEventListener('focus', onFxFocus);
        fxInput.removeEventListener('input', onFxInput);
        fxInput.removeEventListener('keydown', onFxKey);
        fxInput.removeEventListener('blur', onFxBlur);
        tag.removeEventListener('focus', onTagFocus);
        tag.removeEventListener('keydown', onTagKey);
        tag.removeEventListener('blur', onTagBlur);
        unsub();
        renderer.dispose();
        if (ownsWb) wb.dispose();
        host.replaceChildren();
        host.classList.remove('fc-host');
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
