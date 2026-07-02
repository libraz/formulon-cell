import { coerceInput, writeCoerced, writeInputValidated } from '../commands/coerce-input.js';
import { replaceFormulaSelectionWithF9Preview } from '../commands/f9-preview.js';
import { stepWithMerge } from '../commands/merge.js';
import { dblClickRange, extractRefs, rotateRefAt, shiftFormulaRefs } from '../commands/refs.js';
import { addrKey } from '../engine/address.js';
import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { cellRect } from '../render/geometry.js';
import { formatWithPending, sameAddr } from '../store/pending-format.js';
import { type CellFormat, mutators, type SpreadsheetStore } from '../store/store.js';
import { type ArgHelperHandle, type ArgHelperLabels, attachArgHelper } from './arg-helper.js';
import {
  type AutocompleteHandle,
  type AutocompleteLabels,
  attachAutocomplete,
} from './autocomplete.js';

const MAX_ROW = 1_048_575;
const MAX_COL = 16_383;

const syncEditorRefs = (store: SpreadsheetStore, text: string): void => {
  const refs = extractRefs(text).map((r) => ({
    r0: r.r0,
    c0: r.c0,
    r1: r.r1,
    c1: r.c1,
    colorIndex: r.colorIndex,
  }));
  mutators.setEditorRefs(store, refs);
};

export interface EditorDeps {
  host: HTMLElement;
  grid: HTMLElement;
  store: SpreadsheetStore;
  wb: WorkbookHandle;
  /** Called whenever the engine state changed and the surrounding cell
   *  cache needs to be refreshed. */
  onAfterCommit: () => void;
  /** Optional callback fired when validation rejects (severity `stop`) or
   *  warns (`warning` / `information`). The host wires this to a status-bar
   *  toast. When omitted the editor logs to the console. */
  onValidation?: (outcome: {
    severity: 'stop' | 'warning' | 'information';
    title?: string;
    message: string;
  }) => void;
  getLabels?: () => {
    autocomplete?: Partial<AutocompleteLabels>;
    argHelper?: Partial<ArgHelperLabels>;
  };
  getCustomFunctions?: () => readonly string[];
}

/**
 * Inline cell editor — a single-line `<input>` floated over the active
 * cell. Begins on Enter / F2 / printable key. Commits on Enter or Tab,
 * cancels on Escape. Click-outside also commits.
 */
export class InlineEditor {
  private readonly deps: EditorDeps;

  private input: HTMLTextAreaElement | null = null;

  private editingAddr: Addr | null = null;

  private autocomplete: AutocompleteHandle | null = null;

  private argHelper: ArgHelperHandle | null = null;

  constructor(deps: EditorDeps) {
    this.deps = deps;
  }

  /** True when the active editor is sitting on a formula edit (`=`-prefixed)
   *  and is therefore willing to accept range-insert clicks. */
  isFormulaEdit(): boolean {
    return this.input?.value.startsWith('=') ?? false;
  }

  /** Insert `ref` at the current caret, replacing any selection. Used by the
   *  pointer layer to inject a clicked cell/range reference into a live
   *  formula edit. */
  insertRefAtCaret(ref: string): void {
    if (!this.input) return;
    const el = this.input;
    const start = el.selectionStart ?? el.value.length;
    const end = el.selectionEnd ?? start;
    const before = el.value.slice(0, start);
    const after = el.value.slice(end);
    // Replace any trailing partial ref token. Cases handled:
    //   "=A"           → drop "A" (partial ref)
    //   "=A1"          → drop "A1" (full ref)
    //   "=A1:B"        → drop "A1:B" (partial range)
    //   "=SUM(A1:B5"   → drop "A1:B5" (full range, keep prefix)
    const stripped = before.replace(
      /(?:\$?[A-Za-z]+\$?\d+:\$?[A-Za-z]+\$?\d*|\$?[A-Za-z]+\$?\d*|:)$/,
      '',
    );
    el.value = stripped + ref + after;
    const caret = stripped.length + ref.length;
    el.setSelectionRange(caret, caret);
    el.focus();
    this.refreshSize();
    syncEditorRefs(this.deps.store, el.value);
    this.argHelper?.refresh();
  }

  begin(seed: string): void {
    const s = this.deps.store.getState();
    const a = s.selection.active;
    this.editingAddr = a;
    mutators.setEditor(this.deps.store, { kind: 'enter', raw: seed });

    const input = document.createElement('textarea');
    input.className = 'fc-host__editor';
    input.spellcheck = false;
    input.autocapitalize = 'off';
    input.autocomplete = 'off';
    input.rows = 1;
    input.wrap = 'soft';
    input.value = seed;
    this.input = input;
    this.applyTextAlignment(seed);
    this.position(a);
    this.deps.grid.appendChild(input);
    this.refreshSize();

    // Focus synchronously so the *next* keystroke (post-seed) lands on the
    // editor input, not on the host. Deferring this via requestAnimationFrame
    // creates a race: rapid typing (Playwright, real-world fast typists) sends
    // subsequent keystrokes before raf fires; the host's keydown handler then
    // sees `editor.kind !== 'idle'` and silently drops them.
    input.focus();
    input.setSelectionRange(seed.length, seed.length);

    input.addEventListener('keydown', this.onKey);
    input.addEventListener('keyup', this.onKeyUp);
    input.addEventListener('input', this.onInput);
    input.addEventListener('blur', this.onBlur);
    input.addEventListener('dblclick', this.onDblClick);
    input.addEventListener('click', this.onClick);
    this.autocomplete = attachAutocomplete({
      input,
      onAfterInsert: () => syncEditorRefs(this.deps.store, input.value),
      getTables: () => this.deps.wb.getTables(),
      editingAddr: a,
      getColumnValues: (sheet, col, beforeRow) => this.collectColumnHistory(sheet, col, beforeRow),
      getCustomFunctions: () => this.deps.getCustomFunctions?.() ?? [],
      getFunctionNames: () => this.deps.wb.functionNames(),
      labels: this.deps.getLabels?.().autocomplete,
    });
    this.argHelper = attachArgHelper({ input, labels: this.deps.getLabels?.().argHelper });
    this.argHelper.refresh();
    syncEditorRefs(this.deps.store, seed);
  }

  cancel(): void {
    if (!this.input) return;
    this.autocomplete?.detach();
    this.autocomplete = null;
    this.argHelper?.detach();
    this.argHelper = null;
    this.input.removeEventListener('keydown', this.onKey);
    this.input.removeEventListener('keyup', this.onKeyUp);
    this.input.removeEventListener('input', this.onInput);
    this.input.removeEventListener('blur', this.onBlur);
    this.input.removeEventListener('dblclick', this.onDblClick);
    this.input.removeEventListener('click', this.onClick);
    this.input.remove();
    this.input = null;
    this.editingAddr = null;
    mutators.setEditor(this.deps.store, { kind: 'idle' });
    mutators.setPendingFormat(this.deps.store, null);
    mutators.setEditorRefs(this.deps.store, []);
    // Removing the focused input drops focus to <body>; without this, the
    // host's keydown listener stops receiving navigation keys until the
    // user clicks back in.
    this.deps.host.focus({ preventScroll: true });
  }

  commit(advance: 'down' | 'right' | 'none' = 'down'): void {
    if (!this.input || !this.editingAddr) return;
    const raw = this.input.value;
    const a = this.editingAddr;
    const fmt = formatWithPending(this.deps.store.getState(), a);
    let rejected = false;
    let rejectedOutcome: { severity: 'stop'; title?: string; message: string } | null = null;
    try {
      const outcome = writeInputValidated(this.deps.wb, a, raw, fmt?.validation, this.deps.store);
      if (!outcome.ok) {
        rejected = outcome.severity === 'stop';
        if (rejected) {
          rejectedOutcome = {
            severity: 'stop',
            title: fmt?.validation?.errorTitle,
            message: outcome.message,
          };
        } else if (this.deps.onValidation) {
          this.deps.onValidation({
            severity: outcome.severity,
            title: fmt?.validation?.errorTitle,
            message: outcome.message,
          });
        } else {
          console.warn(`formulon-cell: validation ${outcome.severity}: ${outcome.message}`);
        }
      }
    } catch (err) {
      console.warn('formulon-cell: writeInput failed', err);
    }
    if (rejected) {
      // Keep the editor open with the offending value so the user can correct.
      this.input.focus();
      this.input.select();
      if (rejectedOutcome) this.deps.onValidation?.(rejectedOutcome);
      return;
    }
    const pending = this.deps.store.getState().ui.pendingFormat;
    if (pending && sameAddr(pending.addr, a)) {
      mutators.setCellFormat(this.deps.store, a, pending.format);
      mutators.setPendingFormat(this.deps.store, null);
    }
    this.deps.onAfterCommit();
    this.cancel();
    const s = this.deps.store.getState();
    if (advance === 'down') {
      mutators.setActive(
        this.deps.store,
        stepWithMerge(s, s.selection.active, 1, 0, MAX_ROW, MAX_COL),
      );
    } else if (advance === 'right') {
      mutators.setActive(
        this.deps.store,
        stepWithMerge(s, s.selection.active, 0, 1, MAX_ROW, MAX_COL),
      );
    }
  }

  /** the spreadsheet's Ctrl+Enter behavior: write the current editor content to every
   *  cell in `selection.range` (and `extraRanges`), shifting relative refs in
   *  formulas as if filled. The active cell is the anchor — the source for
   *  relative-ref deltas. After committing, the active cell stays put. */
  commitMulti(): void {
    if (!this.input || !this.editingAddr) return;
    const input = this.input;
    const raw = input.value;
    const anchor = this.editingAddr;
    const s = this.deps.store.getState();
    const ranges = [s.selection.range, ...(s.selection.extraRanges ?? [])];
    const sheet = s.data.sheetIndex;
    const isFormula = raw.startsWith('=');
    // Validated write that mirrors the anchor's stop-rejection handling. Returns
    // true when a `stop` rule blocked the entry (the whole fill aborts) so DV
    // bites on every filled cell, not just the anchor.
    const writeValidatedOrAbort = (
      target: Addr,
      text: string,
      fmt: CellFormat | undefined,
    ): boolean => {
      const outcome = writeInputValidated(
        this.deps.wb,
        target,
        text,
        fmt?.validation,
        this.deps.store,
      );
      if (!outcome.ok && outcome.severity === 'stop') {
        this.deps.onValidation?.({
          severity: outcome.severity,
          title: fmt?.validation?.errorTitle,
          message: outcome.message,
        });
        input.focus();
        input.select();
        return true;
      }
      return false;
    };
    for (const r of ranges) {
      for (let row = r.r0; row <= r.r1; row += 1) {
        for (let col = r.c0; col <= r.c1; col += 1) {
          const target = { sheet, row, col };
          const fmt =
            target.sheet === anchor.sheet && target.row === anchor.row && target.col === anchor.col
              ? formatWithPending(this.deps.store.getState(), target)
              : s.format.formats.get(addrKey(target));
          const forceText = fmt?.numFmt?.kind === 'text';
          if (isFormula && !forceText) {
            // Formula fill: relative refs shift by the paste offset. Formula
            // results aren't run through DV here (Excel validates the typed
            // entry, not the recomputed result).
            const shifted = shiftFormulaRefs(raw, row - anchor.row, col - anchor.col);
            try {
              writeCoerced(this.deps.wb, target, { kind: 'formula', text: shifted });
            } catch (err) {
              console.warn('formulon-cell: writeCoerced failed', err);
            }
          } else {
            // Value fill (including text-formatted cells): every target validates
            // against its own rule via the store-aware coercion path.
            if (writeValidatedOrAbort(target, raw, fmt)) return;
          }
        }
      }
    }
    const pending = this.deps.store.getState().ui.pendingFormat;
    if (pending && sameAddr(pending.addr, anchor)) {
      mutators.setCellFormat(this.deps.store, anchor, pending.format);
      mutators.setPendingFormat(this.deps.store, null);
    }
    this.deps.onAfterCommit();
    this.cancel();
  }

  isActive(): boolean {
    return this.input != null;
  }

  private readonly onKey = (e: KeyboardEvent): void => {
    // When the autocomplete is open, intercept arrow/enter/tab/escape so they
    //  drive the popover instead of the surrounding editor.
    if (this.autocomplete?.isOpen()) {
      if (e.key === 'ArrowDown') {
        e.preventDefault();
        this.autocomplete.move(1);
        return;
      }
      if (e.key === 'ArrowUp') {
        e.preventDefault();
        this.autocomplete.move(-1);
        return;
      }
      if (e.key === 'Enter' || e.key === 'Tab') {
        if (this.autocomplete.acceptHighlighted()) {
          e.preventDefault();
          return;
        }
      }
      if (e.key === 'Escape') {
        e.preventDefault();
        this.autocomplete.close();
        return;
      }
    }
    if (e.key === 'Enter') {
      // Ctrl+Enter writes the same value/formula to every cell in the active
      //  selection (spreadsheet parity). On Mac spreadsheets use Control too, not Cmd, so
      //  metaKey keeps the legacy "newline" behavior to avoid surprising Mac
      //  users typing ⌘⏎.
      if (e.ctrlKey && !e.altKey && !e.shiftKey && !e.metaKey) {
        e.preventDefault();
        // stopPropagation: once commitMulti() flips editor.kind back to idle,
        // the same Enter would bubble to the host and start a new edit there.
        e.stopPropagation();
        this.commitMulti();
        return;
      }
      // Alt+Enter / Shift+Enter / Cmd+Enter inserts a literal newline (desktop spreadsheets
      //  Alt+Enter behavior). Plain Enter commits and advances down.
      if (e.altKey || e.shiftKey || e.metaKey) {
        e.preventDefault();
        this.insertNewline();
        return;
      }
      e.preventDefault();
      // Same reason as commitMulti above — commit() returns the editor to
      // idle, the host's keydown listener would then re-process Enter and
      // double-step the cursor.
      e.stopPropagation();
      this.commit('down');
    } else if (e.key === 'Escape') {
      e.preventDefault();
      e.stopPropagation();
      this.cancel();
    } else if (e.key === 'Tab') {
      e.preventDefault();
      e.stopPropagation();
      this.commit(e.shiftKey ? 'none' : 'right');
    } else if (e.key === 'F4' && this.input) {
      // Rotate the cell ref under the cursor: A1 → $A$1 → A$1 → $A1 → A1
      e.preventDefault();
      const caret = this.input.selectionStart ?? this.input.value.length;
      const r = rotateRefAt(this.input.value, caret);
      if (r.text !== this.input.value) {
        this.input.value = r.text;
        this.input.setSelectionRange(r.caret, r.caret);
        syncEditorRefs(this.deps.store, this.input.value);
      }
    } else if (e.key === 'F9' && this.input) {
      const start = this.input.selectionStart ?? 0;
      const end = this.input.selectionEnd ?? start;
      const sheetByName = (name: string): number => {
        for (let i = 0; i < this.deps.wb.sheetCount; i += 1) {
          if (this.deps.wb.sheetName(i) === name) return i;
        }
        return -1;
      };
      const result = replaceFormulaSelectionWithF9Preview(
        this.input.value,
        start,
        end,
        this.deps.store.getState().data.sheetIndex,
        this.deps.store.getState().data.cells,
        sheetByName,
      );
      if (result) {
        e.preventDefault();
        e.stopPropagation();
        this.input.value = result.text;
        this.input.setSelectionRange(result.start, result.end);
        syncEditorRefs(this.deps.store, this.input.value);
        this.autocomplete?.refresh();
        this.argHelper?.refresh();
      }
    }
  };

  private readonly onKeyUp = (): void => {
    // Caret moves on arrow / Home / End / click — those don't fire `input`,
    //  but the active argument can change. Refresh the tooltip alone.
    this.argHelper?.refresh();
  };

  private readonly onBlur = (): void => {
    // Blur commits unless we're already torn down.
    if (this.input) this.commit('none');
  };

  private readonly onInput = (): void => {
    this.refreshSize();
    if (this.input) this.applyTextAlignment(this.input.value);
    if (this.input) {
      mutators.setEditor(this.deps.store, {
        kind: 'edit',
        raw: this.input.value,
        caret: this.input.selectionStart ?? this.input.value.length,
      });
      syncEditorRefs(this.deps.store, this.input.value);
    }
    this.autocomplete?.refresh();
    this.argHelper?.refresh();
  };

  /** Double-click inside a formula edit selects a semantic token rather than
   *  a bare word — the function name, or a whole call argument (so a range
   *  like `F4:F8` selects as a unit). Non-formula edits keep the browser's
   *  native word selection. */
  private readonly onDblClick = (e: MouseEvent): void => {
    const el = this.input;
    if (!el?.value.startsWith('=')) return;
    // The browser has already selected a word; probe its midpoint so the hit
    //  point lands inside the intended token.
    const a = el.selectionStart ?? 0;
    const b = el.selectionEnd ?? a;
    const probe = a === b ? a : Math.floor((a + b) / 2);
    const range = dblClickRange(el.value, probe);
    if (range) {
      e.preventDefault();
      el.setSelectionRange(range.start, range.end);
    }
    this.argHelper?.refresh();
  };

  /** Triple-click inside a formula edit selects the whole formula. */
  private readonly onClick = (e: MouseEvent): void => {
    const el = this.input;
    if (!el || e.detail < 3 || !el.value.startsWith('=')) return;
    e.preventDefault();
    el.setSelectionRange(0, el.value.length);
    this.argHelper?.refresh();
  };

  private insertNewline(): void {
    const el = this.input;
    if (!el) return;
    const start = el.selectionStart ?? el.value.length;
    const end = el.selectionEnd ?? start;
    el.value = `${el.value.slice(0, start)}\n${el.value.slice(end)}`;
    const caret = start + 1;
    el.setSelectionRange(caret, caret);
    this.refreshSize();
  }

  private refreshSize(): void {
    if (!this.input) return;
    const lines = Math.max(1, (this.input.value.match(/\n/g)?.length ?? 0) + 1);
    if (lines === 1) {
      // Hide the per-line growth on a fresh single-line edit so the editor
      //  visually matches the cell rect exactly.
      this.input.style.minHeight = '';
    } else {
      // Spreadsheets grow the editor downward; mirror that with a min-height bump.
      const baseRow = this.deps.store.getState().layout.defaultRowHeight;
      this.input.style.minHeight = `${baseRow * lines}px`;
    }
    this.refreshWidth();
  }

  /** Grow the editor rightward to fit content wider than the cell — desktop
   *  spreadsheet behavior, so a long formula shows from its leading `=`
   *  instead of scrolling to its tail. Capped at the grid's right edge so the
   *  editor never escapes the viewport. */
  private refreshWidth(): void {
    if (!this.input || !this.editingAddr) return;
    const s = this.deps.store.getState();
    const r = cellRect(s.layout, s.viewport, this.editingAddr.row, this.editingAddr.col);
    // Reset to the cell width first so the editor can shrink back as content
    //  is deleted, then measure the natural content width.
    this.input.style.width = `${r.w}px`;
    const content = this.input.scrollWidth;
    if (content <= r.w) {
      this.input.classList.remove('fc-host__editor--overflow');
      return;
    }
    const maxWidth = Math.max(r.w, this.deps.grid.clientWidth - r.x - 2);
    const want = content + 2;
    this.input.style.width = `${Math.min(want, maxWidth)}px`;
    // When the content is still wider than the editor can grow, the text
    //  scrolls; flag the overflow so the editor drops its right border.
    this.input.classList.toggle('fc-host__editor--overflow', want > maxWidth);
  }

  private applyTextAlignment(raw: string): void {
    if (!this.input || !this.editingAddr) return;
    // Explicit cell alignment wins over the type-based defaults below — once
    // a user sets center/right/justify on a cell, the editor honors it even
    // for formula edits so the visual position stays anchored.
    const fmt = this.deps.store.getState().format.formats.get(addrKey(this.editingAddr));
    if (fmt?.align) {
      this.input.style.textAlign = fmt.align;
      return;
    }
    // Formula edits default to left so the leading `=` stays in view rather
    // than scrolling off to the right under content overflow.
    if (raw.startsWith('=')) {
      this.input.style.textAlign = 'left';
      return;
    }
    const coerced = coerceInput(raw);
    if (coerced.kind === 'number') this.input.style.textAlign = 'right';
    else if (coerced.kind === 'bool') this.input.style.textAlign = 'center';
    else this.input.style.textAlign = 'left';
  }

  /** Walk the column upward from `beforeRow - 1` collecting plain-text values
   *  for the autocomplete popover. Mirrors the "pick from list" rules:
   *  text-only (formulas, numbers, blanks all skip), deduped, nearest-first.
   *  Iterates the engine's populated-cells list once rather than probing each
   *  row — we'd otherwise call `cellFormula` (O(n)) per row, blowing up at
   *  every keystroke. */
  private collectColumnHistory(sheet: number, col: number, beforeRow: number): string[] {
    const hits: { row: number; text: string }[] = [];
    for (const e of this.deps.wb.cells(sheet)) {
      if (e.addr.col !== col) continue;
      if (e.addr.row >= beforeRow) continue;
      // Formulas don't contribute — the pick-list is verbatim text only.
      if (e.formula !== null) continue;
      if (e.value.kind !== 'text') continue;
      const text = e.value.value;
      if (text.length === 0) continue;
      hits.push({ row: e.addr.row, text });
    }
    // Nearest-first: highest row index wins.
    hits.sort((a, b) => b.row - a.row);
    const out: string[] = [];
    const seen = new Set<string>();
    for (const h of hits) {
      if (seen.has(h.text)) continue;
      seen.add(h.text);
      out.push(h.text);
      if (out.length >= 10) break;
    }
    return out;
  }

  private position(a: Addr): void {
    const s = this.deps.store.getState();
    const r = cellRect(s.layout, s.viewport, a.row, a.col);
    if (!this.input) return;
    this.input.style.left = `${r.x}px`;
    this.input.style.top = `${r.y}px`;
    this.input.style.width = `${r.w}px`;
    this.input.style.height = `${r.h}px`;
  }
}
