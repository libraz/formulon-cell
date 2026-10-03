import { commitCellInput, inputOperation } from '../commands/cell-input-commit.js';
import { coerceInput } from '../commands/coerce-input.js';
import { replaceFormulaSelectionWithF9Preview } from '../commands/f9-preview.js';
import { interactionControllerFor } from '../commands/interaction-controller.js';
import { dblClickRange, extractRefs, rotateRefAt } from '../commands/refs.js';
import { addrKey } from '../engine/address.js';
import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { bodyBandOrigin, cellRectUnclamped, layoutForView } from '../render/geometry.js';
import { formatWithPending } from '../store/pending-format.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { type ArgHelperHandle, type ArgHelperLabels, attachArgHelper } from './arg-helper.js';
import {
  type AutocompleteHandle,
  type AutocompleteLabels,
  attachAutocomplete,
  pickListValues,
} from './autocomplete.js';
import {
  createFormulaEditLease,
  type FormulaEditLease,
  type FormulaEditLeaseContext,
  type FormulaEditLeaseSnapshot,
} from './formula-edit-lease.js';
import { navigationPolicyFor, nextTabStop } from './navigation-policy.js';
import {
  buildSelectionInputBatch,
  SELECTION_INPUT_LIMIT_MESSAGE,
  writeSelectionInput,
} from './selection-input.js';
import { advanceAfterCommit } from './selection-navigation.js';

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

const policyRejection = (operation: string): { severity: 'stop'; message: string } => ({
  severity: 'stop',
  message: `The ${operation} operation is not permitted for this cell.`,
});

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

  private composing = false;

  private editBaseline = '';

  private leaseGeneration = 0;

  private currentLease: FormulaEditLease | null = null;

  private unsubscribeStore: (() => void) | null = null;

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

  suspendForFormulaPalette(context: FormulaEditLeaseContext): FormulaEditLease | null {
    const input = this.input;
    const anchor = this.editingAddr;
    if (!input || !anchor || this.composing || this.currentLease) return null;
    const state = this.deps.store.getState();
    const generation = ++this.leaseGeneration;
    const sheetCount = this.deps.wb.sheetCount;
    const snapshot: FormulaEditLeaseSnapshot = {
      source: 'inline',
      workbook: this.deps.wb,
      anchor: { ...anchor },
      raw: input.value,
      baseline: this.editBaseline,
      caret: {
        start: input.selectionStart ?? input.value.length,
        end: input.selectionEnd ?? input.value.length,
        direction: input.selectionDirection,
      },
      selection: state.selection,
      editorMode: state.ui.editor,
      pendingFormat: state.ui.pendingFormat ?? null,
      editorRefs: state.ui.editorRefs,
      copy: {
        copyRange: state.ui.copyRange ?? null,
        copyRanges: state.ui.copyRanges ?? null,
        copyMode: state.ui.copyMode ?? null,
        copyRevision: state.ui.copyRevision ?? 0,
      },
      r1c1: state.ui.r1c1,
      locale: context.getLocale(),
    };

    this.detachInput(false);
    this.editingAddr = null;
    this.composing = false;
    mutators.setEditor(this.deps.store, { kind: 'idle' });
    mutators.setEditorRefs(this.deps.store, []);

    let lease: FormulaEditLease | null = null;
    lease = createFormulaEditLease(snapshot, {
      context,
      isOwnerCurrent: () => {
        if (this.currentLease !== lease || this.leaseGeneration !== generation) return false;
        const current = this.deps.store.getState();
        return (
          this.deps.wb === snapshot.workbook &&
          this.deps.wb.sheetCount === sheetCount &&
          current.data.sheetIndex === snapshot.anchor.sheet &&
          current.ui.r1c1 === snapshot.r1c1
        );
      },
      restore: (captured) => this.restoreFromLease(captured),
      release: () => {
        if (this.currentLease === lease) this.currentLease = null;
      },
    });
    this.currentLease = lease;
    return lease;
  }

  begin(seed: string): void {
    this.invalidateLease();
    const s = this.deps.store.getState();
    const a = s.selection.active;
    const controller = interactionControllerFor(this.deps.store);
    if (controller && controller.policy !== undefined) {
      const operation = inputOperation(seed);
      const permission = controller.canExecute({
        operation,
        origin: 'editor',
        effects: [{ kind: 'cells', cells: [a] }],
      });
      if (!permission.allowed) {
        this.deps.onValidation?.(policyRejection(operation));
        return;
      }
    }
    this.editingAddr = a;
    this.editBaseline = seed;
    // Putting a cell into edit mode cancels copy mode — the marquee only
    // survives navigation and paste-family commands.
    if (s.ui.copyRange || s.ui.copyRanges) {
      mutators.setCopyRange(this.deps.store, null);
      mutators.setCopyRanges(this.deps.store, null);
    }
    mutators.setEditor(this.deps.store, { kind: 'enter', raw: seed });

    this.mountInput(a, seed, (input) => {
      // Focus synchronously so the *next* keystroke (post-seed) lands on the
      // editor input, not on the host. Deferring this via requestAnimationFrame
      // creates a race: rapid typing (Playwright, real-world fast typists) sends
      // subsequent keystrokes before raf fires; the host's keydown handler then
      // sees `editor.kind !== 'idle'` and silently drops them.
      input.focus();
      input.setSelectionRange(seed.length, seed.length);
    });
    syncEditorRefs(this.deps.store, seed);
  }

  cancel(): void {
    this.invalidateLease();
    if (!this.input) {
      this.editingAddr = null;
      this.composing = false;
      mutators.setEditor(this.deps.store, { kind: 'idle' });
      mutators.setPendingFormat(this.deps.store, null);
      mutators.setEditorRefs(this.deps.store, []);
      return;
    }
    this.detachInput(true);
    this.editingAddr = null;
    this.composing = false;
    mutators.setEditor(this.deps.store, { kind: 'idle' });
    mutators.setPendingFormat(this.deps.store, null);
    mutators.setEditorRefs(this.deps.store, []);
    // Removing the focused input drops focus to <body>; without this, the
    // host's keydown listener stops receiving navigation keys until the
    // user clicks back in.
    this.deps.host.focus({ preventScroll: true });
  }

  commit(advance: 'down' | 'right' | 'up' | 'left' | 'none' = 'down'): void {
    if (!this.input || !this.editingAddr) return;
    const result = commitCellInput({
      store: this.deps.store,
      wb: this.deps.wb,
      addr: this.editingAddr,
      raw: this.input.value,
      origin: 'editor',
    });
    if (result.status === 'rejected') {
      // Keep the editor open with the offending value so the user can correct.
      this.input.focus();
      this.input.select();
      this.deps.onValidation?.(result.alert ?? policyRejection(result.operation));
      return;
    }
    if (result.status === 'failed') {
      console.warn('formulon-cell: editor write failed', result.error);
      this.input.focus();
      this.deps.onValidation?.({
        severity: 'stop',
        message: 'The cell value could not be written.',
      });
      return;
    } else if (result.notice) {
      if (this.deps.onValidation) this.deps.onValidation(result.notice);
      else {
        console.warn(
          `formulon-cell: validation ${result.notice.severity}: ${result.notice.message}`,
        );
      }
    }
    this.deps.onAfterCommit();
    this.cancel();
    this.advanceAfterCommit(advance);
  }

  private isMacPlatform(): boolean {
    return (
      (this.deps.host.closest<HTMLElement>('.fc-host') ?? this.deps.host).dataset.fcPlatform ===
      'mac'
    );
  }

  private advanceAfterCommit(advance: 'down' | 'right' | 'up' | 'left' | 'none'): void {
    if (advance !== 'none') advanceAfterCommit(this.deps.store, advance, this.isMacPlatform());
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
    const controller = interactionControllerFor(this.deps.store);
    if (controller) {
      const batch = buildSelectionInputBatch(s, raw, anchor);
      if (!batch) {
        this.deps.onValidation?.({ severity: 'stop', message: SELECTION_INPUT_LIMIT_MESSAGE });
        return;
      }
      const { changes, operation } = batch;
      let result: ReturnType<typeof controller.execute>;
      try {
        result = controller.execute({
          type: 'cellBatch',
          operation,
          origin: 'editor',
          changes,
          denied: 'reject',
        });
      } catch (err) {
        console.warn('formulon-cell: restricted editor fill failed', err);
        result = { status: 'rejected', applied: [], rejected: [], revision: 0 };
      }
      if (result.status === 'rejected') {
        input.focus();
        input.select();
        this.deps.onValidation?.({
          severity: 'stop',
          message: result.rejected[0]?.reason ?? policyRejection(operation).message,
        });
        return;
      }
      mutators.setPendingFormat(this.deps.store, null);
      this.deps.onAfterCommit();
      this.cancel();
      return;
    }
    const written = writeSelectionInput(this.deps.wb, this.deps.store, s, raw, anchor);
    if (written.status === 'limitExceeded') {
      this.deps.onValidation?.({ severity: 'stop', message: SELECTION_INPUT_LIMIT_MESSAGE });
      return;
    }
    if (written.status === 'rejected') {
      this.deps.onValidation?.(written.outcome);
      input.focus();
      input.select();
      return;
    }
    this.deps.onAfterCommit();
    this.cancel();
  }

  /** True while editing or while a formula palette holds the suspended edit;
   *  cancel() releases that lease, so teardown paths may call it either way. */
  isActive(): boolean {
    return this.input != null || this.currentLease != null;
  }

  private readonly onKey = (e: KeyboardEvent): void => {
    if (this.composing || e.isComposing || e.key === 'Process') {
      return;
    }
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
    if (
      e.metaKey &&
      !e.ctrlKey &&
      !e.altKey &&
      !e.shiftKey &&
      e.key.toLowerCase() === 't' &&
      this.isMacPlatform()
    ) {
      e.preventDefault();
      e.stopPropagation();
      const caret = this.input?.selectionStart ?? this.input?.value.length ?? 0;
      if (this.input) {
        const r = rotateRefAt(this.input.value, caret);
        if (r.text !== this.input.value) {
          this.input.value = r.text;
          this.input.setSelectionRange(r.caret, r.caret);
          syncEditorRefs(this.deps.store, this.input.value);
          this.autocomplete?.refresh();
          this.argHelper?.refresh();
        }
      }
      return;
    }
    if (e.key === 'Enter') {
      if (this.isMacPlatform() && e.metaKey && !e.ctrlKey && !e.altKey && !e.shiftKey) {
        e.preventDefault();
        e.stopPropagation();
        this.commitMulti();
        return;
      }
      // Ctrl+Enter writes the same value/formula to every cell in the active
      // selection on every platform. Mac Cmd+Enter is handled above.
      if (e.ctrlKey && !e.altKey && !e.shiftKey && !e.metaKey) {
        e.preventDefault();
        // stopPropagation: once commitMulti() flips editor.kind back to idle,
        // the same Enter would bubble to the host and start a new edit there.
        e.stopPropagation();
        this.commitMulti();
        return;
      }
      // Alt+Enter (and Meta+Enter off Mac) inserts a literal newline.
      if (e.altKey || e.metaKey) {
        e.preventDefault();
        this.insertNewline();
        return;
      }
      // Mac Shift+Enter commits upward; elsewhere it inserts a newline.
      if (e.shiftKey) {
        e.preventDefault();
        if (this.isMacPlatform()) {
          e.stopPropagation();
          this.commit('up');
        } else {
          this.insertNewline();
        }
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
      e.stopPropagation();
      const controller = interactionControllerFor(this.deps.store);
      const restricted = controller?.policy !== undefined;
      const navigation = navigationPolicyFor(this.deps.store);
      const hasNavigationPolicy = navigation?.options !== undefined;
      const tabRestricted = restricted || hasNavigationPolicy;
      const selectionDisabled = controller?.policy?.selection === false;
      if (selectionDisabled) {
        this.commit('none');
        return;
      }
      const next =
        tabRestricted && this.editingAddr
          ? nextTabStop(this.deps.store, this.editingAddr, e.shiftKey)
          : null;
      // At the configured boundary, commit the current cell and let the
      // browser continue focus traversal. A component must not trap Tab.
      if (tabRestricted && next === null) {
        this.commit('none');
        if (hasNavigationPolicy && navigation?.options?.tabBoundary !== 'leave') {
          e.preventDefault();
        }
        return;
      }
      e.preventDefault();
      this.commit(
        tabRestricted ? 'none' : e.shiftKey ? (this.isMacPlatform() ? 'left' : 'none') : 'right',
      );
      if (tabRestricted && next && !this.isActive()) mutators.setActive(this.deps.store, next);
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
      const active = this.deps.store.getState().selection.active;
      const result = replaceFormulaSelectionWithF9Preview(
        this.input.value,
        start,
        end,
        this.deps.store.getState().data.sheetIndex,
        this.deps.store.getState().data.cells,
        sheetByName,
        (formula) => this.deps.wb.evaluateFormulaText(active, formula),
        true,
        (formula) => this.deps.wb.evaluateFormulaArray(active, formula),
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

  private readonly onCompositionStart = (): void => {
    this.composing = true;
  };

  private readonly onCompositionEnd = (): void => {
    this.composing = false;
    this.onInput();
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

  /** Remove the transient editor DOM and its subscriptions without changing
   *  the editor slices. A palette lease uses this boundary before blur can
   *  turn a suspended edit into a workbook write; ordinary cancel asks for
   *  host focus after the same teardown. */
  private detachInput(focusHost: boolean): void {
    this.unsubscribeStore?.();
    this.unsubscribeStore = null;
    this.autocomplete?.detach();
    this.autocomplete = null;
    this.argHelper?.detach();
    this.argHelper = null;
    const input = this.input;
    if (!input) {
      if (focusHost) this.deps.host.focus({ preventScroll: true });
      return;
    }
    for (const [type, listener] of this.inputListeners()) {
      input.removeEventListener(type, listener as EventListener);
    }
    input.remove();
    this.input = null;
    if (focusHost) this.deps.host.focus({ preventScroll: true });
  }

  private invalidateLease(): void {
    this.leaseGeneration += 1;
    const lease = this.currentLease;
    this.currentLease = null;
    lease?.discard();
  }

  /** Rebuild an inline editor from a lease snapshot. This deliberately does
   * not call begin(): begin() clears clipboard state, runs policy checks, and
   * focuses a new edit. Palette Cancel owns focus and chooses when to focus
   * this returned element instead. */
  private restoreFromLease(snapshot: Readonly<FormulaEditLeaseSnapshot>): HTMLElement | null {
    if (this.input || this.deps.wb !== snapshot.workbook) return null;
    const state = this.deps.store.getState();
    if (state.data.sheetIndex !== snapshot.anchor.sheet || state.ui.r1c1 !== snapshot.r1c1) {
      return null;
    }
    this.editingAddr = { ...snapshot.anchor };
    this.editBaseline = snapshot.baseline;
    this.composing = false;
    this.deps.store.setState((current) => ({
      ...current,
      selection: snapshot.selection,
      ui: {
        ...current.ui,
        editor: snapshot.editorMode,
        pendingFormat: snapshot.pendingFormat,
        editorRefs: [...snapshot.editorRefs],
        copyRange: snapshot.copy.copyRange,
        copyRanges: snapshot.copy.copyRanges,
        copyMode: snapshot.copy.copyMode,
        copyRevision: snapshot.copy.copyRevision,
        r1c1: snapshot.r1c1,
      },
    }));

    const input = this.mountInput(snapshot.anchor, snapshot.raw, (el) => {
      const max = snapshot.raw.length;
      const start = Math.max(0, Math.min(snapshot.caret.start, max));
      const end = Math.max(start, Math.min(snapshot.caret.end, max));
      el.setSelectionRange(start, end, snapshot.caret.direction);
    });
    mutators.setEditorRefs(this.deps.store, [...snapshot.editorRefs]);
    return input;
  }

  /** Create the editor textarea over `addr`, place its caret, and wire its
   *  listeners, popovers and store tracking. */
  private mountInput(
    addr: Addr,
    raw: string,
    placeCaret: (input: HTMLTextAreaElement) => void,
  ): HTMLTextAreaElement {
    const input = document.createElement('textarea');
    input.className = 'fc-host__editor';
    input.spellcheck = false;
    input.autocapitalize = 'off';
    input.autocomplete = 'off';
    input.rows = 1;
    input.wrap = 'soft';
    input.value = raw;
    this.input = input;
    this.applyTextAlignment(raw);
    this.applyCellAppearance();
    this.position(addr);
    this.deps.grid.appendChild(input);
    this.refreshSize();
    placeCaret(input);

    for (const [type, listener] of this.inputListeners()) {
      input.addEventListener(type, listener as EventListener);
    }
    this.autocomplete = attachAutocomplete({
      input,
      onAfterInsert: () => syncEditorRefs(this.deps.store, input.value),
      getTables: () => this.deps.wb.getTables(),
      editingAddr: addr,
      getColumnValues: (sheet, col, beforeRow) =>
        pickListValues(this.deps.wb.cells(sheet), col, beforeRow),
      getCustomFunctions: () => this.deps.getCustomFunctions?.() ?? [],
      getFunctionNames: () => this.deps.wb.functionNames(),
      labels: this.deps.getLabels?.().autocomplete,
    });
    this.argHelper = attachArgHelper({ input, labels: this.deps.getLabels?.().argHelper });
    this.argHelper.refresh();
    // Scrolling / resizing / freezing moves the cell under the editor. Track
    // the slices that decide the cell rect so the editor stays welded to its
    // cell instead of hanging over whatever scrolled into that spot, and the
    // format slices so a ribbon click mid-edit repaints the editor too.
    this.unsubscribeStore = this.deps.store.subscribe((state, prev) => {
      if (this.editingAddr && (state.viewport !== prev.viewport || state.layout !== prev.layout)) {
        this.position(this.editingAddr);
        this.refreshSize();
      }
      if (state.format !== prev.format || state.ui.pendingFormat !== prev.ui.pendingFormat) {
        this.applyCellAppearance();
        if (this.input) this.applyTextAlignment(this.input.value);
      }
    });
    return input;
  }

  /** DOM listeners the editor textarea carries for its whole lifetime. */
  private inputListeners(): readonly (readonly [string, (event: never) => void])[] {
    return [
      ['keydown', this.onKey],
      ['keyup', this.onKeyUp],
      ['input', this.onInput],
      ['compositionstart', this.onCompositionStart],
      ['compositionend', this.onCompositionEnd],
      ['blur', this.onBlur],
      ['dblclick', this.onDblClick],
      ['click', this.onClick],
    ];
  }

  /** Grow the editor rightward to fit content wider than the cell — desktop
   *  spreadsheet behavior, so a long formula shows from its leading `=`
   *  instead of scrolling to its tail. Capped at the grid's right edge so the
   *  editor never escapes the viewport. */
  private refreshWidth(): void {
    if (!this.input || !this.editingAddr) return;
    const s = this.deps.store.getState();
    const layout = layoutForView(s);
    const r = cellRectUnclamped(layout, s.viewport, this.editingAddr.row, this.editingAddr.col);
    // Reset to the cell width first so the editor can shrink back as content
    //  is deleted, then measure the natural content width.
    this.input.style.width = `${r.w}px`;
    const content = this.input.scrollWidth;
    if (content <= r.w) {
      this.input.classList.remove('fc-host__editor--overflow');
      this.input.classList.remove('fc-host__editor--overflow-rtl');
      return;
    }
    // The editor grows away from the cell's leading edge, which is the right
    // edge on a right-to-left sheet — so there it grows leftwards and its
    // left offset moves with the width.
    const room = layout.rtl ? r.x + r.w - 2 : this.deps.grid.clientWidth - r.x - 2;
    const maxWidth = Math.max(r.w, room);
    const want = content + 2;
    const width = Math.min(want, maxWidth);
    this.input.style.width = `${width}px`;
    if (layout.rtl) this.input.style.left = `${r.x + r.w - width}px`;
    // When the content is still wider than the editor can grow, the text
    //  scrolls; flag the overflow so the editor drops its trailing border.
    this.input.classList.toggle('fc-host__editor--overflow', want > maxWidth && !layout.rtl);
    this.input.classList.toggle('fc-host__editor--overflow-rtl', want > maxWidth && layout.rtl);
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

  /** Give the editor the cell's own fill and text color. Without this the
   *  editor reads as a panel floating above the sheet whenever the cell
   *  carries a fill; with it, the edit happens visually inside the cell. */
  private applyCellAppearance(): void {
    if (!this.input || !this.editingAddr) return;
    const fmt = formatWithPending(this.deps.store.getState(), this.editingAddr);
    // Empty string clears the inline style, falling back to the stylesheet's
    // sheet-background / foreground tokens.
    this.input.style.background = fmt?.fill ?? '';
    this.input.style.color = fmt?.color ?? '';
  }

  private position(a: Addr): void {
    const s = this.deps.store.getState();
    const layout = layoutForView(s);
    const r = cellRectUnclamped(layout, s.viewport, a.row, a.col);
    if (!this.input) return;
    // A cell scrolled off the leading edge sits behind the frozen band and the
    // headers. Park the editor outside the grid's clip rect instead of letting
    // it hang over that chrome — it keeps DOM focus (and the caret, and any IME
    // composition) while its cell is out of view, the way a desktop
    // spreadsheet does.
    const band = bodyBandOrigin(layout, s.viewport);
    const scrolledOut =
      (a.row >= s.layout.freezeRows && r.y < band.y) ||
      (a.col >= s.layout.freezeCols && (layout.rtl ? r.x + r.w > band.x : r.x < band.x));
    this.input.style.left = scrolledOut ? `${-r.w - 4}px` : `${r.x}px`;
    this.input.style.top = scrolledOut ? `${-r.h - 4}px` : `${r.y}px`;
    this.input.style.width = `${r.w}px`;
    this.input.style.height = `${r.h}px`;
  }
}
