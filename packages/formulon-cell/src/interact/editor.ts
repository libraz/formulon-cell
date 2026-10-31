import { commitCellInput, inputOperation } from '../commands/cell-input-commit.js';
import { inputLocaleForWorkbook } from '../commands/coerce-input.js';
import { replaceFormulaSelectionWithF9Preview } from '../commands/f9-preview.js';
import { interactionControllerFor } from '../commands/interaction-controller.js';
import { deferForAllowedRangeUnlock } from '../commands/protection.js';
import { dblClickRange, extractRefs, rotateRefAt } from '../commands/refs.js';
import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { type ArgHelperHandle, type ArgHelperLabels, attachArgHelper } from './arg-helper.js';
import {
  type AutocompleteHandle,
  type AutocompleteInsertOutcome,
  type AutocompleteLabels,
  attachAutocomplete,
  collectAutocompleteNames,
  pickListValues,
} from './autocomplete.js';
import {
  alignEditorText,
  fitEditorToContent,
  paintEditorCellFormat,
  positionEditorOverCell,
} from './editor-cell-surface.js';
import { handleFormulaAutocompleteKey } from './formula-autocomplete-key.js';
import {
  FormulaDraftHistory,
  type FormulaDraftHistoryState,
  type FormulaDraftSnapshot,
} from './formula-draft-history.js';
import {
  createFormulaEditLease,
  type FormulaEditLease,
  type FormulaEditLeaseContext,
  type FormulaEditLeaseSnapshot,
} from './formula-edit-lease.js';
import {
  type FormulaReferenceNavigationState,
  navigateFormulaReference,
  pickedReferenceNavigation,
} from './formula-edit-navigation.js';
import { resolveFormulaHelperDetails } from './formula-helper-catalog.js';
import { classifyFormulaReferenceArrow } from './formula-reference-arrow.js';
import {
  createFormulaReferenceGesture,
  type FormulaReferenceGesture,
} from './formula-reference-gesture.js';
import { catalogLocaleOrdinal, catalogLocaleTag } from './function-catalog-text.js';
import { navigationPolicyFor, nextTabStop, scrollActiveIntoView } from './navigation-policy.js';
import {
  buildSelectionInputBatch,
  SELECTION_INPUT_LIMIT_MESSAGE,
  writeSelectionInput,
} from './selection-input.js';
import { advanceAfterCommit } from './selection-navigation.js';

const syncEditorRefs = (store: SpreadsheetStore, text: string, editedAddr?: Addr | null): void => {
  const state = store.getState();
  const active = editedAddr ?? state.selection.active;
  const refs = extractRefs(text, {
    r1c1: state.ui.r1c1 === true,
    base: { row: active.row, col: active.col },
  }).map((r) => ({
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

interface ActiveReferenceGesture {
  gesture: FormulaReferenceGesture;
  value: string;
  caret: number;
}

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
  getCustomFunctionArgs?: (name: string) => readonly string[] | undefined;
}

export type FormulaArrowMode = 'edit' | 'point';

export interface EditorBeginOptions {
  mode?: FormulaArrowMode;
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

  private suppressSelectionNotifications = false;

  private readonly draftHistory = new FormulaDraftHistory();

  private pendingHistoryInput: 'undo' | 'redo' | null = null;

  private ignoredHistoryInput: { action: 'undo' | 'redo'; snapshot: FormulaDraftSnapshot } | null =
    null;

  private referenceGesture: ActiveReferenceGesture | null = null;

  private pointNavigation: FormulaReferenceNavigationState | null = null;

  private arrowMode: FormulaArrowMode = 'edit';

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
  insertRefAtCaret(ref: string, options?: { phase: 'start' | 'update' }): void {
    if (!this.input) return;
    const historyBefore = this.captureDraftSnapshot();
    if (options?.phase === 'start') {
      this.draftHistory.beginGroup(historyBefore, 'reference');
    } else if (options?.phase !== 'update') {
      this.draftHistory.endGroup();
      this.draftHistory.beginCommand(historyBefore);
    }
    if (options?.phase !== 'update') this.arrowMode = 'point';
    this.pointNavigation = null;
    const el = this.input;
    const selectionStart = el.selectionStart ?? el.value.length;
    const selectionEnd = el.selectionEnd ?? selectionStart;
    let gesture: FormulaReferenceGesture | null = null;
    if (options?.phase === 'start') {
      gesture = createFormulaReferenceGesture(el.value, selectionStart, selectionEnd, {
        r1c1: this.deps.store.getState().ui.r1c1 === true,
      });
    } else if (options?.phase === 'update') {
      const current = this.referenceGesture;
      if (
        !current ||
        current.value !== el.value ||
        current.caret !== selectionStart ||
        current.caret !== selectionEnd
      ) {
        this.referenceGesture = null;
        return;
      }
      gesture = current.gesture;
    } else {
      this.referenceGesture = null;
    }
    const r1c1 = this.deps.store.getState().ui.r1c1 === true;
    const picked =
      gesture ?? createFormulaReferenceGesture(el.value, selectionStart, selectionEnd, { r1c1 });
    const result = picked.insert(ref);
    this.referenceGesture =
      options?.phase && gesture ? { gesture, value: result.text, caret: result.caret } : null;
    const after: FormulaDraftSnapshot = {
      value: result.text,
      start: result.caret,
      end: result.caret,
      direction: 'none',
      mode: this.arrowMode,
    };
    if (options?.phase) {
      this.draftHistory.updateGroup(after, 'reference', historyBefore);
    } else {
      this.draftHistory.finishCommand(after);
    }
    // The next arrow continues from the picked cell or range.
    if (this.editingAddr)
      this.pointNavigation = pickedReferenceNavigation({
        gesture: picked,
        inserted: result,
        ref,
        editedCell: this.editingAddr,
        r1c1,
      });
    el.value = result.text;
    el.setSelectionRange(result.caret, result.caret);
    el.focus();
    this.syncInputState();
  }

  suspendForFormulaPalette(context: FormulaEditLeaseContext): FormulaEditLease | null {
    const input = this.input;
    const anchor = this.editingAddr;
    if (!input || !anchor || this.composing || this.currentLease) return null;
    const state = this.deps.store.getState();
    const suspendedArrowMode = this.arrowMode;
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

    const suspendedHistory = this.draftHistory.captureState();
    this.detachInput(false, true);
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
      restore: (captured) => this.restoreFromLease(captured, suspendedArrowMode, suspendedHistory),
      release: () => {
        if (this.currentLease === lease) this.currentLease = null;
        if (!this.input && !this.currentLease) this.draftHistory.clear();
      },
    });
    this.currentLease = lease;
    return lease;
  }

  begin(seed: string, options: EditorBeginOptions = {}): void {
    this.invalidateLease();
    this.arrowMode = options.mode ?? 'edit';
    this.pointNavigation = null;
    this.referenceGesture = null;
    const s = this.deps.store.getState();
    const a = s.selection.active;
    const cell = { sheet: a.sheet, r0: a.row, c0: a.col, r1: a.row, c1: a.col };
    if (deferForAllowedRangeUnlock(this.deps.store, [cell], () => this.begin(seed, options)))
      return;
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
    this.pendingHistoryInput = null;
    this.ignoredHistoryInput = null;
    this.draftHistory.reset(this.captureDraftSnapshot());
    syncEditorRefs(this.deps.store, seed, this.editingAddr);
  }

  cancel(): void {
    this.invalidateLease();
    this.referenceGesture = null;
    this.pointNavigation = null;
    this.arrowMode = 'edit';
    this.pendingHistoryInput = null;
    this.ignoredHistoryInput = null;
    this.draftHistory.clear();
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
    this.draftHistory.endGroup();
    this.referenceGesture = null;
    this.pointNavigation = null;
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
    this.draftHistory.endGroup();
    this.referenceGesture = null;
    this.pointNavigation = null;
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

  private captureDraftSnapshot(): FormulaDraftSnapshot {
    const input = this.input;
    const value = input?.value ?? '';
    const start = input?.selectionStart ?? value.length;
    const end = input?.selectionEnd ?? start;
    return {
      value,
      start,
      end,
      direction: input?.selectionDirection ?? 'none',
      mode: this.arrowMode,
    };
  }

  private applyDraftSnapshot(snapshot: FormulaDraftSnapshot): void {
    const input = this.input;
    if (!input) return;
    this.referenceGesture = null;
    this.pointNavigation = null;
    this.arrowMode = snapshot.mode;
    input.value = snapshot.value;
    input.setSelectionRange(snapshot.start, snapshot.end, snapshot.direction);
    this.syncInputState();
  }

  private applyDraftHistory(action: 'undo' | 'redo'): FormulaDraftSnapshot {
    const snapshot =
      (action === 'undo' ? this.draftHistory.undo() : this.draftHistory.redo()) ??
      this.draftHistory.current();
    const input = this.input;
    if (!input) return snapshot;
    this.pendingHistoryInput = null;
    this.ignoredHistoryInput = null;
    this.applyDraftSnapshot(snapshot);
    return snapshot;
  }

  private handleDraftHistoryKey(event: KeyboardEvent): boolean {
    const oneModifier = event.ctrlKey !== event.metaKey;
    if (!oneModifier || event.altKey) return false;
    const key = event.key.toLowerCase();
    const action =
      (key === 'z' && event.shiftKey) || (key === 'y' && !event.shiftKey)
        ? 'redo'
        : key === 'z' && !event.shiftKey
          ? 'undo'
          : null;
    if (!action) return false;
    event.preventDefault();
    event.stopPropagation();
    this.applyDraftHistory(action);
    return true;
  }

  private readonly onBeforeInput = (event: InputEvent): void => {
    const action =
      event.inputType === 'historyUndo'
        ? 'undo'
        : event.inputType === 'historyRedo'
          ? 'redo'
          : null;
    if (action) {
      if (this.composing) return;
      if (event.cancelable) {
        event.preventDefault();
        this.pendingHistoryInput = null;
        this.ignoredHistoryInput = { action, snapshot: this.applyDraftHistory(action) };
      } else {
        this.ignoredHistoryInput = null;
        this.pendingHistoryInput = action;
      }
      return;
    }
    this.pendingHistoryInput = null;
    this.ignoredHistoryInput = null;
    if (!this.composing) this.draftHistory.beforeInput(this.captureDraftSnapshot());
  };

  private readonly onKey = (e: KeyboardEvent): void => {
    if (this.composing || e.isComposing || e.key === 'Process') {
      return;
    }
    if (this.handleDraftHistoryKey(e)) return;
    if (
      e.key === 'F2' &&
      !e.ctrlKey &&
      !e.metaKey &&
      !e.altKey &&
      !e.shiftKey &&
      this.isFormulaEdit()
    ) {
      e.preventDefault();
      e.stopPropagation();
      this.arrowMode = this.arrowMode === 'point' ? 'edit' : 'point';
      this.draftHistory.endGroup();
      this.draftHistory.observeSelection(this.captureDraftSnapshot());
      this.referenceGesture = null;
      this.pointNavigation = null;
      if (this.autocomplete?.dismiss) this.autocomplete.dismiss();
      else this.autocomplete?.close();
      return;
    }
    if (
      this.autocomplete &&
      handleFormulaAutocompleteKey(e, this.autocomplete, { formulaEdit: this.isFormulaEdit() })
    ) {
      e.stopPropagation();
      return;
    }
    if (this.handlePointNavigation(e)) return;
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
        const before = this.captureDraftSnapshot();
        this.draftHistory.endGroup();
        const r = rotateRefAt(this.input.value, caret, {
          r1c1: this.deps.store.getState().ui.r1c1 === true,
          base: this.editingAddr
            ? { row: this.editingAddr.row, col: this.editingAddr.col }
            : undefined,
        });
        if (r.text !== this.input.value) {
          this.draftHistory.recordCommand(
            before,
            {
              value: r.text,
              start: r.caret,
              end: r.caret,
              direction: 'none',
              mode: this.arrowMode,
            },
            'f4',
          );
          this.pointNavigation = null;
          this.referenceGesture = null;
          this.input.value = r.text;
          this.input.setSelectionRange(r.caret, r.caret);
          this.syncInputState();
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
      const before = this.captureDraftSnapshot();
      this.draftHistory.endGroup();
      const r = rotateRefAt(this.input.value, caret, {
        r1c1: this.deps.store.getState().ui.r1c1 === true,
        base: this.editingAddr
          ? { row: this.editingAddr.row, col: this.editingAddr.col }
          : undefined,
      });
      if (r.text !== this.input.value) {
        this.draftHistory.recordCommand(
          before,
          {
            value: r.text,
            start: r.caret,
            end: r.caret,
            direction: 'none',
            mode: this.arrowMode,
          },
          'f4',
        );
        this.pointNavigation = null;
        this.referenceGesture = null;
        this.input.value = r.text;
        this.input.setSelectionRange(r.caret, r.caret);
        this.syncInputState();
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
        (formula) =>
          this.deps.wb.evaluateFormulaText(
            active,
            this.deps.wb.canonicalizeLocalizedFormula(formula),
          ),
        true,
        (formula) =>
          this.deps.wb.evaluateFormulaArray(
            active,
            this.deps.wb.canonicalizeLocalizedFormula(formula),
          ),
        {
          localizeLiteral: (literal) => {
            const localized = this.deps.wb.localizeFormulaText(`=${literal}`);
            return localized.startsWith('=') ? localized.slice(1) : literal;
          },
        },
      );
      if (result) {
        e.preventDefault();
        e.stopPropagation();
        const before = this.captureDraftSnapshot();
        this.draftHistory.endGroup();
        this.draftHistory.recordCommand(
          before,
          {
            value: result.text,
            start: result.start,
            end: result.end,
            direction: 'none',
            mode: this.arrowMode,
          },
          'f9',
        );
        this.pointNavigation = null;
        this.referenceGesture = null;
        this.input.value = result.text;
        this.input.setSelectionRange(result.start, result.end);
        this.syncInputState();
      }
    }
  };

  private readonly onKeyUp = (): void => {
    this.refreshHelpersAfterSelection();
  };

  private readonly onBlur = (): void => {
    // Blur commits unless we're already torn down.
    if (this.input) this.commit('none');
  };

  private readonly onInput = (rawEvent?: Event): void => {
    const event = rawEvent as InputEvent | undefined;
    const historyAction =
      this.pendingHistoryInput ??
      (event?.inputType === 'historyUndo'
        ? 'undo'
        : event?.inputType === 'historyRedo'
          ? 'redo'
          : null);
    if (this.composing) {
      this.pendingHistoryInput = null;
      this.ignoredHistoryInput = null;
      this.referenceGesture = null;
      this.pointNavigation = null;
      this.draftHistory.input(this.captureDraftSnapshot(), {
        composing: true,
        inputType: event?.inputType,
      });
      this.syncInputState();
      return;
    }
    const ignored = this.ignoredHistoryInput;
    if (historyAction && ignored?.action === historyAction) {
      this.ignoredHistoryInput = null;
      this.pendingHistoryInput = null;
      this.applyDraftSnapshot(ignored.snapshot);
      return;
    }
    this.ignoredHistoryInput = null;
    if (historyAction) {
      this.pendingHistoryInput = null;
      this.applyDraftHistory(historyAction);
      return;
    }
    this.referenceGesture = null;
    this.pointNavigation = null;
    this.draftHistory.input(this.captureDraftSnapshot(), {
      composing: this.composing,
      inputType: event?.inputType,
    });
    this.syncInputState();
  };

  private handlePointNavigation(event: KeyboardEvent): boolean {
    if (
      this.arrowMode !== 'point' ||
      !this.isFormulaEdit() ||
      this.composing ||
      event.isComposing ||
      event.key === 'Process'
    )
      return false;
    const arrow = classifyFormulaReferenceArrow(event, this.isMacPlatform());
    const input = this.input;
    const editedCell = this.editingAddr;
    if (!arrow || !input || !editedCell) return false;
    const historyBefore = this.captureDraftSnapshot();
    const current = this.pointNavigation;
    const result = navigateFormulaReference({
      state: this.deps.store.getState(),
      text: input.value,
      selectionStart: input.selectionStart ?? input.value.length,
      selectionEnd: input.selectionEnd ?? input.value.length,
      editedCell,
      direction: arrow.direction,
      extend: arrow.extend,
      current,
      r1c1: this.deps.store.getState().ui.r1c1 === true,
      jump: arrow.jump,
    });
    if (!result) return false;
    event.preventDefault();
    event.stopPropagation();
    this.referenceGesture = null;
    this.pointNavigation = result;
    input.value = result.value;
    if (!this.draftHistory.hasActiveGroup) {
      this.draftHistory.beginGroup(historyBefore, 'reference');
    }
    this.draftHistory.updateGroup(
      {
        value: result.value,
        start: result.caret,
        end: result.caret,
        direction: 'none',
        mode: this.arrowMode,
      },
      'reference',
      historyBefore,
    );
    input.setSelectionRange(result.caret, result.caret);
    this.syncInputState();
    // The programmatic edit must not reopen suggestions that swallow the next arrow.
    if (this.autocomplete?.dismiss) this.autocomplete.dismiss();
    else this.autocomplete?.close();
    scrollActiveIntoView(this.deps.store, result.tip);
    return true;
  }

  private syncInputState(): void {
    fitEditorToContent(this.input, this.editingAddr, this.deps);
    if (this.input) alignEditorText(this.input, this.editingAddr, this.deps, this.input.value);
    if (this.input) {
      mutators.setEditor(this.deps.store, {
        kind: 'edit',
        raw: this.input.value,
        caret: this.input.selectionStart ?? this.input.value.length,
      });
      syncEditorRefs(this.deps.store, this.input.value, this.editingAddr);
    }
    this.autocomplete?.refresh();
    this.argHelper?.refresh();
  }

  private readonly onCompositionStart = (): void => {
    this.composing = true;
    this.pendingHistoryInput = null;
    this.ignoredHistoryInput = null;
    this.draftHistory.startComposition(this.captureDraftSnapshot());
    this.pointNavigation = null;
    this.referenceGesture = null;
    this.autocomplete?.setComposing?.(true);
    this.argHelper?.setComposing?.(true);
  };

  private readonly onCompositionEnd = (): void => {
    this.composing = false;
    this.draftHistory.endComposition(this.captureDraftSnapshot());
    this.autocomplete?.setComposing?.(false);
    this.argHelper?.setComposing?.(false);
    this.syncInputState();
  };

  private readonly refreshHelpersAfterSelection = (): void => {
    if (this.suppressSelectionNotifications) return;
    const input = this.input;
    if (input && !this.composing) {
      this.draftHistory.observeSelection(this.captureDraftSnapshot());
    }
    const gesture = this.referenceGesture;
    if (
      gesture &&
      (!input ||
        input.value !== gesture.value ||
        input.selectionStart !== gesture.caret ||
        input.selectionEnd !== gesture.caret)
    ) {
      this.referenceGesture = null;
    }
    const pointNavigation = this.pointNavigation;
    if (
      pointNavigation &&
      (!input ||
        input.value !== pointNavigation.value ||
        input.selectionStart !== pointNavigation.caret ||
        input.selectionEnd !== pointNavigation.caret)
    ) {
      this.pointNavigation = null;
    }
    this.autocomplete?.refresh();
    this.argHelper?.refresh();
  };

  /** Double-click inside a formula edit selects a semantic token rather than
   *  a bare word — the function name, or a whole call argument (so a range
   *  like `F4:F8` selects as a unit). Non-formula edits keep the browser's
   *  native word selection. */
  private readonly onDblClick = (e: MouseEvent): void => {
    this.draftHistory.endGroup();
    this.arrowMode = 'edit';
    this.pointNavigation = null;
    this.referenceGesture = null;
    const el = this.input;
    if (!el?.value.startsWith('=')) {
      this.refreshHelpersAfterSelection();
      return;
    }
    // The browser has already selected a word; probe its midpoint so the hit
    //  point lands inside the intended token.
    const a = el.selectionStart ?? 0;
    const b = el.selectionEnd ?? a;
    const probe = a === b ? a : Math.floor((a + b) / 2);
    const range = dblClickRange(el.value, probe, this.deps.wb.localeFacts()?.listSeparator);
    if (range) {
      e.preventDefault();
      el.setSelectionRange(range.start, range.end);
    }
    this.refreshHelpersAfterSelection();
  };

  /** Triple-click inside a formula edit selects the whole formula. */
  private readonly onClick = (e: MouseEvent): void => {
    this.draftHistory.endGroup();
    this.arrowMode = 'edit';
    this.pointNavigation = null;
    this.referenceGesture = null;
    const el = this.input;
    if (!el || e.detail < 3 || !el.value.startsWith('=')) {
      this.refreshHelpersAfterSelection();
      return;
    }
    e.preventDefault();
    el.setSelectionRange(0, el.value.length);
    this.refreshHelpersAfterSelection();
  };

  private insertNewline(): void {
    const el = this.input;
    if (!el) return;
    const before = this.captureDraftSnapshot();
    this.draftHistory.endGroup();
    const start = el.selectionStart ?? el.value.length;
    const end = el.selectionEnd ?? start;
    const value = `${el.value.slice(0, start)}\n${el.value.slice(end)}`;
    const caret = start + 1;
    this.draftHistory.recordCommand(
      before,
      {
        value,
        start: caret,
        end: caret,
        direction: 'none',
        mode: this.arrowMode,
      },
      'newline',
    );
    el.value = value;
    el.setSelectionRange(caret, caret);
    this.referenceGesture = null;
    this.pointNavigation = null;
    this.syncInputState();
  }

  /** Remove the transient editor DOM and its subscriptions without changing
   *  the editor slices. A palette lease uses this boundary before blur can
   *  turn a suspended edit into a workbook write; ordinary cancel asks for
   *  host focus after the same teardown. */
  private detachInput(focusHost: boolean, preserveHistory = false): void {
    this.unsubscribeStore?.();
    this.unsubscribeStore = null;
    this.autocomplete?.detach();
    this.autocomplete = null;
    this.argHelper?.detach();
    this.argHelper = null;
    this.referenceGesture = null;
    this.pointNavigation = null;
    const input = this.input;
    this.pendingHistoryInput = null;
    this.ignoredHistoryInput = null;
    if (!preserveHistory) this.draftHistory.clear();
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
  private restoreFromLease(
    snapshot: Readonly<FormulaEditLeaseSnapshot>,
    restoredArrowMode: FormulaArrowMode,
    restoredHistory: FormulaDraftHistoryState,
  ): HTMLElement | null {
    if (this.input || this.deps.wb !== snapshot.workbook) return null;
    const state = this.deps.store.getState();
    if (state.data.sheetIndex !== snapshot.anchor.sheet || state.ui.r1c1 !== snapshot.r1c1) {
      return null;
    }
    this.editingAddr = { ...snapshot.anchor };
    this.editBaseline = snapshot.baseline;
    this.composing = false;
    this.arrowMode = restoredArrowMode;
    this.pointNavigation = null;
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
    this.pendingHistoryInput = null;
    this.ignoredHistoryInput = null;
    this.draftHistory.restoreState(restoredHistory);
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
    alignEditorText(this.input, this.editingAddr, this.deps, raw);
    paintEditorCellFormat(this.input, this.editingAddr, this.deps);
    positionEditorOverCell(this.input, addr, this.deps);
    this.deps.grid.appendChild(input);
    fitEditorToContent(this.input, this.editingAddr, this.deps);
    placeCaret(input);

    for (const [type, listener] of this.inputListeners()) {
      input.addEventListener(type, listener as EventListener);
    }
    this.autocomplete = attachAutocomplete({
      input,
      onBeforeInsert: () => {
        this.suppressSelectionNotifications = true;
        this.draftHistory.beginCommand(this.captureDraftSnapshot());
      },
      onAfterInsert: (outcome?: AutocompleteInsertOutcome) => {
        this.suppressSelectionNotifications = false;
        this.referenceGesture = null;
        this.pointNavigation = null;
        if (outcome?.kind === 'function') this.arrowMode = 'point';
        this.draftHistory.finishCommand(this.captureDraftSnapshot());
        this.syncInputState();
      },
      getTables: () => this.deps.wb.getTables(),
      getNames: () =>
        collectAutocompleteNames(this.deps.wb.definedNames(), this.deps.wb.getTables(), addr.sheet),
      editingAddr: addr,
      getColumnValues: (sheet, col, beforeRow) =>
        pickListValues(this.deps.wb.cells(sheet), col, beforeRow),
      getCustomFunctions: () => this.deps.getCustomFunctions?.() ?? [],
      getCustomFunctionArgs: (name) => this.deps.getCustomFunctionArgs?.(name),
      getFunctionNames: () => this.deps.wb.functionNames(),
      getFunctionDisplayNames: () => {
        const wb = this.deps.wb;
        const names = wb.functionNames();
        if (!names) return null;
        const locale = catalogLocaleOrdinal(inputLocaleForWorkbook(wb) ?? 'en-US');
        return names.map((name) => wb.localizeFunctionName(name, locale) || name);
      },
      resolveFunctionName: (rawName) => {
        const locale = catalogLocaleOrdinal(inputLocaleForWorkbook(this.deps.wb) ?? 'en-US');
        return (
          this.deps.wb.canonicalizeFunctionName(rawName, locale) || rawName.toLocaleUpperCase()
        );
      },
      getListSeparator: () => this.deps.wb.localeFacts()?.listSeparator ?? ',',
      getFunctionDetails: (canonicalName) =>
        resolveFormulaHelperDetails(canonicalName, {
          localeTag: catalogLocaleTag(inputLocaleForWorkbook(this.deps.wb) ?? 'en-US'),
          listSeparator: this.deps.wb.localeFacts()?.listSeparator ?? ',',
          getMetadata: (name, locale) => this.deps.wb.functionMetadataForLocale(name, locale),
          localizeName: (name) => {
            const locale = catalogLocaleOrdinal(inputLocaleForWorkbook(this.deps.wb) ?? 'en-US');
            return this.deps.wb.localizeFunctionName(name, locale) || name;
          },
          getCustomFunctionArgs: (name) => this.deps.getCustomFunctionArgs?.(name),
        }),
      labels: this.deps.getLabels?.().autocomplete,
    });
    this.argHelper = attachArgHelper({
      input,
      getListSeparator: () => this.deps.wb.localeFacts()?.listSeparator ?? ',',
      getFunctionDisplayName: (canonicalName) => {
        const wb = this.deps.wb;
        const localeTag = catalogLocaleTag(inputLocaleForWorkbook(wb) ?? 'en-US');
        const metadataName = wb.functionMetadataForLocale(canonicalName, localeTag)?.localizedName;
        return (
          metadataName ||
          wb.localizeFunctionName(canonicalName, catalogLocaleOrdinal(localeTag)) ||
          canonicalName
        );
      },
      resolveFunctionName: (rawName) => {
        const locale = catalogLocaleOrdinal(inputLocaleForWorkbook(this.deps.wb) ?? 'en-US');
        return (
          this.deps.wb.canonicalizeFunctionName(rawName, locale) || rawName.toLocaleUpperCase()
        );
      },
      getCustomFunctionArgs: (name) => this.deps.getCustomFunctionArgs?.(name),
      labels: this.deps.getLabels?.().argHelper,
    });
    this.argHelper.refresh();
    // Scrolling / resizing / freezing moves the cell under the editor. Track
    // the slices that decide the cell rect so the editor stays welded to its
    // cell instead of hanging over whatever scrolled into that spot, and the
    // format slices so a ribbon click mid-edit repaints the editor too.
    this.unsubscribeStore = this.deps.store.subscribe((state, prev) => {
      if (this.editingAddr && (state.viewport !== prev.viewport || state.layout !== prev.layout)) {
        positionEditorOverCell(this.input, this.editingAddr, this.deps);
        fitEditorToContent(this.input, this.editingAddr, this.deps);
        paintEditorCellFormat(this.input, this.editingAddr, this.deps);
      }
      if (state.format !== prev.format || state.ui.pendingFormat !== prev.ui.pendingFormat) {
        paintEditorCellFormat(this.input, this.editingAddr, this.deps);
        if (this.input) alignEditorText(this.input, this.editingAddr, this.deps, this.input.value);
      }
    });
    return input;
  }

  /** DOM listeners the editor textarea carries for its whole lifetime. */
  private inputListeners(): readonly (readonly [string, (event: never) => void])[] {
    return [
      ['keydown', this.onKey],
      ['keyup', this.onKeyUp],
      ['beforeinput', this.onBeforeInput],
      ['select', this.refreshHelpersAfterSelection],
      ['mouseup', this.refreshHelpersAfterSelection],
      ['input', this.onInput],
      ['compositionstart', this.onCompositionStart],
      ['compositionend', this.onCompositionEnd],
      ['blur', this.onBlur],
      ['dblclick', this.onDblClick],
      ['click', this.onClick],
    ];
  }
}
