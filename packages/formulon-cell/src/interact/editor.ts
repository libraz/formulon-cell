import { autoCorrectInput, correctTypedEntry } from '../commands/autocorrect.js';
import {
  type CellInputAlert,
  commitCellInput,
  inputOperation,
  protectedCellAlert,
  selectRefusedSpan,
} from '../commands/cell-input-commit.js';
import { replaceFormulaSelectionWithWorkbookF9 } from '../commands/f9-preview.js';
import { interactionControllerFor } from '../commands/interaction-controller.js';
import { refusedBatch } from '../commands/interaction-policy.js';
import { PERMISSION_REASONS, type RefusalLike } from '../commands/permission-reasons.js';
import { deferForAllowedRangeUnlock, isCellWritable } from '../commands/protection.js';
import { dblClickRange, editorRefHighlights, extractRefs, rotateRefAt } from '../commands/refs.js';
import { sheetHasCells } from '../engine/sheet-kind.js';
import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { isMacHost } from '../extensions/ui-options.js';
import { rangeOfAddr } from '../store/selection-geometry.js';
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
import { focusSheet } from './dialog-focus.js';
import { caretIndexAtClientX } from './editor-caret.js';
import {
  alignEditorText,
  fitEditorToContent,
  paintEditorCellFormat,
  positionEditorOverCell,
} from './editor-cell-surface.js';
import {
  dismissFormulaAutocomplete,
  handleFormulaAutocompleteKey,
} from './formula-autocomplete-key.js';
import {
  caretSnapshot,
  FormulaDraftHistory,
  type FormulaDraftHistoryState,
  type FormulaDraftMode,
  type FormulaDraftSnapshot,
  historyActionForInputType,
  historyActionForKey,
  planReferenceInsert,
  recordPointNavigation,
} from './formula-draft-history.js';
import {
  applyLeaseStoreState,
  captureLeaseStoreState,
  createFormulaEditLease,
  type FormulaEditLease,
  type FormulaEditLeaseContext,
  type FormulaEditLeaseSnapshot,
  r1c1Active,
} from './formula-edit-lease.js';
import {
  type FormulaReferenceNavigationState,
  navigateFormulaReference,
  pickedReferenceNavigation,
} from './formula-edit-navigation.js';
import { classifyFormulaReferenceArrow } from './formula-reference-arrow.js';
import type { ActiveReferenceGesture } from './formula-reference-gesture.js';
import { isImeKey } from './mac-dialog-keys.js';
import {
  navigationPolicyFor,
  nextTabStop,
  resolveTabBoundary,
  scrollActiveIntoView,
} from './navigation-policy.js';
import {
  buildSelectionInputBatch,
  SELECTION_INPUT_LIMIT_MESSAGE,
  writeSelectionInput,
} from './selection-input.js';
import { advanceAfterCommit, type CommitAdvance } from './selection-navigation.js';
import {
  workbookArgHelperCatalog,
  workbookAutocompleteCatalog,
  workbookCustomFunctions,
} from './workbook-function-catalog.js';

const syncEditorRefs = (store: SpreadsheetStore, text: string, editedAddr: Addr): void => {
  mutators.setEditorRefs(
    store,
    editorRefHighlights(text, editedAddr, r1c1Active(store.getState())),
  );
};

/** A stop alert for a refusal; the alert dialog localizes it through `refusal`. */
const refusalAlert = (refusal: RefusalLike): CellInputAlert => ({
  severity: 'stop',
  message: refusal.reason ?? PERMISSION_REASONS[refusal.reasonKey ?? 'operationDenied'],
  refusal,
});

const operationDeniedAlert = (): CellInputAlert =>
  refusalAlert({ code: 'operationDenied', reasonKey: 'operationDenied' });

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
  onValidation?: (outcome: CellInputAlert) => void;
  getLabels?: () => {
    autocomplete?: Partial<AutocompleteLabels>;
    argHelper?: Partial<ArgHelperLabels>;
  };
  getCustomFunctions?: () => readonly string[];
  getCustomFunctionArgs?: (name: string) => readonly string[] | undefined;
}

export interface EditorBeginOptions {
  mode?: FormulaDraftMode;
  /** Client x of the click that opened the editor; the caret goes to the
   *  character boundary under it instead of the end of the text. */
  caretClientX?: number;
}

/** The part of the inline cell editor hosts may drive. */
export interface CellEditorHandle {
  /** Open the editor on the active cell with `seed` as its text. */
  begin(seed: string, options?: EditorBeginOptions): void;
  /** Whether a cell edit is in progress. */
  isActive(): boolean;
}

/**
 * Inline cell editor — a single-line `<input>` floated over the active
 * cell. Begins on Enter / F2 / printable key. Commits on Enter or Tab,
 * cancels on Escape. Click-outside also commits.
 */

const A1_ATOM = /^(\$?)([A-Za-z]+)(\$?)(\d+)$/;

/** Re-apply the `$` pins of an A1 reference to its moved or resized
 *  replacement: the first corner keeps the first corner's pins, the last the
 *  last's. Text that is not a plain cell or range passes through. */
function keepAbsolutePins(previous: string, next: string): string {
  const pins = previous.split(':').map((atom) => A1_ATOM.exec(atom));
  const parts = next.split(':');
  if (pins.some((m) => !m) || parts.length > 2) return next;
  return parts
    .map((atom, i) => {
      const pin = (i === 0 ? pins[0] : pins[pins.length - 1]) as RegExpExecArray;
      const m = A1_ATOM.exec(atom);
      return m ? `${pin[1]}${m[2]}${pin[3]}${m[4]}` : atom;
    })
    .join(':');
}

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

  private arrowMode: FormulaDraftMode = 'edit';

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

  /** Cell being edited; only valid while an edit is open. */
  editedCell(): Addr {
    if (!this.editingAddr) throw new Error('no cell is being edited');
    return this.editingAddr;
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
    const r1c1 = r1c1Active(this.deps.store.getState());
    const plan = planReferenceInsert({
      value: el.value,
      selectionStart,
      selectionEnd,
      ref,
      phase: options?.phase,
      active: this.referenceGesture,
      r1c1,
    });
    this.referenceGesture = plan?.next ?? null;
    if (!plan) return;
    const { picked, result } = plan;
    const after = caretSnapshot(result.text, result.caret, this.arrowMode);
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

  replaceReferenceAt(index: number, ref: string, options?: { phase: 'start' | 'update' }): void {
    const el = this.input;
    if (!el || !this.editingAddr) return;
    const r1c1 = r1c1Active(this.deps.store.getState());
    const target = extractRefs(el.value, { r1c1, base: this.editingAddr })[index];
    if (!target) return;
    const historyBefore = this.captureDraftSnapshot();
    if (options?.phase === 'start') this.draftHistory.beginGroup(historyBefore, 'reference');
    const current = el.value.slice(target.start, target.end);
    const qualifier = /^(?:'(?:[^']|'')*'|[^'!\s]+)!/.exec(current);
    const pinned = r1c1 ? ref : keepAbsolutePins(current.slice(qualifier?.[0].length ?? 0), ref);
    const text = `${el.value.slice(0, target.start)}${qualifier?.[0] ?? ''}${pinned}${el.value.slice(target.end)}`;
    const caret = target.start + (qualifier?.[0].length ?? 0) + pinned.length;
    this.referenceGesture = null;
    this.pointNavigation = null;
    const after = caretSnapshot(text, caret, this.arrowMode);
    if (options?.phase) this.draftHistory.updateGroup(after, 'reference', historyBefore);
    else this.draftHistory.finishCommand(after);
    el.value = text;
    el.setSelectionRange(caret, caret);
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
      ...captureLeaseStoreState(state),
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
    // A sheet without cells has nothing to edit.
    if (!sheetHasCells(s, a.sheet)) return;
    const cell = rangeOfAddr(a);
    if (deferForAllowedRangeUnlock(this.deps.store, [cell], () => this.begin(seed, options)))
      return;
    // A locked cell of a protected sheet refuses edit mode up front, as the
    // desktop does on a double-click, F2 or the first keystroke.
    if (!isCellWritable(s, a)) {
      this.deps.onValidation?.(protectedCellAlert());
      return;
    }
    const controller = interactionControllerFor(this.deps.store);
    if (controller && controller.policy !== undefined) {
      const operation = inputOperation(seed);
      const permission = controller.canExecute({
        operation,
        origin: 'editor',
        effects: [{ kind: 'cells', cells: [a] }],
      });
      if (!permission.allowed) {
        this.deps.onValidation?.(refusalAlert(permission));
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
      const caret =
        options.caretClientX === undefined
          ? null
          : caretIndexAtClientX(input, options.caretClientX);
      input.setSelectionRange(caret ?? seed.length, caret ?? seed.length);
    });
    this.pendingHistoryInput = null;
    this.ignoredHistoryInput = null;
    this.draftHistory.reset(this.captureDraftSnapshot());
    syncEditorRefs(this.deps.store, seed, a);
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
      mutators.setEditorRefs(this.deps.store, []);
      return;
    }
    this.detachInput(false);
    this.editingAddr = null;
    this.composing = false;
    mutators.setEditor(this.deps.store, { kind: 'idle' });
    mutators.setEditorRefs(this.deps.store, []);
    // Removing the focused input drops focus to <body>; without this, the
    // host's keydown listener stops receiving navigation keys until the
    // user clicks back in.
    focusSheet(this.deps.host, { preventScroll: true });
  }

  commit(advance: CommitAdvance = 'down'): void {
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
      typed: true,
    });
    if (result.status === 'rejected') {
      // Keep the editor open with the offending value so the user can correct.
      this.input.focus();
      if (!selectRefusedSpan(this.input, result.selection)) this.input.select();
      this.deps.onValidation?.(result.alert ?? operationDeniedAlert());
      return;
    }
    if (result.status === 'failed') {
      console.warn('formulon-cell: editor write failed', result.error);
      this.input.focus();
      selectRefusedSpan(this.input, result.selection);
      this.deps.onValidation?.(refusalAlert({ code: 'invalid', reasonKey: 'valueNotWritten' }));
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
    return isMacHost(this.deps.host);
  }

  private advanceAfterCommit(advance: CommitAdvance): void {
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
    const raw = correctTypedEntry(this.deps.store, input.value);
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
        });
      } catch (err) {
        console.warn('formulon-cell: restricted editor fill failed', err);
        result = refusedBatch({
          code: 'invalid',
          reasonKey: 'cellsNotFilled',
          reason: String(err),
        });
      }
      if (result.status === 'refused') {
        input.focus();
        input.select();
        const rejection = result.rejections[0];
        this.deps.onValidation?.(rejection ? refusalAlert(rejection) : operationDeniedAlert());
        return;
      }
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
    const action = historyActionForKey(event);
    if (!action) return false;
    event.preventDefault();
    event.stopPropagation();
    this.applyDraftHistory(action);
    return true;
  }

  private readonly onBeforeInput = (event: InputEvent): void => {
    const action = historyActionForInputType(event.inputType);
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
    if (this.composing || isImeKey(e) || e.key === 'Process') {
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
      if (this.autocomplete) dismissFormulaAutocomplete(this.autocomplete);
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
      this.rotateReferenceAtCaret();
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
        if (
          hasNavigationPolicy &&
          resolveTabBoundary(this.deps.host, this.deps.store) !== 'leave'
        ) {
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
      this.rotateReferenceAtCaret();
    } else if (e.key === 'F9' && this.input && this.editingAddr) {
      const start = this.input.selectionStart ?? 0;
      const result = replaceFormulaSelectionWithWorkbookF9(
        this.deps.wb,
        this.editingAddr,
        this.deps.store.getState().data,
        this.input.value,
        start,
        this.input.selectionEnd ?? start,
      );
      if (result) {
        e.preventDefault();
        e.stopPropagation();
        this.applyRecordedEdit(result.text, result.start, result.end, 'f9');
      }
    }
  };

  /** Replace the draft with a one-step rewrite recorded as its own Undo entry. */
  private applyRecordedEdit(value: string, start: number, end: number, identity: string): void {
    const input = this.input;
    if (!input) return;
    this.draftHistory.recordCommand(
      this.captureDraftSnapshot(),
      { value, start, end, direction: 'none', mode: this.arrowMode },
      identity,
    );
    this.pointNavigation = null;
    this.referenceGesture = null;
    input.value = value;
    input.setSelectionRange(start, end);
    this.syncInputState();
  }

  /** F4 / Cmd+T: cycle the absolute/relative form of the ref under the caret. */
  private rotateReferenceAtCaret(): void {
    const input = this.input;
    const base = this.editingAddr;
    if (!input || !base) return;
    this.draftHistory.endGroup();
    const r = rotateRefAt(input.value, input.selectionStart ?? input.value.length, {
      r1c1: r1c1Active(this.deps.store.getState()),
      base: { row: base.row, col: base.col },
    });
    if (r.text !== input.value) this.applyRecordedEdit(r.text, r.caret, r.caret, 'f4');
  }

  private readonly onKeyUp = (): void => {
    this.refreshHelpersAfterSelection();
  };

  private readonly onBlur = (): void => {
    // Blur commits unless we're already torn down.
    if (this.input) this.commit('none');
  };

  private readonly onInput = (rawEvent?: Event): void => {
    const event = rawEvent as InputEvent | undefined;
    const historyAction = this.pendingHistoryInput ?? historyActionForInputType(event?.inputType);
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
    if (event?.inputType === 'insertText' && this.input) {
      autoCorrectInput(this.deps.store, this.input);
    }
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
      r1c1: r1c1Active(this.deps.store.getState()),
      jump: arrow.jump,
    });
    if (!result) return false;
    event.preventDefault();
    event.stopPropagation();
    this.referenceGesture = null;
    this.pointNavigation = result;
    input.value = result.value;
    recordPointNavigation(this.draftHistory, historyBefore, result, this.arrowMode);
    input.setSelectionRange(result.caret, result.caret);
    this.syncInputState();
    // The programmatic edit must not reopen suggestions that swallow the next arrow.
    if (this.autocomplete) dismissFormulaAutocomplete(this.autocomplete);
    scrollActiveIntoView(this.deps.store, result.tip);
    return true;
  }

  private syncInputState(): void {
    fitEditorToContent(this.input, this.editingAddr, this.deps);
    if (this.input) alignEditorText(this.input, this.editingAddr, this.deps, this.input.value);
    if (this.input && this.editingAddr) {
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
    const start = el.selectionStart ?? el.value.length;
    const end = el.selectionEnd ?? start;
    const caret = start + 1;
    this.applyRecordedEdit(
      `${el.value.slice(0, start)}\n${el.value.slice(end)}`,
      caret,
      caret,
      'newline',
    );
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
      if (focusHost) focusSheet(this.deps.host, { preventScroll: true });
      return;
    }
    for (const [type, listener] of this.inputListeners()) {
      input.removeEventListener(type, listener as EventListener);
    }
    input.remove();
    this.input = null;
    if (focusHost) focusSheet(this.deps.host, { preventScroll: true });
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
    restoredArrowMode: FormulaDraftMode,
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
    applyLeaseStoreState(this.deps.store, snapshot, [...snapshot.editorRefs]);

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
    const customFunctions = workbookCustomFunctions(
      () => this.deps.wb,
      () => addr.sheet,
      {
        list: () => this.deps.getCustomFunctions?.() ?? [],
        args: (name) => this.deps.getCustomFunctionArgs?.(name),
      },
    );
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
      ...customFunctions,
      ...workbookAutocompleteCatalog(() => this.deps.wb, customFunctions.getCustomFunctionArgs),
      groupNames: () => this.isMacPlatform(),
      labels: this.deps.getLabels?.().autocomplete,
    });
    this.argHelper = attachArgHelper({
      input,
      ...workbookArgHelperCatalog(() => this.deps.wb),
      getCustomFunctionArgs: customFunctions.getCustomFunctionArgs,
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
      if (state.format !== prev.format) {
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
