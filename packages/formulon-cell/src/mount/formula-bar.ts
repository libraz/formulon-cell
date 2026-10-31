import { commitCellInput } from '../commands/cell-input-commit.js';
import { replaceFormulaSelectionWithF9Preview } from '../commands/f9-preview.js';
import { interactionControllerFor } from '../commands/interaction-controller.js';
import { deferForAllowedRangeUnlock, selectionEditRanges } from '../commands/protection.js';
import { extractRefs, rotateRefAt } from '../commands/refs.js';
import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import type { AutocompleteInsertOutcome } from '../interact/autocomplete.js';
import { handleFormulaAutocompleteKey } from '../interact/formula-autocomplete-key.js';
import {
  FormulaDraftHistory,
  type FormulaDraftHistoryState,
  type FormulaDraftSnapshot,
} from '../interact/formula-draft-history.js';
import {
  createFormulaEditLease,
  type FormulaEditLease,
  type FormulaEditLeaseContext,
  type FormulaEditLeaseSnapshot,
} from '../interact/formula-edit-lease.js';
import {
  type FormulaReferenceNavigationState,
  navigateFormulaReference,
  pickedReferenceNavigation,
} from '../interact/formula-edit-navigation.js';
import { classifyFormulaReferenceArrow } from '../interact/formula-reference-arrow.js';
import {
  createFormulaReferenceGesture,
  type FormulaReferenceGesture,
} from '../interact/formula-reference-gesture.js';
import {
  isNavigationAddrAllowed,
  navigationBoundsFor,
  navigationSelectionBoundsFor,
  nextTabStop,
  scrollActiveIntoView,
} from '../interact/navigation-policy.js';
import {
  buildSelectionInputBatch,
  SELECTION_INPUT_LIMIT_MESSAGE,
  writeSelectionInput,
} from '../interact/selection-input.js';
import { advanceAfterCommit } from '../interact/selection-navigation.js';
import { sameAddr } from '../store/pending-format.js';
import type { SpreadsheetStore } from '../store/store.js';
import { mutators } from '../store/store.js';
import type { EditorRefHighlight } from '../store/types.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';

interface FormulaBarAutocomplete {
  isOpen(): boolean;
  move(n: number): void;
  acceptHighlighted(): boolean;
  close(): void;
  refresh(): void;
  navigate?(action: 'first' | 'last' | 'pageUp' | 'pageDown'): void;
  dismiss?(): void;
  toggle?(): void;
  setComposing?(composing: boolean): void;
}

interface FormulaArgHelper {
  refresh(): void;
  close?(): void;
  setComposing?(composing: boolean): void;
}

interface ActiveReferenceGesture {
  gesture: FormulaReferenceGesture;
  value: string;
  caret: number;
}

type FormulaBarArrowMode = 'edit' | 'point';

interface AttachFormulaBarInput {
  formulabar: HTMLElement;
  fxAccept: HTMLButtonElement;
  fxCancel: HTMLButtonElement;
  fxInput: HTMLTextAreaElement;
  getArgHelper: () => FormulaArgHelper | null;
  getAutocomplete: () => FormulaBarAutocomplete;
  getStrings: () => Strings;
  cancelBindingEditor: () => void;
  host: HTMLElement;
  onValidation?: (outcome: {
    severity: 'stop' | 'warning' | 'information';
    title?: string;
    message: string;
  }) => void;
  store: SpreadsheetStore;
  updateChrome: () => void;
  wb: () => WorkbookHandle;
}

interface ExternalFormulaDraftView {
  raw: string;
  caret: {
    start: number;
    end: number;
    direction: HTMLTextAreaElement['selectionDirection'];
  };
}

type ExternalFormulaDraftOutcome = 'committed' | 'cancelled' | 'discarded';

export interface ExternalFormulaDraftHooks {
  onFinish(outcome: ExternalFormulaDraftOutcome, restoredFocusTarget?: HTMLElement | null): void;
}

export interface ExternalFormulaDraftHandle {
  readonly anchor: Addr;
  value(): string;
  snapshot(): ExternalFormulaDraftView | null;
  setValue(raw: string, caret?: number): void;
  commit(): boolean;
  cancel(): void;
  discard(): void;
  subscribe(fn: (view: ExternalFormulaDraftView) => void): () => void;
}

type ExternalDraftStatus = 'open' | 'committing' | 'committed' | 'cancelled' | 'discarded';

interface ExternalDraftState {
  anchor: Addr;
  workbook: WorkbookHandle;
  restoreValue: string;
  hooks: ExternalFormulaDraftHooks;
  listeners: Set<(view: ExternalFormulaDraftView) => void>;
  lastNotifiedView: ExternalFormulaDraftView | null;
  status: ExternalDraftStatus;
  cancelRequested: boolean;
  cancelMode: 'restore' | 'discard' | null;
  lease: FormulaEditLease | null;
}

export interface FormulaBarController {
  acceptFx(): void;
  beginExternalDraft(
    anchor: Addr,
    seed: string,
    hooks: ExternalFormulaDraftHooks,
    options?: { lease?: FormulaEditLease },
  ): ExternalFormulaDraftHandle | null;
  suspendForFormulaPalette(context: FormulaEditLeaseContext): FormulaEditLease | null;
  cancelFx(): void;
  discardFx(): void;
  commitFx(advance: 'down' | 'right' | 'up' | 'left' | 'none'): boolean;
  detach(): void;
  insertRefAtCaret(ref: string, options?: { phase: 'start' | 'update' }): void;
  beforeAutocompleteInsert?(): void;
  notifyAutocompleteInsert?(outcome?: AutocompleteInsertOutcome): void;
  isEditing(): boolean;
  isFormulaEdit(): boolean;
  refreshActions(): void;
  syncFxRefs(): void;
}

export function attachFormulaBarController(input: AttachFormulaBarInput): FormulaBarController {
  const {
    formulabar,
    fxAccept,
    fxCancel,
    fxInput,
    getArgHelper,
    getAutocomplete,
    getStrings,
    cancelBindingEditor,
    host,
    onValidation,
    store,
    updateChrome,
    wb,
  } = input;
  let fxEditing = false;
  let fxBaseline = '';
  let editingAnchor: Addr | null = null;
  let composing = false;
  let referenceGesture: ActiveReferenceGesture | null = null;
  let pointNavigation: FormulaReferenceNavigationState | null = null;
  let arrowMode: FormulaBarArrowMode = 'edit';
  let leaseGeneration = 0;
  let currentLease: FormulaEditLease | null = null;
  let suspended = false;
  let commitRejected = false;
  /** A commit waits for a range password; the prompt's focus change must not re-commit. */
  let commitDeferred = false;
  let externalDraft: ExternalDraftState | null = null;
  let detached = false;
  let suppressSelectionNotifications = false;
  let ownedEditorRefs: EditorRefHighlight[] | null = null;
  const draftHistory = new FormulaDraftHistory();
  let pendingHistoryInput: 'undo' | 'redo' | null = null;
  let ignoredHistoryInput: { action: 'undo' | 'redo'; snapshot: FormulaDraftSnapshot } | null =
    null;

  const isMacPlatform = (): boolean =>
    (host.closest<HTMLElement>('.fc-host') ?? host).dataset.fcPlatform === 'mac';

  const isFormulaEdit = (): boolean => fxEditing && fxInput.value.trimStart().startsWith('=');
  const draftActive = (): boolean => fxEditing || externalDraft?.status === 'open';

  const captureDraftSnapshot = (): FormulaDraftSnapshot => {
    const value = fxInput.value;
    const start = fxInput.selectionStart ?? value.length;
    const end = fxInput.selectionEnd ?? start;
    return {
      value,
      start,
      end,
      direction: fxInput.selectionDirection,
      mode: arrowMode,
    };
  };

  const applyDraftSnapshot = (snapshot: FormulaDraftSnapshot): void => {
    referenceGesture = null;
    pointNavigation = null;
    arrowMode = snapshot.mode;
    fxInput.value = snapshot.value;
    fxInput.setSelectionRange(snapshot.start, snapshot.end, snapshot.direction);
    commitRejected = false;
    syncFormulaBarInputState();
  };

  const applyDraftHistory = (action: 'undo' | 'redo'): FormulaDraftSnapshot => {
    const snapshot =
      (action === 'undo' ? draftHistory.undo() : draftHistory.redo()) ?? draftHistory.current();
    pendingHistoryInput = null;
    ignoredHistoryInput = null;
    applyDraftSnapshot(snapshot);
    return snapshot;
  };

  const handleDraftHistoryKey = (event: KeyboardEvent): boolean => {
    if (!draftActive()) return false;
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
    applyDraftHistory(action);
    return true;
  };

  const refreshActions = (): void => {
    const dirty = fxEditing && fxInput.value !== fxBaseline;
    const strings = getStrings().a11y;
    const cancelReason = fxEditing ? null : strings.cancelFormulaEditUnavailable;
    const acceptReason = dirty
      ? null
      : fxEditing
        ? strings.enterFormulaNoChanges
        : strings.enterFormulaUnavailable;
    projectDisabledState(fxCancel, !fxEditing, cancelReason, {
      datasetKey: 'disabledReason',
      titlePrefix: strings.cancelFormulaEdit,
    });
    projectDisabledState(fxAccept, !dirty, acceptReason, {
      datasetKey: 'disabledReason',
      titlePrefix: strings.enterFormula,
    });
    formulabar.dataset.fcEditing = fxEditing ? '1' : '0';
  };

  const syncFxRefs = (): void => {
    if (detached || (externalDraft && externalDraft.status !== 'open')) return;
    const active = editingAnchor ?? externalDraft?.anchor ?? store.getState().selection.active;
    const refs = extractRefs(fxInput.value, {
      r1c1: store.getState().ui.r1c1 === true,
      base: { row: active.row, col: active.col },
    }).map((r) => ({
      r0: r.r0,
      c0: r.c0,
      r1: r.r1,
      c1: r.c1,
      colorIndex: r.colorIndex,
    }));
    ownedEditorRefs = refs;
    mutators.setEditorRefs(store, refs);
  };

  const clearFxRefs = (): void => {
    const owned = ownedEditorRefs;
    // Relinquish before the mutator so a synchronous subscriber can establish
    // a newer owner without a trailing cleanup clobbering its refs.
    ownedEditorRefs = null;
    if (!owned || store.getState().ui.editorRefs !== owned) return;
    mutators.setEditorRefs(store, []);
  };

  const closeFormulaHelpers = (): void => {
    getAutocomplete().close();
    getArgHelper()?.close?.();
  };

  const finishEditing = (): void => {
    referenceGesture = null;
    pointNavigation = null;
    arrowMode = 'edit';
    fxEditing = false;
    editingAnchor = null;
    commitRejected = false;
    fxBaseline = fxInput.value;
    pendingHistoryInput = null;
    ignoredHistoryInput = null;
    draftHistory.clear();
    refreshActions();
    clearFxRefs();
    closeFormulaHelpers();
  };

  const restoreFromLease = (
    snapshot: Readonly<FormulaEditLeaseSnapshot>,
    restoredArrowMode: FormulaBarArrowMode,
    restoredHistory: FormulaDraftHistoryState,
  ): HTMLElement | null => {
    if (detached || wb() !== snapshot.workbook) return null;
    const state = store.getState();
    if (state.data.sheetIndex !== snapshot.anchor.sheet || state.ui.r1c1 !== snapshot.r1c1) {
      return null;
    }
    const restoredRefs = [...snapshot.editorRefs];
    ownedEditorRefs = restoredRefs;
    store.setState((current) => ({
      ...current,
      selection: snapshot.selection,
      ui: {
        ...current.ui,
        editor: snapshot.editorMode,
        pendingFormat: snapshot.pendingFormat,
        editorRefs: restoredRefs,
        copyRange: snapshot.copy.copyRange,
        copyRanges: snapshot.copy.copyRanges,
        copyMode: snapshot.copy.copyMode,
        copyRevision: snapshot.copy.copyRevision,
        r1c1: snapshot.r1c1,
      },
    }));
    fxInput.value = snapshot.raw;
    fxInput.setSelectionRange(
      Math.max(0, Math.min(snapshot.caret.start, snapshot.raw.length)),
      Math.max(0, Math.min(snapshot.caret.end, snapshot.raw.length)),
      snapshot.caret.direction,
    );
    fxBaseline = snapshot.baseline;
    editingAnchor = { ...snapshot.anchor };
    fxEditing = true;
    arrowMode = restoredArrowMode;
    pointNavigation = null;
    suspended = false;
    commitRejected = false;
    pendingHistoryInput = null;
    ignoredHistoryInput = null;
    draftHistory.restoreState(restoredHistory);
    refreshActions();
    mutators.setEditorRefs(store, restoredRefs);
    getAutocomplete().refresh();
    getArgHelper()?.refresh();
    return fxInput;
  };

  const invalidateLease = (): void => {
    leaseGeneration += 1;
    const lease = currentLease;
    currentLease = null;
    lease?.discard();
    suspended = false;
  };

  const suspendForFormulaPalette = (context: FormulaEditLeaseContext): FormulaEditLease | null => {
    if (detached || !fxEditing || composing || externalDraft || currentLease) return null;
    const anchor = editingAnchor ?? store.getState().selection.active;
    const state = store.getState();
    const generation = ++leaseGeneration;
    const suspendedArrowMode = arrowMode;
    const suspendedHistory = draftHistory.captureState();
    const workbook = wb();
    const sheetCount = workbook.sheetCount;
    const snapshot: FormulaEditLeaseSnapshot = {
      source: 'formulaBar',
      workbook,
      anchor: { ...anchor },
      raw: fxInput.value,
      baseline: fxBaseline,
      caret: {
        start: fxInput.selectionStart ?? fxInput.value.length,
        end: fxInput.selectionEnd ?? fxInput.value.length,
        direction: fxInput.selectionDirection,
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
    suspended = true;
    clearFxRefs();
    closeFormulaHelpers();
    mutators.setEditor(store, { kind: 'idle' });
    let lease: FormulaEditLease | null = null;
    lease = createFormulaEditLease(snapshot, {
      context,
      isOwnerCurrent: () => {
        if (currentLease !== lease || leaseGeneration !== generation || detached) return false;
        const current = store.getState();
        return (
          wb() === snapshot.workbook &&
          wb().sheetCount === sheetCount &&
          current.data.sheetIndex === snapshot.anchor.sheet &&
          current.ui.r1c1 === snapshot.r1c1
        );
      },
      restore: (captured) => restoreFromLease(captured, suspendedArrowMode, suspendedHistory),
      release: () => {
        if (currentLease === lease) currentLease = null;
        if (suspended) {
          // Go inert without finishEditing: a newer owner may already hold the store.
          suspended = false;
          draftHistory.clear();
          fxEditing = false;
          editingAnchor = null;
          commitRejected = false;
          fxBaseline = fxInput.value;
        }
      },
    });
    currentLease = lease;
    return lease;
  };

  const externalDraftWorkbookCurrent = (draft: ExternalDraftState): boolean => {
    try {
      const current = wb();
      return (
        current === draft.workbook &&
        Number.isInteger(current.sheetCount) &&
        current.sheetCount >= 0
      );
    } catch {
      return false;
    }
  };

  const externalDraftStale = (draft: ExternalDraftState): boolean =>
    !externalDraftWorkbookCurrent(draft) || (draft.lease !== null && !draft.lease.valid());

  const currentExternalDraftView = (draft: ExternalDraftState): ExternalFormulaDraftView | null => {
    if (detached || externalDraft !== draft || draft.status !== 'open') return null;
    if (externalDraftStale(draft)) {
      finishExternalDraft(draft, 'discarded', false, false);
      return null;
    }
    return {
      raw: fxInput.value,
      caret: {
        start: fxInput.selectionStart ?? fxInput.value.length,
        end: fxInput.selectionEnd ?? fxInput.value.length,
        direction: fxInput.selectionDirection,
      },
    };
  };

  const makeExternalDraftInert = (): void => {
    // Drop only the discarded draft's local ownership markers. Clear its refs
    // by identity so a newer owner survives a synchronous discard re-entry.
    fxEditing = false;
    editingAnchor = null;
    commitRejected = false;
    fxBaseline = fxInput.value;
    suspended = false;
    referenceGesture = null;
    clearFxRefs();
    pendingHistoryInput = null;
    ignoredHistoryInput = null;
    draftHistory.clear();
  };

  const refreshDiscardedChrome = (): void => {
    // Discard is a system boundary: clear only this controller's local chrome
    // state, then project the context that is current now. Ref ownership was
    // handled before entering this helper, so newer editor refs survive.
    try {
      closeFormulaHelpers();
    } catch {
      // A helper may already be detached while the host is being torn down.
    }
    if (!detached) {
      try {
        const current = wb();
        if (Number.isInteger(current.sheetCount) && current.sheetCount >= 0) {
          updateChrome();
          fxBaseline = fxInput.value;
        }
      } catch {
        // Workbook teardown can race a stale draft; leave the callback path live.
      }
    }
    try {
      refreshActions();
    } catch {
      // Keep terminal callbacks reachable if local chrome has already detached.
    }
  };

  function finishExternalDraft(
    draft: ExternalDraftState,
    requestedOutcome: ExternalFormulaDraftOutcome,
    restore: boolean,
    focusHost: boolean,
  ): void {
    if (draft.status !== 'open' && draft.status !== 'committing') return;

    let outcome = requestedOutcome;
    // Invalidity takes precedence over a user-style restore, including when a
    // synchronous callback invalidates the lease while the write is settling.
    if (outcome === 'cancelled' && restore && draft.lease && !draft.lease.valid()) {
      outcome = 'discarded';
      restore = false;
      focusHost = false;
    }

    draft.status = outcome;
    // Clear the binding and subscribers before releasing the lease or invoking
    // hooks so synchronous re-entry can only observe a terminal draft.
    if (externalDraft === draft) externalDraft = null;
    draft.listeners.clear();

    let restoredFocusTarget: HTMLElement | null = null;
    if (outcome === 'discarded') {
      makeExternalDraftInert();
      draft.lease?.discard();
      refreshDiscardedChrome();
    } else if (draft.lease) {
      finishEditing();
      if (outcome === 'cancelled' && restore) {
        restoredFocusTarget = draft.lease.userCancel();
      } else {
        draft.lease.finalize();
      }
      suspended = false;
    } else {
      if (outcome === 'cancelled' && restore) {
        fxInput.value = draft.restoreValue;
        mutators.setPendingFormat(store, null);
      }
      finishEditing();
      if (focusHost && !detached) {
        host.focus();
        updateChrome();
      }
    }
    try {
      if (draft.lease && outcome === 'cancelled' && restore) {
        draft.hooks.onFinish(outcome, restoredFocusTarget);
      } else {
        draft.hooks.onFinish(outcome);
      }
    } catch (err) {
      console.warn('formulon-cell: external formula draft finish callback failed', err);
    }
  }

  const notifyExternalInput = (): void => {
    const draft = externalDraft;
    if (detached || draft?.status !== 'open') return;
    const view = currentExternalDraftView(draft);
    if (!view || draft.listeners.size === 0) return;
    const last = draft.lastNotifiedView;
    if (
      last &&
      last.raw === view.raw &&
      last.caret.start === view.caret.start &&
      last.caret.end === view.caret.end &&
      last.caret.direction === view.caret.direction
    )
      return;
    draft.lastNotifiedView = view;
    for (const listener of [...draft.listeners]) {
      if (draft.status !== 'open' || externalDraft !== draft || draft.lastNotifiedView !== view)
        break;
      listener(view);
    }
  };

  const syncFormulaBarInputState = (notify = true): void => {
    refreshActions();
    if (fxEditing) syncFxRefs();
    getAutocomplete().refresh();
    getArgHelper()?.refresh();
    if (notify) notifyExternalInput();
  };

  const discardStaleLeaseDraft = (draft: ExternalDraftState): void => {
    finishExternalDraft(draft, 'discarded', false, false);
  };

  const currentDraftValue = (draft: ExternalDraftState): string =>
    currentExternalDraftView(draft)?.raw ?? '';

  const syncExternalInput = (draft: ExternalDraftState, raw: string, caret?: number): void => {
    if (detached || externalDraft !== draft || draft.status !== 'open') return;
    if (externalDraftStale(draft)) {
      discardStaleLeaseDraft(draft);
      return;
    }
    const before = captureDraftSnapshot();
    draftHistory.endGroup();
    referenceGesture = null;
    pointNavigation = null;
    const nextCaret = Math.max(0, Math.min(caret ?? raw.length, raw.length));
    draftHistory.recordCommand(
      before,
      {
        value: raw,
        start: nextCaret,
        end: nextCaret,
        direction: 'none',
        mode: arrowMode,
      },
      'external-set-value',
    );
    fxInput.value = raw;
    fxInput.setSelectionRange(nextCaret, nextCaret);
    commitRejected = false;
    syncFormulaBarInputState(false);
    notifyExternalInput();
  };

  const notifyValidation = (outcome: {
    severity: 'stop' | 'warning' | 'information';
    title?: string;
    message: string;
  }): void => {
    if (!onValidation) return;
    try {
      onValidation(outcome);
    } catch (err) {
      console.warn('formulon-cell: formula-bar validation callback failed', err);
    }
  };

  const showCommitFailure = (message: string, title?: string): void => {
    commitRejected = true;
    // Focusing the bar under an adopted draft would blur and close the palette.
    if (!detached && !externalDraft?.lease) fxInput.focus();
    if (onValidation) notifyValidation({ severity: 'stop', title, message });
    else console.warn(`formulon-cell: formula-bar commit rejected: ${message}`);
  };

  const moveAfterCommit = (advance: 'down' | 'right' | 'up' | 'left' | 'none'): void => {
    if (advance === 'none') return;
    const controller = interactionControllerFor(store);
    if (controller?.policy?.selection === false) return;
    const hasTabStops =
      controller?.policy !== undefined ||
      navigationBoundsFor(store) !== undefined ||
      navigationSelectionBoundsFor(store) !== undefined;
    // Mac Tab honours explicit tab stops ahead of the selected rectangle.
    if (isMacPlatform() && (advance === 'right' || advance === 'left') && hasTabStops) {
      const target = nextTabStop(store, store.getState().selection.active, advance === 'left');
      if (target && isNavigationAddrAllowed(store, target)) mutators.setActive(store, target);
      return;
    }
    advanceAfterCommit(store, advance, isMacPlatform());
  };

  const commitSelection = (): void => {
    commitRejected = false;
    draftHistory.endGroup();
    referenceGesture = null;
    pointNavigation = null;
    const state = store.getState();
    if (commitDeferred) return;
    if (
      deferForAllowedRangeUnlock(store, selectionEditRanges(state), commitSelection, () => {
        commitDeferred = false;
      })
    ) {
      commitDeferred = true;
      return;
    }
    const anchor = state.selection.active;
    const batch = buildSelectionInputBatch(state, fxInput.value, anchor);
    if (!batch) {
      showCommitFailure(SELECTION_INPUT_LIMIT_MESSAGE);
      return;
    }
    const controller = interactionControllerFor(store);
    if (!controller) {
      // The direct path serves only hosts without a registered controller.
      const written = writeSelectionInput(wb(), store, state, fxInput.value, anchor);
      if (written.status === 'limitExceeded') {
        showCommitFailure(SELECTION_INPUT_LIMIT_MESSAGE);
        return;
      }
      if (written.status === 'rejected') {
        showCommitFailure(written.outcome.message, written.outcome.title);
        return;
      }
      mutators.replaceCells(store, wb().cells(store.getState().data.sheetIndex));
      finishEditing();
      host.focus();
      return;
    }

    let result: ReturnType<typeof controller.execute>;
    try {
      result = controller.execute({
        type: 'cellBatch',
        operation: batch.operation,
        origin: 'formulaBar',
        changes: batch.changes,
        denied: 'reject',
      });
    } catch (err) {
      console.warn('formulon-cell: formula-bar selection fill failed', err);
      showCommitFailure('The selected cells could not be filled.');
      return;
    }
    if (result.status === 'rejected') {
      showCommitFailure(
        result.rejected[0]?.reason ?? `The ${batch.operation} operation was rejected.`,
      );
      return;
    }

    const currentWb = wb();
    const pending = store.getState().ui.pendingFormat;
    if (pending && sameAddr(pending.addr, anchor)) mutators.setPendingFormat(store, null);
    mutators.replaceCells(store, currentWb.cells(store.getState().data.sheetIndex));
    finishEditing();
    host.focus();
  };

  const commitRawAt = (a: Addr, raw: string, draft?: ExternalDraftState): boolean => {
    const currentWb = wb();
    const result = commitCellInput({ store, wb: currentWb, addr: a, raw, origin: 'formulaBar' });
    if (result.status === 'rejected') {
      showCommitFailure(
        result.alert?.message ?? `The ${result.operation} operation was rejected.`,
        result.alert?.title,
      );
      return false;
    }
    if (result.status === 'failed') {
      console.warn('formulon-cell: formula-bar write failed', result.error);
      showCommitFailure('The formula-bar value could not be written.');
      return false;
    }
    if (result.notice) {
      if (onValidation) notifyValidation(result.notice);
      else {
        console.warn(
          `formulon-cell: validation ${result.notice.severity}: ${result.notice.message}`,
        );
      }
    }
    if (
      draft &&
      (detached ||
        externalDraft !== draft ||
        draft.cancelMode === 'discard' ||
        !externalDraftWorkbookCurrent(draft))
    ) {
      return true;
    }
    mutators.replaceCells(store, currentWb.cells(store.getState().data.sheetIndex));
    return true;
  };

  function commitExternalDraft(draft: ExternalDraftState): boolean {
    draftHistory.endGroup();
    referenceGesture = null;
    pointNavigation = null;
    if (detached || externalDraft !== draft || draft.status !== 'open') return false;
    if (externalDraftStale(draft)) {
      discardStaleLeaseDraft(draft);
      return false;
    }
    const raw = fxInput.value;
    draft.status = 'committing';
    let committed = false;
    try {
      committed = commitRawAt(draft.anchor, raw, draft);
    } catch (err) {
      console.warn('formulon-cell: external formula-bar commit failed', err);
    }
    // Context discard or invalidity wins over a write that completed while a
    // controller subscriber synchronously re-entered the session.
    const contextDiscarded =
      detached ||
      externalDraft !== draft ||
      draft.cancelMode === 'discard' ||
      externalDraftStale(draft);
    if (contextDiscarded) {
      discardStaleLeaseDraft(draft);
      return committed;
    }
    if (committed) {
      finishExternalDraft(draft, 'committed', false, false);
      return true;
    }
    if (draft.cancelRequested || detached) {
      finishExternalDraft(draft, 'cancelled', true, false);
      return false;
    }
    if (draft.status === 'committing') {
      draft.status = 'open';
    }
    return false;
  }

  function cancelExternalDraft(focusHost = true, expected?: ExternalDraftState): void {
    const draft = externalDraft;
    if (!draft || (draft !== expected && expected !== undefined)) return;
    if (draft.status === 'committing') {
      draft.cancelRequested = true;
      if (draft.cancelMode !== 'discard') draft.cancelMode = 'restore';
      return;
    }
    if (draft.status !== 'open') return;
    if (externalDraftStale(draft)) {
      discardStaleLeaseDraft(draft);
      return;
    }
    // Only a user cancellation may restore a leased owner.
    finishExternalDraft(draft, 'cancelled', true, focusHost && !detached);
  }

  function discardExternalDraft(expected?: ExternalDraftState): void {
    const draft = externalDraft;
    if (!draft || (draft !== expected && expected !== undefined)) return;
    if (draft.status === 'committing') {
      draft.cancelRequested = true;
      draft.cancelMode = 'discard';
      return;
    }
    if (draft.status !== 'open') return;
    finishExternalDraft(draft, 'discarded', false, false);
  }

  const commitFx = (advance: 'down' | 'right' | 'up' | 'left' | 'none'): boolean => {
    draftHistory.endGroup();
    referenceGesture = null;
    pointNavigation = null;
    if (externalDraft) {
      return externalDraft.status === 'open' ? commitExternalDraft(externalDraft) : false;
    }
    if (detached || suspended || !fxEditing) return false;
    commitRejected = false;
    const a = editingAnchor ?? store.getState().selection.active;
    // A password-guarded range under the cell is unlocked first; the commit then repeats.
    const cell = { sheet: a.sheet, r0: a.row, c0: a.col, r1: a.row, c1: a.col };
    // One deferred commit at a time; it keeps its Enter / Tab direction.
    if (commitDeferred) return false;
    if (
      deferForAllowedRangeUnlock(
        store,
        [cell],
        () => void commitFx(advance),
        () => {
          commitDeferred = false;
        },
      )
    ) {
      commitDeferred = true;
      return false;
    }
    if (!commitRawAt(a, fxInput.value)) return false;
    finishEditing();
    moveAfterCommit(advance);
    host.focus();
    return true;
  };

  const cancelFx = (): void => {
    referenceGesture = null;
    pointNavigation = null;
    arrowMode = 'edit';
    pendingHistoryInput = null;
    ignoredHistoryInput = null;
    draftHistory.clear();
    if (externalDraft) {
      cancelExternalDraft();
      return;
    }
    if (detached) return;
    const baseline = fxBaseline;
    invalidateLease();
    fxInput.value = baseline;
    fxEditing = false;
    commitRejected = false;
    mutators.setPendingFormat(store, null);
    clearFxRefs();
    closeFormulaHelpers();
    refreshActions();
    host.focus();
    updateChrome();
  };

  const discardFx = (): void => {
    referenceGesture = null;
    pointNavigation = null;
    arrowMode = 'edit';
    pendingHistoryInput = null;
    ignoredHistoryInput = null;
    draftHistory.clear();
    if (externalDraft) {
      discardExternalDraft();
      return;
    }
    if (!fxEditing && !suspended) return;
    invalidateLease();
    // System boundaries abandon a regular formula-bar edit in place. Do not
    // restore the old pending format or move focus while the host replaces the
    // context; clear only refs still owned by this edit.
    fxEditing = false;
    editingAnchor = null;
    commitRejected = false;
    fxBaseline = fxInput.value;
    suspended = false;
    pendingHistoryInput = null;
    ignoredHistoryInput = null;
    draftHistory.clear();
    clearFxRefs();
    refreshDiscardedChrome();
  };

  const acceptFx = (): void => {
    if (externalDraft) {
      if (externalDraft.status === 'open') commitExternalDraft(externalDraft);
      return;
    }
    if (detached) return;
    if (fxInput.value !== fxBaseline) commitFx('none');
  };

  const beginFormulaBarEdit = (baseline = fxInput.value): void => {
    if (detached || externalDraft || suspended) return;
    if (fxEditing) return;
    cancelBindingEditor();
    fxEditing = true;
    arrowMode = 'edit';
    pointNavigation = null;
    referenceGesture = null;
    fxBaseline = baseline;
    editingAnchor = { ...store.getState().selection.active };
    pendingHistoryInput = null;
    ignoredHistoryInput = null;
    draftHistory.reset(captureDraftSnapshot());
    refreshActions();
    syncFxRefs();
  };

  const onFxFocus = (): void => {
    // A data table's result cells show their TABLE() formula but cannot be edited.
    if (fxInput.readOnly) return;
    beginFormulaBarEdit();
  };

  const onFxBeforeInput = (event: InputEvent): void => {
    if (detached || (externalDraft && externalDraft.status !== 'open')) {
      pendingHistoryInput = null;
      ignoredHistoryInput = null;
      event.preventDefault();
      event.stopPropagation();
      return;
    }
    const currentDraft = externalDraft;
    if (currentDraft && externalDraftStale(currentDraft)) {
      discardStaleLeaseDraft(currentDraft);
      pendingHistoryInput = null;
      ignoredHistoryInput = null;
      event.preventDefault();
      event.stopPropagation();
      return;
    }
    if (!fxEditing && !externalDraft) return;
    const action =
      event.inputType === 'historyUndo'
        ? 'undo'
        : event.inputType === 'historyRedo'
          ? 'redo'
          : null;
    if (action) {
      if (composing) return;
      if (event.cancelable) {
        event.preventDefault();
        pendingHistoryInput = null;
        ignoredHistoryInput = { action, snapshot: applyDraftHistory(action) };
      } else {
        ignoredHistoryInput = null;
        pendingHistoryInput = action;
      }
      return;
    }
    pendingHistoryInput = null;
    ignoredHistoryInput = null;
    if (!composing) draftHistory.beforeInput(captureDraftSnapshot());
  };

  const onFxInput = (rawEvent?: Event): void => {
    if (detached) return;
    if (externalDraft && externalDraft.status !== 'open') {
      const snapshot = draftHistory.current();
      fxInput.value = snapshot.value;
      fxInput.setSelectionRange(snapshot.start, snapshot.end, snapshot.direction);
      pendingHistoryInput = null;
      ignoredHistoryInput = null;
      return;
    }
    const currentDraft = externalDraft;
    if (currentDraft && externalDraftStale(currentDraft)) {
      discardStaleLeaseDraft(currentDraft);
      pendingHistoryInput = null;
      ignoredHistoryInput = null;
      return;
    }
    const event = rawEvent as InputEvent | undefined;
    const historyAction =
      pendingHistoryInput ??
      (event?.inputType === 'historyUndo'
        ? 'undo'
        : event?.inputType === 'historyRedo'
          ? 'redo'
          : null);
    if (composing) {
      pendingHistoryInput = null;
      ignoredHistoryInput = null;
      referenceGesture = null;
      pointNavigation = null;
      draftHistory.input(captureDraftSnapshot(), { composing: true, inputType: event?.inputType });
      syncFormulaBarInputState();
      return;
    }
    const ignored = ignoredHistoryInput;
    if (historyAction && ignored?.action === historyAction) {
      ignoredHistoryInput = null;
      pendingHistoryInput = null;
      applyDraftSnapshot(ignored.snapshot);
      return;
    }
    ignoredHistoryInput = null;
    if (historyAction) {
      pendingHistoryInput = null;
      applyDraftHistory(historyAction);
      return;
    }
    referenceGesture = null;
    pointNavigation = null;
    // A system discard has already cached the current display in fxBaseline;
    // preserve it because the browser updates textarea.value before input.
    if (!fxEditing) beginFormulaBarEdit(fxBaseline);
    commitRejected = false;
    draftHistory.input(captureDraftSnapshot(), {
      composing,
      inputType: event?.inputType,
    });
    syncFormulaBarInputState();
  };

  const onFxCompositionStart = (): void => {
    if (detached || (externalDraft && externalDraft.status !== 'open')) {
      composing = false;
      return;
    }
    const currentDraft = externalDraft;
    if (currentDraft && externalDraftStale(currentDraft)) {
      composing = false;
      discardStaleLeaseDraft(currentDraft);
      return;
    }
    composing = true;
    pendingHistoryInput = null;
    ignoredHistoryInput = null;
    draftHistory.startComposition(captureDraftSnapshot());
    referenceGesture = null;
    pointNavigation = null;
    getAutocomplete().setComposing?.(true);
    getArgHelper()?.setComposing?.(true);
  };

  const onFxCompositionEnd = (): void => {
    composing = false;
    if (detached || (externalDraft && externalDraft.status !== 'open')) return;
    const currentDraft = externalDraft;
    if (currentDraft && externalDraftStale(currentDraft)) {
      discardStaleLeaseDraft(currentDraft);
      return;
    }
    draftHistory.endComposition(captureDraftSnapshot());
    getAutocomplete().setComposing?.(false);
    getArgHelper()?.setComposing?.(false);
    syncFormulaBarInputState();
  };

  const onFxKeyUp = (): void => {
    if (detached || (externalDraft && externalDraft.status !== 'open')) return;
    refreshHelpersAfterSelection();
    notifyExternalInput();
  };

  const onFxSelection = (): void => {
    if (suppressSelectionNotifications) return;
    refreshHelpersAfterSelection();
    notifyExternalInput();
  };

  const onFxClick = (): void => {
    arrowMode = 'edit';
    referenceGesture = null;
    pointNavigation = null;
    onFxSelection();
  };

  const refreshHelpersAfterSelection = (): void => {
    if (!composing && (fxEditing || externalDraft?.status === 'open')) {
      draftHistory.observeSelection(captureDraftSnapshot());
    }
    const gesture = referenceGesture;
    if (
      gesture &&
      (fxInput.value !== gesture.value ||
        fxInput.selectionStart !== gesture.caret ||
        fxInput.selectionEnd !== gesture.caret)
    ) {
      referenceGesture = null;
    }
    const point = pointNavigation;
    if (
      point &&
      (fxInput.value !== point.value ||
        fxInput.selectionStart !== point.caret ||
        fxInput.selectionEnd !== point.caret)
    ) {
      pointNavigation = null;
    }
    if (!fxEditing || detached || (externalDraft && externalDraft.status !== 'open')) return;
    getAutocomplete().refresh();
    getArgHelper()?.refresh();
  };

  const handlePointNavigation = (event: KeyboardEvent): boolean => {
    if (arrowMode !== 'point' || !isFormulaEdit() || composing || event.isComposing) return false;
    const arrow = classifyFormulaReferenceArrow(event, isMacPlatform());
    const anchor = editingAnchor ?? store.getState().selection.active;
    if (!arrow) return false;
    const historyBefore = captureDraftSnapshot();
    const result = navigateFormulaReference({
      state: store.getState(),
      text: fxInput.value,
      selectionStart: fxInput.selectionStart ?? fxInput.value.length,
      selectionEnd: fxInput.selectionEnd ?? fxInput.value.length,
      editedCell: anchor,
      direction: arrow.direction,
      extend: arrow.extend,
      current: pointNavigation,
      r1c1: store.getState().ui.r1c1 === true,
      jump: arrow.jump,
    });
    if (!result) return false;
    event.preventDefault();
    event.stopPropagation();
    referenceGesture = null;
    pointNavigation = result;
    fxInput.value = result.value;
    if (!draftHistory.hasActiveGroup) {
      draftHistory.beginGroup(historyBefore, 'reference');
    }
    draftHistory.updateGroup(
      {
        value: result.value,
        start: result.caret,
        end: result.caret,
        direction: 'none',
        mode: arrowMode,
      },
      'reference',
      historyBefore,
    );
    fxInput.setSelectionRange(result.caret, result.caret);
    syncFormulaBarInputState();
    // The programmatic edit must not reopen suggestions that swallow the next arrow.
    const autocomplete = getAutocomplete();
    if (autocomplete.dismiss) autocomplete.dismiss();
    else autocomplete.close();
    scrollActiveIntoView(store, result.tip);
    return true;
  };

  const onFxKey = (e: KeyboardEvent): void => {
    if (detached) return;
    if (composing || e.isComposing || e.key === 'Process') return;
    if (externalDraft && externalDraft.status !== 'open') {
      if (e.key === 'Escape') {
        e.preventDefault();
        e.stopPropagation();
        cancelExternalDraft();
        return;
      }
      e.preventDefault();
      e.stopPropagation();
      return;
    }
    if (handleDraftHistoryKey(e)) return;
    if (e.key === 'F2' && !e.ctrlKey && !e.metaKey && !e.altKey && !e.shiftKey && isFormulaEdit()) {
      e.preventDefault();
      e.stopPropagation();
      draftHistory.endGroup();
      arrowMode = arrowMode === 'point' ? 'edit' : 'point';
      draftHistory.observeSelection(captureDraftSnapshot());
      referenceGesture = null;
      pointNavigation = null;
      const autocomplete = getAutocomplete();
      if (autocomplete.dismiss) autocomplete.dismiss();
      else autocomplete.close();
      return;
    }
    const autocomplete = getAutocomplete();
    if (handleFormulaAutocompleteKey(e, autocomplete, { formulaEdit: isFormulaEdit() })) {
      e.stopPropagation();
      return;
    }
    if (handlePointNavigation(e)) return;
    if (externalDraft?.status === 'open') {
      if (e.key === 'Enter' && !e.altKey) {
        e.preventDefault();
        e.stopPropagation();
        commitExternalDraft(externalDraft);
        return;
      }
      if (e.key === 'Escape') {
        e.preventDefault();
        e.stopPropagation();
        cancelExternalDraft();
        return;
      }
      if (e.key === 'Tab') {
        e.stopPropagation();
        return;
      }
    }
    if (
      e.metaKey &&
      !e.ctrlKey &&
      !e.altKey &&
      !e.shiftKey &&
      e.key.toLowerCase() === 't' &&
      isMacPlatform()
    ) {
      e.preventDefault();
      e.stopPropagation();
      const caret = fxInput.selectionStart ?? fxInput.value.length;
      const before = captureDraftSnapshot();
      draftHistory.endGroup();
      const anchor = editingAnchor ?? externalDraft?.anchor ?? store.getState().selection.active;
      const r = rotateRefAt(fxInput.value, caret, {
        r1c1: store.getState().ui.r1c1 === true,
        base: { row: anchor.row, col: anchor.col },
      });
      if (r.text !== fxInput.value) {
        draftHistory.recordCommand(
          before,
          {
            value: r.text,
            start: r.caret,
            end: r.caret,
            direction: 'none',
            mode: arrowMode,
          },
          'f4',
        );
        referenceGesture = null;
        pointNavigation = null;
        fxInput.value = r.text;
        fxInput.setSelectionRange(r.caret, r.caret);
        syncFormulaBarInputState();
      }
      return;
    }
    if (e.key === 'Enter') {
      if (e.altKey) {
        e.stopPropagation();
        return;
      }
      const macFill = isMacPlatform() && e.metaKey && !e.ctrlKey && !e.shiftKey;
      if (macFill || (e.ctrlKey && !e.metaKey && !e.shiftKey)) {
        e.preventDefault();
        e.stopPropagation();
        commitSelection();
        return;
      }
      if (e.metaKey) {
        e.stopPropagation();
        return;
      }
      if (e.shiftKey) {
        if (!isMacPlatform()) {
          // On the default platform, Shift+Enter remains a literal newline in
          // the formula-bar editor, matching the existing textarea behavior.
          e.stopPropagation();
          return;
        }
        e.preventDefault();
        e.stopPropagation();
        commitFx('up');
        return;
      }
      e.preventDefault();
      e.stopPropagation();
      commitFx('down');
    } else if (e.key === 'Tab') {
      e.stopPropagation();
      if (isMacPlatform()) {
        e.preventDefault();
        commitFx(e.shiftKey ? 'left' : 'right');
        return;
      }
      const controller = interactionControllerFor(store);
      const restricted = controller?.policy !== undefined;
      const selectionDisabled = controller?.policy?.selection === false;
      if (selectionDisabled) {
        commitFx('none');
        return;
      }
      const active = store.getState().selection.active;
      const next = restricted ? nextTabStop(store, active, e.shiftKey) : null;
      if (restricted && next === null) {
        commitFx('none');
        return;
      }
      e.preventDefault();
      commitFx(restricted ? 'none' : e.shiftKey ? 'none' : 'right');
      if (restricted && next && !fxEditing) mutators.setActive(store, next);
    } else if (e.key === 'Escape') {
      e.preventDefault();
      e.stopPropagation();
      cancelFx();
    } else if (e.key === 'F9' && draftActive()) {
      const start = fxInput.selectionStart ?? 0;
      const end = fxInput.selectionEnd ?? start;
      const currentWb = wb();
      const sheetByName = (name: string): number => {
        for (let index = 0; index < currentWb.sheetCount; index += 1) {
          if (currentWb.sheetName(index) === name) return index;
        }
        return -1;
      };
      const active = editingAnchor ?? externalDraft?.anchor ?? store.getState().selection.active;
      const result = replaceFormulaSelectionWithF9Preview(
        fxInput.value,
        start,
        end,
        store.getState().data.sheetIndex,
        store.getState().data.cells,
        sheetByName,
        (formula) =>
          currentWb.evaluateFormulaText(active, currentWb.canonicalizeLocalizedFormula(formula)),
        true,
        (formula) =>
          currentWb.evaluateFormulaArray(active, currentWb.canonicalizeLocalizedFormula(formula)),
        {
          localizeLiteral: (literal) => {
            const localized = currentWb.localizeFormulaText(`=${literal}`);
            return localized.startsWith('=') ? localized.slice(1) : literal;
          },
        },
      );
      if (result) {
        e.preventDefault();
        e.stopPropagation();
        const before = captureDraftSnapshot();
        draftHistory.endGroup();
        draftHistory.recordCommand(
          before,
          {
            value: result.text,
            start: result.start,
            end: result.end,
            direction: 'none',
            mode: arrowMode,
          },
          'f9',
        );
        referenceGesture = null;
        pointNavigation = null;
        fxInput.value = result.text;
        fxInput.setSelectionRange(result.start, result.end);
        syncFormulaBarInputState();
      }
    } else if (e.key === 'F4' && draftActive()) {
      e.preventDefault();
      e.stopPropagation();
      const caret = fxInput.selectionStart ?? fxInput.value.length;
      const before = captureDraftSnapshot();
      draftHistory.endGroup();
      const anchor = editingAnchor ?? externalDraft?.anchor ?? store.getState().selection.active;
      const r = rotateRefAt(fxInput.value, caret, {
        r1c1: store.getState().ui.r1c1 === true,
        base: { row: anchor.row, col: anchor.col },
      });
      if (r.text !== fxInput.value) {
        draftHistory.recordCommand(
          before,
          {
            value: r.text,
            start: r.caret,
            end: r.caret,
            direction: 'none',
            mode: arrowMode,
          },
          'f4',
        );
        referenceGesture = null;
        pointNavigation = null;
        fxInput.value = r.text;
        fxInput.setSelectionRange(r.caret, r.caret);
        syncFormulaBarInputState();
      }
    }
  };

  const onFxBlur = (): void => {
    if (detached || externalDraft || suspended) return;
    if (!fxEditing) {
      clearFxRefs();
      closeFormulaHelpers();
      return;
    }
    if (commitRejected || commitDeferred) return;
    if (fxInput.value !== fxBaseline) commitFx('none');
    else {
      pointNavigation = null;
      arrowMode = 'edit';
      fxEditing = false;
      pendingHistoryInput = null;
      ignoredHistoryInput = null;
      draftHistory.clear();
      refreshActions();
      clearFxRefs();
      closeFormulaHelpers();
    }
  };

  const insertRefAtCaret = (ref: string, options?: { phase: 'start' | 'update' }): void => {
    if (detached || (externalDraft && externalDraft.status !== 'open')) return;
    const historyBefore = captureDraftSnapshot();
    if (options?.phase === 'start') {
      draftHistory.beginGroup(historyBefore, 'reference');
    } else if (options?.phase !== 'update') {
      draftHistory.endGroup();
      draftHistory.beginCommand(historyBefore);
    }
    if (options?.phase !== 'update') arrowMode = 'point';
    pointNavigation = null;
    const selectionStart = fxInput.selectionStart ?? fxInput.value.length;
    const selectionEnd = fxInput.selectionEnd ?? selectionStart;
    let gesture: FormulaReferenceGesture | null = null;
    if (options?.phase === 'start') {
      gesture = createFormulaReferenceGesture(fxInput.value, selectionStart, selectionEnd, {
        r1c1: store.getState().ui.r1c1 === true,
      });
    } else if (options?.phase === 'update') {
      const current = referenceGesture;
      if (
        !current ||
        current.value !== fxInput.value ||
        current.caret !== selectionStart ||
        current.caret !== selectionEnd
      ) {
        referenceGesture = null;
        return;
      }
      gesture = current.gesture;
    } else {
      referenceGesture = null;
    }
    const r1c1 = store.getState().ui.r1c1 === true;
    const picked =
      gesture ??
      createFormulaReferenceGesture(fxInput.value, selectionStart, selectionEnd, { r1c1 });
    const result = picked.insert(ref);
    referenceGesture =
      options?.phase && gesture ? { gesture, value: result.text, caret: result.caret } : null;
    const after: FormulaDraftSnapshot = {
      value: result.text,
      start: result.caret,
      end: result.caret,
      direction: 'none',
      mode: arrowMode,
    };
    if (options?.phase) {
      draftHistory.updateGroup(after, 'reference', historyBefore);
    } else {
      draftHistory.finishCommand(after);
    }
    // The next arrow continues from the picked cell or range.
    pointNavigation = pickedReferenceNavigation({
      gesture: picked,
      inserted: result,
      ref,
      editedCell: editingAnchor ?? externalDraft?.anchor ?? store.getState().selection.active,
      r1c1,
    });
    fxInput.value = result.text;
    fxInput.setSelectionRange(result.caret, result.caret);
    fxInput.focus();
    syncFormulaBarInputState();
  };

  const beginExternalDraft = (
    anchor: Addr,
    seed: string,
    hooks: ExternalFormulaDraftHooks,
    options?: { lease?: FormulaEditLease },
  ): ExternalFormulaDraftHandle | null => {
    const lease = options?.lease ?? null;
    if (detached || externalDraft || store.getState().ui.editor.kind !== 'idle') return null;
    if (lease) {
      const snapshot = lease.snapshot;
      if (!lease.valid() || !sameAddr(snapshot.anchor, anchor)) {
        return null;
      }
      if (wb() !== snapshot.workbook) return null;
    } else {
      if (fxEditing) return null;
      cancelBindingEditor();
    }
    const snapshot = lease?.snapshot;
    const draftSeed = snapshot?.raw ?? seed;
    const draft: ExternalDraftState = {
      anchor: { ...anchor },
      workbook: wb(),
      restoreValue: fxInput.value,
      hooks,
      listeners: new Set(),
      lastNotifiedView: null,
      status: 'open',
      cancelRequested: false,
      cancelMode: null,
      lease,
    };
    externalDraft = draft;
    referenceGesture = null;
    pointNavigation = null;
    arrowMode = 'edit';
    fxEditing = true;
    suspended = lease !== null;
    commitRejected = false;
    editingAnchor = { ...anchor };
    fxBaseline = snapshot?.baseline ?? draftSeed;
    fxInput.value = draftSeed;
    if (snapshot) {
      fxInput.setSelectionRange(snapshot.caret.start, snapshot.caret.end, snapshot.caret.direction);
    } else {
      fxInput.setSelectionRange(draftSeed.length, draftSeed.length);
    }
    pendingHistoryInput = null;
    ignoredHistoryInput = null;
    draftHistory.reset(captureDraftSnapshot());
    refreshActions();
    syncFxRefs();
    getAutocomplete().refresh();
    getArgHelper()?.refresh();

    const exposedAnchor = { ...draft.anchor };
    return {
      anchor: exposedAnchor,
      value: () => currentDraftValue(draft),
      snapshot: () => currentExternalDraftView(draft),
      setValue: (raw, caret) => syncExternalInput(draft, raw, caret),
      commit: () => commitExternalDraft(draft),
      cancel: () => cancelExternalDraft(true, draft),
      discard: () => discardExternalDraft(draft),
      subscribe: (listener) => {
        if (detached || externalDraft !== draft || draft.status !== 'open') return () => {};
        if (externalDraftStale(draft)) {
          discardStaleLeaseDraft(draft);
          return () => {};
        }
        draft.listeners.add(listener);
        return () => draft.listeners.delete(listener);
      },
    };
  };

  const keepFxFocus = (e: MouseEvent): void => e.preventDefault();

  fxInput.addEventListener('focus', onFxFocus);
  fxInput.addEventListener('beforeinput', onFxBeforeInput);
  fxInput.addEventListener('input', onFxInput);
  fxInput.addEventListener('compositionstart', onFxCompositionStart);
  fxInput.addEventListener('compositionend', onFxCompositionEnd);
  fxInput.addEventListener('keyup', onFxKeyUp);
  fxInput.addEventListener('select', onFxSelection);
  fxInput.addEventListener('mouseup', onFxSelection);
  fxInput.addEventListener('click', onFxClick);
  fxInput.addEventListener('keydown', onFxKey);
  fxInput.addEventListener('blur', onFxBlur);
  fxCancel.addEventListener('mousedown', keepFxFocus);
  fxAccept.addEventListener('mousedown', keepFxFocus);
  fxCancel.addEventListener('click', cancelFx);
  fxAccept.addEventListener('click', acceptFx);
  refreshActions();

  return {
    acceptFx,
    beginExternalDraft,
    cancelFx,
    discardFx,
    commitFx,
    detach(): void {
      if (detached) return;
      detached = true;
      fxInput.removeEventListener('focus', onFxFocus);
      fxInput.removeEventListener('beforeinput', onFxBeforeInput);
      fxInput.removeEventListener('input', onFxInput);
      fxInput.removeEventListener('compositionstart', onFxCompositionStart);
      fxInput.removeEventListener('compositionend', onFxCompositionEnd);
      fxInput.removeEventListener('keyup', onFxKeyUp);
      fxInput.removeEventListener('select', onFxSelection);
      fxInput.removeEventListener('mouseup', onFxSelection);
      fxInput.removeEventListener('click', onFxClick);
      fxInput.removeEventListener('keydown', onFxKey);
      fxInput.removeEventListener('blur', onFxBlur);
      fxCancel.removeEventListener('mousedown', keepFxFocus);
      fxAccept.removeEventListener('mousedown', keepFxFocus);
      fxCancel.removeEventListener('click', cancelFx);
      fxAccept.removeEventListener('click', acceptFx);
      if (externalDraft) discardExternalDraft();
      else {
        referenceGesture = null;
        pendingHistoryInput = null;
        ignoredHistoryInput = null;
        draftHistory.clear();
        invalidateLease();
      }
    },
    insertRefAtCaret,
    beforeAutocompleteInsert(): void {
      if (fxEditing || externalDraft?.status === 'open') {
        suppressSelectionNotifications = true;
        draftHistory.beginCommand(captureDraftSnapshot());
      }
    },
    notifyAutocompleteInsert(outcome?: AutocompleteInsertOutcome): void {
      suppressSelectionNotifications = false;
      pointNavigation = null;
      referenceGesture = null;
      if (outcome?.kind === 'function') arrowMode = 'point';
      draftHistory.finishCommand(captureDraftSnapshot());
      syncFormulaBarInputState(false);
    },
    isEditing: () => fxEditing,
    isFormulaEdit,
    refreshActions,
    syncFxRefs,
    suspendForFormulaPalette,
  };
}
