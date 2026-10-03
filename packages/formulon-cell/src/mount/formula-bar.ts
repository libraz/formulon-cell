import { commitCellInput } from '../commands/cell-input-commit.js';
import { interactionControllerFor } from '../commands/interaction-controller.js';
import { extractRefs, rotateRefAt } from '../commands/refs.js';
import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import {
  createFormulaEditLease,
  type FormulaEditLease,
  type FormulaEditLeaseContext,
  type FormulaEditLeaseSnapshot,
} from '../interact/formula-edit-lease.js';
import {
  isNavigationAddrAllowed,
  navigationBoundsFor,
  navigationSelectionBoundsFor,
  nextTabStop,
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
import { projectDisabledState } from '../toolbar/menu-a11y.js';

interface FormulaBarAutocomplete {
  isOpen(): boolean;
  move(n: number): void;
  acceptHighlighted(): boolean;
  close(): void;
  refresh(): void;
}

interface FormulaArgHelper {
  refresh(): void;
  close?(): void;
}

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

export interface ExternalFormulaDraftHooks {
  onFinish(outcome: 'committed' | 'cancelled', restoredFocusTarget?: HTMLElement | null): void;
}

export interface ExternalFormulaDraftHandle {
  readonly anchor: Addr;
  value(): string;
  setValue(raw: string, caret?: number): void;
  commit(): boolean;
  cancel(): void;
  discard(): void;
  subscribe(fn: (raw: string) => void): () => void;
}

type ExternalDraftStatus = 'open' | 'committing' | 'committed' | 'cancelled';

interface ExternalDraftState {
  anchor: Addr;
  workbook: WorkbookHandle;
  restoreValue: string;
  hooks: ExternalFormulaDraftHooks;
  listeners: Set<(raw: string) => void>;
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
  commitFx(advance: 'down' | 'right' | 'up' | 'left' | 'none'): boolean;
  detach(): void;
  insertRefAtCaret(ref: string): void;
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
  let leaseGeneration = 0;
  let currentLease: FormulaEditLease | null = null;
  let suspended = false;
  let commitRejected = false;
  let externalDraft: ExternalDraftState | null = null;
  let detached = false;

  const isMacPlatform = (): boolean =>
    (host.closest<HTMLElement>('.fc-host') ?? host).dataset.fcPlatform === 'mac';

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

  const closeFormulaHelpers = (): void => {
    getAutocomplete().close();
    getArgHelper()?.close?.();
  };

  const finishEditing = (): void => {
    fxEditing = false;
    editingAnchor = null;
    commitRejected = false;
    fxBaseline = fxInput.value;
    refreshActions();
    clearFxRefs();
    closeFormulaHelpers();
  };

  const restoreFromLease = (snapshot: Readonly<FormulaEditLeaseSnapshot>): HTMLElement | null => {
    if (detached || wb() !== snapshot.workbook) return null;
    const state = store.getState();
    if (state.data.sheetIndex !== snapshot.anchor.sheet || state.ui.r1c1 !== snapshot.r1c1) {
      return null;
    }
    store.setState((current) => ({
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
    fxInput.value = snapshot.raw;
    fxInput.setSelectionRange(
      Math.max(0, Math.min(snapshot.caret.start, snapshot.raw.length)),
      Math.max(0, Math.min(snapshot.caret.end, snapshot.raw.length)),
      snapshot.caret.direction,
    );
    fxBaseline = snapshot.baseline;
    editingAnchor = { ...snapshot.anchor };
    fxEditing = true;
    suspended = false;
    commitRejected = false;
    refreshActions();
    mutators.setEditorRefs(store, [...snapshot.editorRefs]);
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
      restore: (captured) => restoreFromLease(captured),
      release: () => {
        if (currentLease === lease) currentLease = null;
        if (suspended) {
          // Go inert without finishEditing: a newer owner may already hold the store.
          suspended = false;
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

  const notifyExternalInput = (): void => {
    const draft = externalDraft;
    if (detached || draft?.status !== 'open') return;
    for (const listener of [...draft.listeners]) listener(fxInput.value);
  };

  const discardStaleLeaseDraft = (draft: ExternalDraftState): void => {
    draft.status = 'cancelled';
    draft.lease?.discard();
    // Go inert without finishEditing so a later Enter/blur cannot write stale text.
    fxEditing = false;
    editingAnchor = null;
    commitRejected = false;
    fxBaseline = fxInput.value;
    suspended = false;
    if (externalDraft === draft) externalDraft = null;
  };

  const currentDraftValue = (draft: ExternalDraftState): string => {
    if (detached || externalDraft !== draft || draft.status !== 'open') return '';
    if (draft.lease && (!draft.lease.valid() || wb() !== draft.workbook)) {
      discardStaleLeaseDraft(draft);
      return '';
    }
    return fxInput.value;
  };

  const syncExternalInput = (draft: ExternalDraftState, raw: string, caret?: number): void => {
    if (detached || externalDraft !== draft || draft.status !== 'open') return;
    if (draft.lease && (!draft.lease.valid() || wb() !== draft.workbook)) {
      discardStaleLeaseDraft(draft);
      return;
    }
    if (wb() !== draft.workbook) return;
    fxInput.value = raw;
    const nextCaret = Math.max(0, Math.min(caret ?? raw.length, raw.length));
    fxInput.setSelectionRange(nextCaret, nextCaret);
    commitRejected = false;
    refreshActions();
    syncFxRefs();
    getAutocomplete().refresh();
    getArgHelper()?.refresh();
    notifyExternalInput();
  };

  function finishExternalDraft(
    draft: ExternalDraftState,
    outcome: 'committed' | 'cancelled',
    restore: boolean,
    focusHost: boolean,
  ): void {
    if (draft.status !== 'open' && draft.status !== 'committing') return;
    draft.status = outcome;
    let restoredFocusTarget: HTMLElement | null = null;
    if (draft.lease) {
      // An invalid lease discards without touching the current store.
      if (restore) {
        if (!draft.lease.valid()) {
          draft.lease.discard();
        } else {
          finishEditing();
          restoredFocusTarget = draft.lease.userCancel();
        }
      } else {
        finishEditing();
        draft.lease.finalize();
      }
      suspended = false;
    } else {
      if (restore) {
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
      if (draft.lease && restore) draft.hooks.onFinish(outcome, restoredFocusTarget);
      else draft.hooks.onFinish(outcome);
    } catch (err) {
      console.warn('formulon-cell: external formula draft finish callback failed', err);
    } finally {
      if (externalDraft === draft) externalDraft = null;
    }
  }

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
    const state = store.getState();
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

  const commitRawAt = (a: Addr, raw: string): boolean => {
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
    mutators.replaceCells(store, currentWb.cells(store.getState().data.sheetIndex));
    return true;
  };

  function commitExternalDraft(draft: ExternalDraftState): boolean {
    if (detached || externalDraft !== draft || draft.status !== 'open') return false;
    if (draft.lease && (!draft.lease.valid() || wb() !== draft.workbook)) {
      discardStaleLeaseDraft(draft);
      return false;
    }
    if (wb() !== draft.workbook) {
      finishExternalDraft(draft, 'cancelled', true, !detached);
      return false;
    }
    const raw = fxInput.value;
    draft.status = 'committing';
    let committed = false;
    try {
      committed = commitRawAt(draft.anchor, raw);
    } catch (err) {
      console.warn('formulon-cell: external formula-bar commit failed', err);
    }
    // Subscribers may have replaced the context during execute(); leave a newer owner alone.
    if (draft.lease && (!draft.lease.valid() || detached || draft.cancelMode === 'discard')) {
      discardStaleLeaseDraft(draft);
      return committed;
    }
    if (committed) {
      finishExternalDraft(draft, 'committed', false, false);
      return true;
    }
    if (draft.cancelRequested || detached) {
      finishExternalDraft(draft, 'cancelled', draft.cancelMode !== 'discard', false);
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
    if (draft.lease && (!draft.lease.valid() || wb() !== draft.workbook)) {
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
    if (draft.lease) {
      if (!draft.lease.valid() || wb() !== draft.workbook) {
        discardStaleLeaseDraft(draft);
        return;
      }
      draft.status = 'cancelled';
      finishEditing();
      draft.lease.discard();
      suspended = false;
      if (externalDraft === draft) externalDraft = null;
      return;
    }
    finishExternalDraft(draft, 'cancelled', true, false);
  }

  const commitFx = (advance: 'down' | 'right' | 'up' | 'left' | 'none'): boolean => {
    if (externalDraft) {
      return externalDraft.status === 'open' ? commitExternalDraft(externalDraft) : false;
    }
    if (detached || suspended || !fxEditing) return false;
    commitRejected = false;
    const a = editingAnchor ?? store.getState().selection.active;
    if (!commitRawAt(a, fxInput.value)) return false;
    finishEditing();
    moveAfterCommit(advance);
    host.focus();
    return true;
  };

  const cancelFx = (): void => {
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

  const acceptFx = (): void => {
    if (externalDraft) {
      if (externalDraft.status === 'open') commitExternalDraft(externalDraft);
      return;
    }
    if (detached) return;
    if (fxInput.value !== fxBaseline) commitFx('none');
  };

  const onFxFocus = (): void => {
    if (detached || externalDraft || suspended) return;
    if (fxEditing) return;
    cancelBindingEditor();
    fxEditing = true;
    fxBaseline = fxInput.value;
    editingAnchor = { ...store.getState().selection.active };
    refreshActions();
    syncFxRefs();
  };

  const onFxInput = (): void => {
    if (detached || (externalDraft && externalDraft.status !== 'open')) return;
    commitRejected = false;
    refreshActions();
    if (fxEditing) syncFxRefs();
    getAutocomplete().refresh();
    getArgHelper()?.refresh();
    notifyExternalInput();
  };

  const onFxCompositionStart = (): void => {
    composing = true;
  };

  const onFxCompositionEnd = (): void => {
    composing = false;
    onFxInput();
  };

  const onFxKeyUp = (): void => {
    if (detached || (externalDraft && externalDraft.status !== 'open')) return;
    if (fxEditing) getArgHelper()?.refresh();
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
    const autocomplete = getAutocomplete();
    if (autocomplete.isOpen()) {
      if (e.key === 'ArrowDown') {
        e.preventDefault();
        e.stopPropagation();
        autocomplete.move(1);
        return;
      }
      if (e.key === 'ArrowUp') {
        e.preventDefault();
        e.stopPropagation();
        autocomplete.move(-1);
        return;
      }
      if ((e.key === 'Enter' || e.key === 'Tab') && autocomplete.acceptHighlighted()) {
        e.preventDefault();
        e.stopPropagation();
        notifyExternalInput();
        return;
      }
      if (e.key === 'Escape') {
        e.preventDefault();
        e.stopPropagation();
        autocomplete.close();
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
      const r = rotateRefAt(fxInput.value, caret);
      if (r.text !== fxInput.value) {
        fxInput.value = r.text;
        fxInput.setSelectionRange(r.caret, r.caret);
        syncFxRefs();
        getAutocomplete().refresh();
        getArgHelper()?.refresh();
        notifyExternalInput();
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
    } else if (e.key === 'F4') {
      e.preventDefault();
      e.stopPropagation();
      const caret = fxInput.selectionStart ?? fxInput.value.length;
      const r = rotateRefAt(fxInput.value, caret);
      if (r.text !== fxInput.value) {
        fxInput.value = r.text;
        fxInput.setSelectionRange(r.caret, r.caret);
        syncFxRefs();
        notifyExternalInput();
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
    if (commitRejected) return;
    if (fxInput.value !== fxBaseline) commitFx('none');
    else {
      fxEditing = false;
      refreshActions();
      clearFxRefs();
      closeFormulaHelpers();
    }
  };

  const insertRefAtCaret = (ref: string): void => {
    if (detached || (externalDraft && externalDraft.status !== 'open')) return;
    const start = fxInput.selectionStart ?? fxInput.value.length;
    const end = fxInput.selectionEnd ?? start;
    const before = fxInput.value.slice(0, start);
    const after = fxInput.value.slice(end);
    // Replace any trailing partial reference token. This lets repeated pointer
    // updates during a drag extend or replace the in-progress reference.
    const stripped = before.replace(
      /(?:\$?[A-Za-z]+\$?\d+:\$?[A-Za-z]+\$?\d*|\$?[A-Za-z]+\$?\d*|:)$/,
      '',
    );
    fxInput.value = stripped + ref + after;
    const caret = stripped.length + ref.length;
    fxInput.setSelectionRange(caret, caret);
    fxInput.focus();
    refreshActions();
    syncFxRefs();
    getAutocomplete().refresh();
    getArgHelper()?.refresh();
    notifyExternalInput();
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
      status: 'open',
      cancelRequested: false,
      cancelMode: null,
      lease,
    };
    externalDraft = draft;
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
    refreshActions();
    syncFxRefs();
    getAutocomplete().refresh();
    getArgHelper()?.refresh();

    const exposedAnchor = { ...draft.anchor };
    return {
      anchor: exposedAnchor,
      value: () => currentDraftValue(draft),
      setValue: (raw, caret) => syncExternalInput(draft, raw, caret),
      commit: () => commitExternalDraft(draft),
      cancel: () => cancelExternalDraft(true, draft),
      discard: () => discardExternalDraft(draft),
      subscribe: (listener) => {
        if (detached || externalDraft !== draft || draft.status !== 'open') return () => {};
        if (draft.lease && (!draft.lease.valid() || wb() !== draft.workbook)) {
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
  fxInput.addEventListener('input', onFxInput);
  fxInput.addEventListener('compositionstart', onFxCompositionStart);
  fxInput.addEventListener('compositionend', onFxCompositionEnd);
  fxInput.addEventListener('keyup', onFxKeyUp);
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
    commitFx,
    detach(): void {
      if (detached) return;
      detached = true;
      fxInput.removeEventListener('focus', onFxFocus);
      fxInput.removeEventListener('input', onFxInput);
      fxInput.removeEventListener('compositionstart', onFxCompositionStart);
      fxInput.removeEventListener('compositionend', onFxCompositionEnd);
      fxInput.removeEventListener('keyup', onFxKeyUp);
      fxInput.removeEventListener('keydown', onFxKey);
      fxInput.removeEventListener('blur', onFxBlur);
      fxCancel.removeEventListener('mousedown', keepFxFocus);
      fxAccept.removeEventListener('mousedown', keepFxFocus);
      fxCancel.removeEventListener('click', cancelFx);
      fxAccept.removeEventListener('click', acceptFx);
      if (externalDraft) discardExternalDraft();
      else invalidateLease();
    },
    insertRefAtCaret,
    isEditing: () => fxEditing,
    isFormulaEdit: () => fxEditing && fxInput.value.trimStart().startsWith('='),
    refreshActions,
    syncFxRefs,
    suspendForFormulaPalette,
  };
}
