import {
  buildFunctionCatalog,
  type CatalogFunctionCategory,
  type FunctionCatalogEntry,
  type FunctionCatalogReader,
  type FunctionCatalogSnapshot,
  type FunctionCategory,
  functionSyntax,
  isFunctionUnavailableForInsertion,
} from '../commands/function-categories.js';
import {
  getRecentFunctions,
  recordRecentFunction,
  subscribeRecentFunctions,
} from '../commands/function-history.js';
import type { Addr } from '../engine/types.js';
import { formatCell, fromEngineValue } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import type { ExternalFormulaDraftHandle, FormulaBarController } from '../mount/formula-bar.js';
import type { SpreadsheetStore } from '../store/store.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import type { FormulaEditLease, FormulaEditLeaseContext } from './formula-edit-lease.js';
import {
  CATEGORY_LABEL_KEY,
  catalogLocaleOrdinal,
  functionDescription,
} from './function-catalog-text.js';
import type { FxDialogHandle, FxDialogOpenOptions } from './fx-dialog.js';
import {
  assembledFormula,
  canonicalName,
  formulaWithArgumentCount,
  type ProjectedFormulaCall,
  parseOuterCall,
  projectFormulaCallAtCaret,
  replaceProjectedFormulaCall,
} from './mac-formula-call.js';
import {
  argumentFieldCount,
  type FunctionArgumentHelp,
  type FunctionArgumentHelpProvider,
  insertableEntry,
  paletteStrings,
  pickerSections,
  resolveArgumentHelp,
  sameAvailability,
} from './mac-formula-palette-catalog.js';
import {
  type ArgumentsRefs,
  type ArgumentsViewState,
  createMacFormulaPaletteView,
  type PickerRefs,
  type PickerViewState,
} from './mac-formula-palette-view.js';
import type { RangeInsertTarget } from './range-insert.js';

export type {
  FunctionArgumentHelp,
  FunctionArgumentHelpProvider,
} from './mac-formula-palette-catalog.js';

export interface MacFormulaPaletteDeps {
  host: HTMLElement;
  dock: HTMLElement;
  store: SpreadsheetStore;
  getWb: () => WorkbookHandle;
  getLocale: () => string;
  getStrings: () => Strings;
  getAnchor: () => Addr;
  getInitialArguments?: (name: string) => readonly string[] | null;
  beginDraft: FormulaBarController['beginExternalDraft'];
  /** Suspend an in-progress cell or formula-bar edit so the palette can take it over. */
  suspendActiveEdit?: (context: FormulaEditLeaseContext) => FormulaEditLease | null;
  projectMirror: (anchor: Addr, raw: string | null) => void;
  getFunctionArgumentHelp?: FunctionArgumentHelpProvider;
}

export interface MacFormulaPaletteHandle extends FxDialogHandle {
  discard(): void;
  isOpen(): boolean;
  rangeInsertTarget(): RangeInsertTarget | null;
  setStrings(next: Strings): void;
  setReturnFocusTarget(target: HTMLElement | null): void;
}

type PaletteMode = 'closed' | 'picker' | 'arguments-editing' | 'arguments-committed';

interface CommitSnapshot {
  name: string;
  args: string[];
  raw: string;
  result: string;
  caret: number;
}

interface DraftBinding {
  handle: ExternalFormulaDraftHandle;
  unsubscribe: () => void;
  leased: boolean;
}

interface DraftView {
  raw: string;
  caret: {
    start: number;
    end: number;
    direction: HTMLTextAreaElement['selectionDirection'];
  };
}
const connectedElement = (value: Element | null): HTMLElement | null =>
  value instanceof HTMLElement && value.isConnected ? value : null;

export function attachMacFormulaPalette(deps: MacFormulaPaletteDeps): MacFormulaPaletteHandle {
  const root = document.createElement('aside');
  root.className = 'fc-mac-formula-palette';
  root.setAttribute('role', 'complementary');
  root.style.inlineSize = '300px';
  root.style.flexBasis = '300px';
  root.hidden = true;
  deps.dock.appendChild(root);
  deps.dock.hidden = true;

  let strings = deps.getStrings();
  let labels = paletteStrings(strings);
  let mode: PaletteMode = 'closed';
  let detached = false;
  let opener: HTMLElement | null = null;
  let restoredFocusTarget: HTMLElement | null = null;
  let returnFocusTarget: HTMLElement | null = null;
  let anchor: Addr | null = null;
  let sessionWb: WorkbookHandle | null = null;
  let catalog = buildFunctionCatalog(null, catalogLocaleOrdinal(deps.getLocale()));
  let pickerCategory: FunctionCategory = 'all';
  let pickerSelectionName: string | null = null;
  let pickerSelectionEntry: FunctionCatalogEntry | null = null;
  let selectedName: string | null = null;
  let selectedEntry: FunctionCatalogEntry | null = null;
  let args: string[] = [];
  let synchronized = true;
  let explicitArgumentCount: number | null = null;
  let preserveExplicitArgumentCount = false;
  let writingDraftRaw = false;
  let focusedArgument = 0;
  let currentRaw = '';
  let currentCaret = 0;
  let callProjection: ProjectedFormulaCall | null = null;
  let guardedUnavailable = false;
  let paletteWriteFocusIndex: number | null = null;
  let searchQuery = '';
  let guardMessage = '';
  let postCommit: CommitSnapshot | null = null;
  let activeDraft: DraftBinding | null = null;
  let pickerRefs: PickerRefs | null = null;
  let argumentsRefs: ArgumentsRefs | null = null;
  let recentUnsubscribe: (() => void) | null = null;
  let closing = false;
  let guarding = false;
  let committing = false;
  let staleProjectionWrite = false;

  const currentDraftBinding = (): DraftBinding | null => activeDraft;

  const view = createMacFormulaPaletteView(root, {
    labels: () => labels,
    pickerState: () => pickerViewState(),
    argumentsState: () => argumentsViewState(),
    close: () => api.close(),
    showAll: () => returnToPicker(),
    search: (query) => {
      searchQuery = query;
      updatePickerList();
    },
    pick: (name) => pickEntry(name),
    insert: () =>
      insertSelected(pickerSelectionName ?? selectedName, pickerSelectionEntry ?? selectedEntry),
    argumentFocus: (index) => {
      focusedArgument = index;
    },
    argumentInput: (index, value, sourceField) => onArgumentInput(index, value, sourceField),
    addArgument: () => addArgument(),
    done: () => commitSelected(),
  });

  const leaseContext: FormulaEditLeaseContext = {
    getLocale: () => deps.getLocale(),
    contextCurrent: () => !detached && mode !== 'closed' && safeWorkbook() === sessionWb,
  };

  const currentAnchor = (): Addr => anchor ?? { ...deps.getAnchor() };

  const safeWorkbook = (): WorkbookHandle | null => {
    try {
      return deps.getWb();
    } catch {
      return null;
    }
  };

  const isClosed = (): boolean => mode === 'closed';

  const setReturnFocusTarget = (target: HTMLElement | null): void => {
    if (detached || isClosed() || !(target instanceof HTMLElement) || !target.isConnected) return;
    returnFocusTarget = target;
  };

  const readCatalog = (): FunctionCatalogSnapshot => {
    let reader: FunctionCatalogReader | null = null;
    try {
      reader = deps.getWb();
      return buildFunctionCatalog(reader, catalogLocaleOrdinal(deps.getLocale()));
    } catch {
      return buildFunctionCatalog(null, catalogLocaleOrdinal(deps.getLocale()));
    }
  };

  const resolveProjectedName = (rawName: string): string | null => {
    const name = canonicalName(rawName);
    return catalog.knownNames.has(name) ? name : null;
  };

  const activeEndpoint = (view: DraftView): number =>
    view.caret.direction === 'backward' ? view.caret.start : view.caret.end;

  const reprojectDraft = (view: DraftView): ProjectedFormulaCall | null => {
    currentCaret = activeEndpoint(view);
    callProjection = projectFormulaCallAtCaret(view.raw, currentCaret, resolveProjectedName);
    return callProjection;
  };

  const sameProjectedCall = (
    left: ProjectedFormulaCall | null,
    right: ProjectedFormulaCall | null,
  ): boolean => {
    if (left === null || right === null) return left === right;
    return (
      left.source === right.source &&
      left.span.canonicalName === right.span.canonicalName &&
      left.span.name.start === right.span.name.start &&
      left.span.name.end === right.span.name.end &&
      left.span.call.start === right.span.call.start &&
      left.span.call.end === right.span.call.end &&
      left.span.openParen === right.span.openParen &&
      left.span.closeParen === right.span.closeParen &&
      left.span.complete === right.span.complete
    );
  };

  const textForEntry = (entry: FunctionCatalogEntry): string =>
    functionDescription(entry, catalogLocaleOrdinal(deps.getLocale()));

  const argumentHelp = (entry: FunctionCatalogEntry, index: number): FunctionArgumentHelp =>
    resolveArgumentHelp(
      deps.getFunctionArgumentHelp,
      entry,
      index,
      deps.getLocale(),
      labels.argument,
    );

  const notifyMirror = (raw: string | null): void => {
    if (!anchor) return;
    try {
      deps.projectMirror({ ...anchor }, raw);
    } catch {
      // A mirror is presentation-only and cannot interrupt draft ownership.
    }
  };

  const clearRoot = (): void => {
    while (root.firstChild) root.removeChild(root.firstChild);
  };

  const selectedDescription = (): string => {
    if (!selectedEntry) return '';
    return textForEntry(selectedEntry) ?? '';
  };

  const selectedSyntax = (): string => (selectedEntry ? functionSyntax(selectedEntry) : '');

  const categoryTitle = (category: CatalogFunctionCategory): string =>
    strings.fxDialog[CATEGORY_LABEL_KEY[category]];

  const resetDraftProjection = (): void => {
    explicitArgumentCount = null;
    preserveExplicitArgumentCount = false;
    synchronized = true;
    currentRaw = '';
    currentCaret = 0;
    callProjection = null;
    guardedUnavailable = false;
    staleProjectionWrite = false;
  };

  const resetFunctionSelection = (): void => {
    selectedName = null;
    selectedEntry = null;
    args = [];
    resetDraftProjection();
  };

  const resetPickerSelection = (): void => {
    pickerCategory = 'all';
    pickerSelectionName = null;
    pickerSelectionEntry = null;
  };

  /** True when the selected function left the catalog or became unavailable. */
  const selectedEntryWithdrawn = (): boolean => {
    if (!selectedName) return false;
    return insertableEntry(catalog, selectedName) === null;
  };

  const reconcilePickerSelection = (): void => {
    if (!pickerSelectionName) {
      pickerSelectionEntry = null;
      return;
    }
    const entry = insertableEntry(catalog, pickerSelectionName);
    if (!entry) {
      pickerSelectionName = null;
      pickerSelectionEntry = null;
      return;
    }
    pickerSelectionEntry = entry;
  };

  const updateDataset = (): void => {
    root.dataset.state = mode;
    root.dataset.rawSynchronized = synchronized ? 'true' : 'false';
    root.hidden = mode === 'closed';
    deps.dock.hidden = mode === 'closed' || detached;
    root.setAttribute('aria-label', labels.title);
  };

  const updateFieldValues = (): void => {
    const fields = argumentsRefs?.fields.querySelectorAll<HTMLInputElement>(
      'input[data-argument-index]',
    );
    fields?.forEach((field) => {
      const index = Number(field.dataset.argumentIndex);
      field.value = args[index] ?? '';
    });
  };

  const renderPreview = (): void => {
    const preview = argumentsRefs?.preview;
    if (!preview || !selectedEntry) return;
    preview.textContent = '';
    const pendingRequired =
      synchronized && args.slice(0, selectedEntry.minArity).some((value) => value.trim() === '');
    if (pendingRequired || !currentRaw.trim().startsWith('=')) {
      preview.textContent = labels.pending;
      return;
    }
    try {
      const workbook = safeWorkbook();
      if (!workbook) {
        preview.textContent = labels.pending;
        return;
      }
      const result = workbook.evaluateFormulaText(currentAnchor(), currentRaw);
      if (!result.status.ok) {
        preview.textContent = labels.pending;
        return;
      }
      preview.textContent = formatCell(fromEngineValue(result.value), deps.getLocale());
    } catch {
      preview.textContent = labels.pending;
    }
  };

  const sameArgs = (left: readonly string[], right: readonly string[]): boolean =>
    left.length === right.length && left.every((value, index) => value === right[index]);

  /** Adopt one authoritative draft snapshot without causing a render loop. */
  const adoptDraftView = (view: DraftView, preserveArgument?: number | null): boolean => {
    const previousName = selectedName;
    const previousArgs = args;
    const previousProjection = callProjection;
    const previousSynchronized = synchronized;
    currentRaw = view.raw;
    const projection = reprojectDraft(view);

    if (selectedName) {
      if (projection) {
        const nextEntry = catalog.entries.get(projection.span.canonicalName) ?? null;
        selectedName = projection.span.canonicalName;
        selectedEntry = nextEntry;
        args = [...projection.args];
        explicitArgumentCount = projection.args.length;
        if (!writingDraftRaw) preserveExplicitArgumentCount = true;
        guardedUnavailable =
          nextEntry === null || isFunctionUnavailableForInsertion(nextEntry.availability);
        synchronized = !guardedUnavailable;
        if (guardedUnavailable) {
          explicitArgumentCount = null;
          preserveExplicitArgumentCount = false;
        }
      } else {
        const nextEntry = catalog.entries.get(selectedName) ?? null;
        guardedUnavailable =
          nextEntry === null || isFunctionUnavailableForInsertion(nextEntry.availability);
        synchronized = false;
        explicitArgumentCount = null;
        preserveExplicitArgumentCount = false;
      }
    } else {
      guardedUnavailable = false;
      synchronized = projection !== null || view.raw.trim() === '=';
    }

    if (preserveArgument !== null && preserveArgument !== undefined) {
      focusedArgument = preserveArgument;
    } else if (projection) {
      focusedArgument = projection.span.activeArgumentIndex;
    }

    return (
      previousName !== selectedName ||
      !sameArgs(previousArgs, args) ||
      previousSynchronized !== synchronized ||
      previousProjection?.span.canonicalName !== projection?.span.canonicalName ||
      previousProjection?.span.call.start !== projection?.span.call.start ||
      previousProjection?.span.call.end !== projection?.span.call.end
    );
  };

  const sectionTitle = (key: string): string =>
    key === 'recent'
      ? labels.recent
      : key === 'all'
        ? labels.all
        : categoryTitle(key as CatalogFunctionCategory);

  const pickerViewState = (): PickerViewState => {
    const recent = getRecentFunctions(deps.store, catalog.knownNames);
    const sections = pickerSections(pickerCategory, catalog, recent, searchQuery).map(
      (section) => ({
        key: section.key,
        title: sectionTitle(section.key),
        rows: section.names.flatMap((name) => {
          const entry = catalog.entries.get(name);
          if (!entry) return [];
          return [
            {
              name: entry.canonicalName,
              displayName: entry.displayName,
              description: textForEntry(entry),
              unavailable: isFunctionUnavailableForInsertion(entry.availability),
            },
          ];
        }),
      }),
    );
    const entry = pickerSelectionEntry ?? selectedEntry;
    return {
      searchQuery,
      guardMessage,
      activeName: pickerSelectionName ?? selectedName,
      sections,
      summary: entry
        ? {
            displayName: entry.displayName,
            description: textForEntry(entry),
            syntax: functionSyntax(entry),
          }
        : null,
      insertDisabled:
        (pickerSelectionName ?? selectedName) === null ||
        entry === null ||
        isFunctionUnavailableForInsertion(entry.availability),
    };
  };

  const argumentsViewState = (): ArgumentsViewState => {
    const entry = selectedEntry;
    const count = entry ? argumentFieldCount(entry, args.length) : args.length;
    const fields: ArgumentsViewState['fields'] = [];
    for (let index = 0; index < count; index += 1) {
      const help = entry
        ? argumentHelp(entry, index)
        : { label: `${labels.argument} ${index + 1}`, description: undefined };
      fields.push({
        label: help.label ?? `${labels.argument} ${index + 1}`,
        description: help.description,
        value: args[index] ?? '',
      });
    }
    return {
      title: entry?.displayName ?? selectedName ?? '',
      fields,
      canAddArgument: entry?.maxArity === null,
      doneDisabled:
        selectedUnavailable() ||
        mode === 'arguments-committed' ||
        !synchronized ||
        callProjection === null ||
        !callProjection.span.complete,
      doneDisabledReason: selectedUnavailable() ? labels.unavailable : labels.draftConflict,
      guardMessage,
      description: selectedDescription(),
      syntax: selectedSyntax(),
      helpUrl: (entry ? argumentHelp(entry, 0) : null)?.url,
    };
  };

  const selectedUnavailable = (): boolean =>
    guardedUnavailable ||
    (selectedEntry !== null && isFunctionUnavailableForInsertion(selectedEntry.availability));

  /** Drop a stale workbook/context without running the user cancel/restore path. */
  const discardSessionInertly = (): void => {
    const binding = activeDraft;
    if (binding) {
      guarding = true;
      try {
        binding.handle.discard();
      } finally {
        guarding = false;
      }
      return;
    }
    mode = 'closed';
    resetFunctionSelection();
    resetPickerSelection();
    postCommit = null;
    notifyMirror(null);
    recentUnsubscribe?.();
    recentUnsubscribe = null;
    pickerRefs = null;
    argumentsRefs = null;
    clearRoot();
    updateDataset();
    opener = null;
    restoredFocusTarget = null;
    returnFocusTarget = null;
    anchor = null;
    sessionWb = null;
  };

  /** Keep the authoritative draft alive while a same-workbook catalog entry is withdrawn. */
  const showGuardedPicker = (message: string): void => {
    const live = safeWorkbook();
    sessionWb = live;
    catalog = readCatalog();
    const nextEntry = selectedName ? (catalog.entries.get(selectedName) ?? null) : null;
    if (nextEntry) selectedEntry = nextEntry;
    guardedUnavailable = true;
    synchronized = false;
    resetPickerSelection();
    guardMessage = message;
    notifyMirror(activeDraft ? currentRaw : null);
    render();
  };

  const liveEntryGuard = (): FunctionCatalogEntry | null => {
    if (!selectedName) return null;
    const live = safeWorkbook();
    const nextCatalog = readCatalog();
    const nextEntry = nextCatalog.entries.get(selectedName) ?? null;
    const workbookChanged = live === null || live !== sessionWb;
    const mismatch =
      workbookChanged ||
      !nextCatalog.knownNames.has(selectedName) ||
      nextEntry === null ||
      selectedEntry === null ||
      !sameAvailability(selectedEntry, nextEntry);
    catalog = nextCatalog;
    reconcilePickerSelection();
    if (mismatch) {
      if (workbookChanged) {
        discardSessionInertly();
        return null;
      }
      showGuardedPicker(labels.unavailable);
      return null;
    }
    sessionWb = live;
    selectedEntry = nextEntry;
    guardedUnavailable = false;
    return nextEntry;
  };

  const setDraftRaw = (
    raw: string,
    preserveExplicit = preserveExplicitArgumentCount,
    caret?: number,
  ): void => {
    const binding = activeDraft;
    if (!binding) return;
    preserveExplicitArgumentCount = preserveExplicit;
    writingDraftRaw = true;
    try {
      binding.handle.setValue(raw, caret);
    } finally {
      writingDraftRaw = false;
    }
    if (activeDraft === binding) currentRaw = raw;
  };

  const handleRawUpdate = (binding: DraftBinding, view: DraftView): void => {
    if (activeDraft !== binding) return;
    const viewCaret = activeEndpoint(view);
    const preserveArgument = paletteWriteFocusIndex;
    if (currentRaw === view.raw && currentCaret === viewCaret) {
      notifyMirror(view.raw);
      updateDataset();
      updateFieldValues();
      renderPreview();
      return;
    }
    const projectionChanged = adoptDraftView(view, preserveArgument);
    notifyMirror(view.raw);
    updateDataset();
    if (mode !== 'picker' && projectionChanged && preserveArgument === null) {
      render();
      return;
    }
    updateFieldValues();
    renderPreview();
  };

  const finishDraft = (
    binding: DraftBinding,
    outcome: 'committed' | 'cancelled' | 'discarded',
  ): void => {
    if (activeDraft !== binding) return;
    if (outcome === 'discarded') {
      activeDraft = null;
      binding.unsubscribe();
      recentUnsubscribe?.();
      recentUnsubscribe = null;
      notifyMirror(null);
      mode = 'closed';
      resetFunctionSelection();
      resetPickerSelection();
      postCommit = null;
      restoredFocusTarget = null;
      opener = null;
      returnFocusTarget = null;
      sessionWb = null;
      anchor = null;
      pickerRefs = null;
      argumentsRefs = null;
      clearRoot();
      updateDataset();
      return;
    }
    const committedName = selectedName;
    const committedRaw = currentRaw;
    const committedProjection =
      callProjection &&
      committedName &&
      callProjection.source === committedRaw &&
      callProjection.span.canonicalName === committedName
        ? callProjection
        : null;
    const committedArgs = [...(committedProjection?.args ?? args)];
    const committedResult = argumentsRefs?.preview.textContent ?? labels.pending;
    const committedCount = explicitArgumentCount;
    const committedPreservesCount = preserveExplicitArgumentCount;
    const parsed =
      outcome === 'committed' && committedName ? parseOuterCall(committedRaw, committedName) : null;
    const assembled =
      parsed && committedName
        ? formulaWithArgumentCount(
            committedName,
            parsed.args,
            committedPreservesCount ? committedCount : null,
          )
        : null;
    const committedValid =
      outcome === 'committed' &&
      committedName !== null &&
      (committedProjection?.span.complete || parsed !== null);
    const shouldRecord =
      committedValid && (committedProjection !== null || assembled === committedRaw.trim());
    activeDraft = null;
    binding.unsubscribe();
    if (outcome === 'committed' && committedName) {
      postCommit = {
        name: committedName,
        args: committedArgs,
        raw: committedRaw,
        result: committedResult,
        caret: committedProjection?.span.call.end ?? currentCaret,
      };
      mode = 'arguments-committed';
      synchronized = committedValid;
      explicitArgumentCount = committedProjection?.args.length ?? parsed?.args.length ?? null;
      preserveExplicitArgumentCount = committedProjection !== null || parsed !== null;
      guardMessage = '';
      if (shouldRecord) recordRecentFunction(deps.store, committedName, catalog.knownNames);
      notifyMirror(null);
      if (detached || closing || mode !== 'arguments-committed') return;
      render();
      return;
    }
    notifyMirror(null);
    if (outcome === 'cancelled' && !closing && !guarding && !committing && mode !== 'closed') {
      mode = 'picker';
      resetFunctionSelection();
      render();
    }
  };

  const startDraft = (seed: string, lease?: FormulaEditLease, initialCaret?: number): boolean => {
    if (detached || activeDraft || !anchor) return activeDraft !== null;
    const handle = deps.beginDraft(
      { ...anchor },
      seed,
      {
        onFinish: (outcome, restored) => {
          if (binding) finishDraft(binding, outcome);
          if (!restored) return;
          // A restored edit owns the interaction again, so the palette session ends.
          restoredFocusTarget = restored;
          if (!closing && mode !== 'closed') close();
        },
      },
      lease ? { lease } : undefined,
    );
    if (!handle) return false;
    let binding!: DraftBinding;
    binding = {
      handle,
      unsubscribe: () => {},
      leased: lease !== undefined,
    };
    activeDraft = binding;
    const initialView = handle.snapshot();
    if (!initialView) {
      if (activeDraft === binding) activeDraft = null;
      return false;
    }
    binding.unsubscribe = handle.subscribe((view) => handleRawUpdate(binding, view));
    handleRawUpdate(binding, initialView);
    if (initialCaret !== undefined && initialCaret !== activeEndpoint(initialView)) {
      handle.setValue(seed, initialCaret);
    }
    return true;
  };

  const ensurePostCommitDraft = (): boolean => {
    if (activeDraft) return true;
    if ((mode !== 'arguments-committed' && mode !== 'picker') || !postCommit) return false;
    if (!startDraft(postCommit.raw, undefined, postCommit.caret)) {
      args = [...postCommit.args];
      currentRaw = postCommit.raw;
      callProjection = projectFormulaCallAtCaret(
        postCommit.raw,
        postCommit.caret,
        resolveProjectedName,
      );
      const parsed = selectedName ? parseOuterCall(currentRaw, selectedName) : null;
      explicitArgumentCount = callProjection?.args.length ?? parsed?.args.length ?? null;
      preserveExplicitArgumentCount = callProjection !== null || parsed !== null;
      synchronized =
        (callProjection?.span.canonicalName === selectedName && callProjection.span.complete) ||
        parsed !== null;
      guardMessage = labels.draftConflict;
      notifyMirror(null);
      render();
      return false;
    }
    mode = 'arguments-editing';
    const binding = currentDraftBinding();
    if (!binding) return false;
    const view = binding.handle.snapshot();
    const projection = view ? reprojectDraft(view) : null;
    const parsed = parseOuterCall(currentRaw, postCommit.name);
    synchronized =
      (projection?.span.canonicalName === postCommit.name && projection.span.complete) ||
      parsed !== null;
    explicitArgumentCount = projection?.args.length ?? parsed?.args.length ?? null;
    preserveExplicitArgumentCount = projection !== null || parsed !== null;
    if (projection?.span.canonicalName === postCommit.name) {
      args = [...projection.args];
      focusedArgument = projection.span.activeArgumentIndex;
    }
    guardMessage = '';
    updateDataset();
    if (argumentsRefs) {
      projectDisabledState(
        argumentsRefs.done,
        selectedUnavailable(),
        selectedUnavailable() ? labels.unavailable : null,
      );
    }
    return true;
  };

  const projectionForWrite = (): ProjectedFormulaCall | null => {
    staleProjectionWrite = false;
    const binding = activeDraft;
    if (!binding) return null;
    const view = binding.handle.snapshot();
    if (!view) return null;
    const snapshotCaret = activeEndpoint(view);
    const snapshotProjection = projectFormulaCallAtCaret(
      view.raw,
      snapshotCaret,
      resolveProjectedName,
    );
    if (
      currentRaw !== view.raw ||
      currentCaret !== snapshotCaret ||
      !sameProjectedCall(callProjection, snapshotProjection)
    ) {
      staleProjectionWrite = true;
      adoptDraftView(view);
      if (!isClosed() && mode !== 'picker') render();
      return null;
    }
    adoptDraftView(view, focusedArgument);
    const projection = callProjection;
    if (!projection || !selectedName || projection.span.canonicalName !== selectedName) {
      synchronized = false;
      return null;
    }
    return projection;
  };

  /** Rewrite the projected call with `nextArgs`; false when the draft refused the write. */
  const writeProjectedArgs = (
    build: (projected: readonly string[]) => { args: string[]; focus: number },
  ): boolean => {
    const projection = projectionForWrite();
    if (!projection || !selectedName) {
      if (staleProjectionWrite) return false;
      guardMessage = labels.draftConflict;
      render();
      return false;
    }
    const next = build(projection.args);
    const replaced = replaceProjectedFormulaCall(
      currentRaw,
      projection,
      selectedName,
      next.args,
      next.args.length,
    );
    if (!replaced) {
      synchronized = false;
      guardMessage = labels.draftConflict;
      render();
      return false;
    }
    args = next.args;
    explicitArgumentCount = next.args.length;
    preserveExplicitArgumentCount = true;
    focusedArgument = next.focus;
    paletteWriteFocusIndex = next.focus;
    try {
      setDraftRaw(replaced.raw, true, replaced.caret);
    } finally {
      paletteWriteFocusIndex = null;
    }
    return true;
  };

  const withArgument = (projected: readonly string[], index: number, value: string) => {
    const nextArgs = [...projected];
    while (nextArgs.length <= index) nextArgs.push('');
    nextArgs[index] = value;
    return { args: nextArgs, focus: index };
  };

  const onArgumentInput = (index: number, value: string, sourceField?: HTMLInputElement): void => {
    if (
      sourceField &&
      argumentsRefs?.fields.querySelector<HTMLInputElement>(
        `input[data-argument-index="${index}"]`,
      ) !== sourceField
    ) {
      return;
    }
    focusedArgument = index;
    if (mode === 'arguments-committed' && !ensurePostCommitDraft()) return;
    if (mode !== 'arguments-editing' || !activeDraft || !synchronized) return;
    if (!liveEntryGuard() || !activeDraft) return;
    if (!writeProjectedArgs((projected) => withArgument(projected, index, value))) return;
    renderPreview();
  };

  const updatePickerList = (): void => {
    if (pickerRefs) view.renderPickerList();
  };

  const pickEntry = (name: string): void => {
    if (detached || mode !== 'picker') return;
    const currentEntry = insertableEntry(catalog, name);
    if (!currentEntry) {
      reconcilePickerSelection();
      render();
      return;
    }
    pickerSelectionName = currentEntry.canonicalName;
    pickerSelectionEntry = currentEntry;
    guardMessage = '';
    render();
  };

  const addArgument = (): void => {
    if (mode === 'arguments-committed' && !ensurePostCommitDraft()) return;
    if (!synchronized || !activeDraft || !selectedName) return;
    if (!liveEntryGuard() || !activeDraft || !selectedName) return;
    const appended = writeProjectedArgs((projected) => ({
      args: [...projected, ''],
      focus: projected.length,
    }));
    if (!appended) return;
    renderArguments();
    argumentsRefs?.fields
      .querySelector<HTMLInputElement>(`input[data-argument-index="${focusedArgument}"]`)
      ?.focus();
  };

  const renderArguments = (): void => {
    clearRoot();
    updateDataset();
    argumentsRefs = view.renderArguments();
    updateDataset();
    updateFieldValues();
    renderPreview();
  };

  const renderPicker = (): void => {
    clearRoot();
    updateDataset();
    pickerRefs = view.renderPicker();
    argumentsRefs = null;
    updateDataset();
  };

  const render = (): void => {
    const view = activeDraft?.handle.snapshot();
    if (view) adoptDraftView(view);
    updateDataset();
    if (mode === 'closed') {
      clearRoot();
      pickerRefs = null;
      argumentsRefs = null;
      return;
    }
    if (mode === 'picker') renderPicker();
    else renderArguments();
  };

  const returnToPicker = (): void => {
    if (!activeDraft && postCommit && !ensurePostCommitDraft()) return;
    if (!activeDraft && !startDraft('=')) return;
    // Show All changes only the palette selection. The live draft, raw source,
    // caret projection, and lease remain authoritative for the next choice.
    selectedName = null;
    selectedEntry = null;
    args = [];
    postCommit = null;
    resetPickerSelection();
    guardMessage = '';
    mode = 'picker';
    render();
    pickerRefs?.search.focus();
  };

  const insertSelected = (
    requestedName: string | null = selectedName,
    requestedEntry: FunctionCatalogEntry | null = selectedEntry,
  ): void => {
    if (
      !requestedName ||
      !requestedEntry ||
      isFunctionUnavailableForInsertion(requestedEntry.availability)
    ) {
      guardMessage = labels.unavailable;
      render();
      return;
    }

    let raw = currentRaw;
    let projection = callProjection;
    if (mode === 'picker' && activeDraft) {
      const view = activeDraft.handle.snapshot();
      raw = view?.raw ?? activeDraft.handle.value();
      projection = view ? reprojectDraft(view) : null;
      const initialDraft = selectedName === null && raw.trim() === '=';
      if (!initialDraft && !projection) {
        guardMessage = labels.draftConflict;
        synchronized = false;
        render();
        return;
      }
    } else if (mode === 'picker' && postCommit && !activeDraft) {
      if (!ensurePostCommitDraft() || !synchronized) {
        guardMessage = labels.draftConflict;
        render();
        return;
      }
      const binding = currentDraftBinding();
      if (!binding) {
        guardMessage = labels.draftConflict;
        render();
        return;
      }
      const view = binding.handle.snapshot();
      raw = view?.raw ?? currentRaw;
      projection = view ? reprojectDraft(view) : callProjection;
    }

    if (projection && activeDraft && raw !== '=') {
      selectedName = requestedName;
      selectedEntry = requestedEntry;
      pickerSelectionName = null;
      pickerSelectionEntry = null;
      if (!liveEntryGuard() || !activeDraft) return;
      if (projection.span.canonicalName === requestedName) {
        adoptProjectedCall(projection, requestedEntry);
        return;
      }

      const initial = deps.getInitialArguments?.(requestedName) ?? [];
      const count = Math.max(
        requestedEntry.minArity,
        requestedEntry.argumentLabels.length,
        initial.length,
      );
      const initialArgs = Array.from({ length: count }, (_, index) => initial[index] ?? '');
      const replaced = replaceProjectedFormulaCall(
        raw,
        projection,
        requestedName,
        initialArgs,
        null,
      );
      if (!replaced) {
        synchronized = false;
        guardMessage = labels.draftConflict;
        render();
        return;
      }
      args = initialArgs;
      explicitArgumentCount = null;
      preserveExplicitArgumentCount = false;
      synchronized = true;
      guardMessage = '';
      mode = 'arguments-editing';
      postCommit = null;
      setDraftRaw(replaced.raw, false, replaced.caret);
      render();
      const first = argumentsRefs?.fields.querySelector<HTMLInputElement>(
        'input[data-argument-index="0"]',
      );
      first?.focus();
      return;
    }

    selectedName = requestedName;
    selectedEntry = requestedEntry;
    pickerSelectionName = null;
    pickerSelectionEntry = null;
    if (!liveEntryGuard()) return;
    if (!activeDraft && !startDraft('=')) return;
    const initial = deps.getInitialArguments?.(selectedName) ?? [];
    const count = Math.max(
      selectedEntry.minArity,
      selectedEntry.argumentLabels.length,
      initial.length,
    );
    args = Array.from({ length: count }, (_, index) => initial[index] ?? '');
    explicitArgumentCount = null;
    preserveExplicitArgumentCount = false;
    synchronized = true;
    guardMessage = '';
    const formula = assembledFormula(selectedName, args);
    setDraftRaw(formula);
    mode = 'arguments-editing';
    postCommit = null;
    render();
    const first = argumentsRefs?.fields.querySelector<HTMLInputElement>(
      'input[data-argument-index="0"]',
    );
    first?.focus();
  };

  const focusCurrentSurface = (): void => {
    if (mode === 'picker') {
      pickerRefs?.search.focus();
      return;
    }
    const field =
      argumentsRefs?.fields.querySelector<HTMLInputElement>(
        `input[data-argument-index="${focusedArgument}"]`,
      ) ?? argumentsRefs?.fields.querySelector<HTMLInputElement>('input[data-argument-index]');
    field?.focus();
  };

  const rereadOpenCatalog = (): boolean => {
    strings = deps.getStrings();
    labels = paletteStrings(strings);
    const live = safeWorkbook();
    catalog = readCatalog();
    reconcilePickerSelection();
    if (live !== sessionWb) {
      discardSessionInertly();
      return false;
    }
    if (mode !== 'picker' && selectedEntryWithdrawn()) {
      showGuardedPicker(labels.unavailable);
      return false;
    }
    selectedEntry = selectedName ? (catalog.entries.get(selectedName) ?? null) : null;
    guardedUnavailable = false;
    const binding = activeDraft;
    const view = binding?.handle.snapshot();
    if (binding && view) handleRawUpdate(binding, view);
    return true;
  };

  const adoptProjectedCall = (
    projection: ProjectedFormulaCall,
    entry: FunctionCatalogEntry,
  ): boolean => {
    selectedName = entry.canonicalName;
    selectedEntry = entry;
    if (!liveEntryGuard()) return false;
    args = [...projection.args];
    explicitArgumentCount = projection.args.length;
    preserveExplicitArgumentCount = true;
    synchronized = true;
    focusedArgument = projection.span.activeArgumentIndex;
    pickerSelectionName = null;
    pickerSelectionEntry = null;
    guardMessage = '';
    mode = 'arguments-editing';
    postCommit = null;
    render();
    const field = argumentsRefs?.fields.querySelector<HTMLInputElement>(
      `input[data-argument-index="${focusedArgument}"]`,
    );
    (
      field ?? argumentsRefs?.fields.querySelector<HTMLInputElement>('input[data-argument-index]')
    )?.focus();
    return true;
  };

  const applyOpenRequest = (seedName?: string, options?: FxDialogOpenOptions): void => {
    if (!rereadOpenCatalog()) return;
    if (deps.store.getState().ui.editor.kind !== 'idle') {
      guardMessage = labels.draftConflict;
      render();
      return;
    }
    if (!seedName) {
      if (mode === 'picker' && options?.category === undefined && activeDraft) {
        const view = activeDraft.handle.snapshot();
        const projection = view ? reprojectDraft(view) : null;
        const entry = projection
          ? (catalog.entries.get(projection.span.canonicalName) ?? null)
          : null;
        if (projection && entry && !isFunctionUnavailableForInsertion(entry.availability)) {
          adoptProjectedCall(projection, entry);
          return;
        }
      }
      if (options?.category !== undefined) {
        pickerCategory = options.category;
        pickerSelectionName = null;
        pickerSelectionEntry = null;
        if (mode !== 'picker') mode = 'picker';
      }
      render();
      focusCurrentSurface();
      return;
    }

    const requestedName = canonicalName(seedName);
    const requestedEntry = catalog.entries.get(requestedName) ?? null;
    if (!requestedEntry || isFunctionUnavailableForInsertion(requestedEntry.availability)) {
      guardMessage = labels.unavailable;
      render();
      focusCurrentSurface();
      return;
    }

    if (mode === 'arguments-editing') {
      const view = activeDraft?.handle.snapshot();
      const projection = view ? reprojectDraft(view) : callProjection;
      if (!activeDraft || !selectedName || !projection) {
        guardMessage = labels.unavailable;
        render();
        focusCurrentSurface();
        return;
      }
    } else if (mode === 'arguments-committed') {
      if (!ensurePostCommitDraft() || !synchronized) {
        guardMessage = labels.unavailable;
        render();
        focusCurrentSurface();
        return;
      }
    } else if (mode !== 'picker') {
      return;
    }

    guardMessage = '';
    insertSelected(requestedEntry.canonicalName, requestedEntry);
  };

  const commitSelected = (): void => {
    if (mode !== 'arguments-editing' || !activeDraft || !selectedName || !selectedEntry) return;
    if (selectedUnavailable()) {
      guardMessage = labels.unavailable;
      render();
      return;
    }
    const entry = liveEntryGuard();
    if (!entry || !activeDraft) return;
    const projection = projectionForWrite();
    if (!projection && staleProjectionWrite) return;
    if (!projection?.span.complete) {
      synchronized = false;
      guardMessage = labels.draftConflict;
      render();
      return;
    }
    const binding = activeDraft;
    committing = true;
    let committed = false;
    try {
      committed = binding.handle.commit();
    } finally {
      committing = false;
    }
    if (committed && activeDraft === binding) finishDraft(binding, 'committed');
  };

  const close = (): void => {
    if (detached || mode === 'closed') return;
    closing = true;
    const binding = activeDraft;
    if (binding) binding.handle.cancel();
    closing = false;
    // An invalid or already-discarded session closes inertly and must not
    // return focus to the opener.
    if (binding && activeDraft !== binding && isClosed()) return;
    mode = 'closed';
    resetFunctionSelection();
    resetPickerSelection();
    postCommit = null;
    notifyMirror(null);
    render();
    const target =
      connectedElement(restoredFocusTarget) ??
      connectedElement(returnFocusTarget) ??
      connectedElement(opener) ??
      deps.host;
    target.focus();
    restoredFocusTarget = null;
    returnFocusTarget = null;
    opener = null;
    sessionWb = null;
    anchor = null;
  };

  const discard = (): void => {
    if (detached || mode === 'closed') return;
    closing = true;
    const binding = activeDraft;
    if (binding) binding.handle.discard();
    closing = false;
    if (isClosed() && activeDraft !== binding) return;
    mode = 'closed';
    resetFunctionSelection();
    resetPickerSelection();
    postCommit = null;
    notifyMirror(null);
    recentUnsubscribe?.();
    recentUnsubscribe = null;
    render();
    restoredFocusTarget = null;
    returnFocusTarget = null;
    opener = null;
    sessionWb = null;
    anchor = null;
  };

  const open = (seedName?: string, options?: FxDialogOpenOptions): void => {
    if (detached) return;
    if (mode !== 'closed') {
      applyOpenRequest(seedName, options);
      return;
    }
    opener = connectedElement(document.activeElement);
    anchor = { ...deps.getAnchor() };
    sessionWb = safeWorkbook();
    resetFunctionSelection();
    resetPickerSelection();
    searchQuery = '';
    guardMessage = '';
    postCommit = null;
    restoredFocusTarget = null;
    returnFocusTarget = null;
    mode = 'picker';
    const lease = deps.suspendActiveEdit?.(leaseContext) ?? undefined;
    if (!startDraft('=', lease) && lease) {
      // The draft refused the lease; hand the edit straight back to its owner.
      lease.userCancel()?.focus();
      if (isClosed()) return;
    }
    recentUnsubscribe?.();
    recentUnsubscribe = subscribeRecentFunctions(deps.store, () => {
      if (mode === 'picker') updatePickerList();
    });
    applyOpenRequest(seedName, options);
  };

  const refresh = (): void => {
    if (detached) return;
    strings = deps.getStrings();
    labels = paletteStrings(strings);
    if (safeWorkbook() !== sessionWb) {
      discardSessionInertly();
      return;
    }
    catalog = readCatalog();
    reconcilePickerSelection();
    if (mode !== 'picker' && selectedEntryWithdrawn()) {
      showGuardedPicker(labels.unavailable);
      return;
    }
    selectedEntry = selectedName ? (catalog.entries.get(selectedName) ?? null) : null;
    guardedUnavailable = false;
    const binding = activeDraft;
    const view = binding?.handle.snapshot();
    if (binding && view) handleRawUpdate(binding, view);
    render();
  };

  const setStrings = (next: Strings): void => {
    strings = next;
    labels = paletteStrings(strings);
    render();
  };

  const rangeInsertTarget = (): RangeInsertTarget | null => {
    if ((mode !== 'arguments-editing' && mode !== 'arguments-committed') || !selectedName)
      return null;
    return {
      isFormulaEdit: () => Boolean((activeDraft || postCommit) && synchronized && callProjection),
      insertRefAtCaret: (ref: string) => {
        if (mode === 'arguments-committed' && !ensurePostCommitDraft()) return;
        if (!activeDraft || !synchronized) return;
        if (!liveEntryGuard() || !activeDraft || !selectedName) return;
        const index = focusedArgument;
        if (!writeProjectedArgs((projected) => withArgument(projected, index, ref))) return;
        renderPreview();
      },
    };
  };

  const detach = (): void => {
    if (detached) return;
    detached = true;
    closing = true;
    if (activeDraft) activeDraft.handle.discard();
    closing = false;
    recentUnsubscribe?.();
    recentUnsubscribe = null;
    notifyMirror(null);
    deps.dock.hidden = true;
    root.remove();
    mode = 'closed';
    opener = null;
    restoredFocusTarget = null;
    returnFocusTarget = null;
    anchor = null;
    sessionWb = null;
  };

  const api: MacFormulaPaletteHandle = {
    open,
    close,
    discard,
    refresh,
    detach,
    isOpen: () => mode !== 'closed',
    rangeInsertTarget,
    setStrings,
    setReturnFocusTarget,
  };
  return api;
}
