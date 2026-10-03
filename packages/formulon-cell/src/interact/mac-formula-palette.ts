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
  parseOuterCall,
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
  isOpen(): boolean;
  rangeInsertTarget(): RangeInsertTarget | null;
  setStrings(next: Strings): void;
}

type PaletteMode = 'closed' | 'picker' | 'arguments-editing' | 'arguments-committed';

interface CommitSnapshot {
  name: string;
  args: string[];
  raw: string;
  result: string;
}

interface DraftBinding {
  handle: ExternalFormulaDraftHandle;
  unsubscribe: () => void;
  leased: boolean;
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
    argumentInput: (index, value) => onArgumentInput(index, value),
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

  const readCatalog = (): FunctionCatalogSnapshot => {
    let reader: FunctionCatalogReader | null = null;
    try {
      reader = deps.getWb();
      return buildFunctionCatalog(reader, catalogLocaleOrdinal(deps.getLocale()));
    } catch {
      return buildFunctionCatalog(null, catalogLocaleOrdinal(deps.getLocale()));
    }
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
      doneDisabled: selectedUnavailable() || mode === 'arguments-committed',
      guardMessage,
      description: selectedDescription(),
      syntax: selectedSyntax(),
      helpUrl: (entry ? argumentHelp(entry, 0) : null)?.url,
    };
  };

  const selectedUnavailable = (): boolean =>
    selectedEntry !== null && isFunctionUnavailableForInsertion(selectedEntry.availability);

  const showGuardedPicker = (message: string): void => {
    guarding = true;
    if (activeDraft) activeDraft.handle.cancel();
    guarding = false;
    const live = safeWorkbook();
    sessionWb = live;
    catalog = readCatalog();
    selectedEntry = selectedName ? (catalog.entries.get(selectedName) ?? null) : null;
    resetPickerSelection();
    mode = 'picker';
    guardMessage = message;
    resetDraftProjection();
    notifyMirror(null);
    render();
  };

  const liveEntryGuard = (): FunctionCatalogEntry | null => {
    if (!selectedName) return null;
    const live = safeWorkbook();
    const nextCatalog = readCatalog();
    const nextEntry = nextCatalog.entries.get(selectedName) ?? null;
    const mismatch =
      live === null ||
      live !== sessionWb ||
      !nextCatalog.knownNames.has(selectedName) ||
      nextEntry === null ||
      selectedEntry === null ||
      !sameAvailability(selectedEntry, nextEntry);
    catalog = nextCatalog;
    reconcilePickerSelection();
    if (mismatch) {
      showGuardedPicker(labels.unavailable);
      return null;
    }
    sessionWb = live;
    selectedEntry = nextEntry;
    return nextEntry;
  };

  const setDraftRaw = (raw: string, preserveExplicit = preserveExplicitArgumentCount): void => {
    if (!activeDraft) return;
    preserveExplicitArgumentCount = preserveExplicit;
    writingDraftRaw = true;
    try {
      activeDraft.handle.setValue(raw);
    } finally {
      writingDraftRaw = false;
    }
    currentRaw = raw;
  };

  const assembledCurrentFormula = (name: string): string =>
    formulaWithArgumentCount(
      name,
      args,
      preserveExplicitArgumentCount ? explicitArgumentCount : null,
    );

  const handleRawUpdate = (binding: DraftBinding, raw: string): void => {
    if (activeDraft !== binding) return;
    currentRaw = raw;
    notifyMirror(raw);
    if (selectedName) {
      const parsed = parseOuterCall(raw, selectedName);
      if (parsed) {
        const changedCount = args.length !== parsed.args.length;
        args = parsed.args;
        explicitArgumentCount = parsed.args.length;
        if (!writingDraftRaw) preserveExplicitArgumentCount = true;
        synchronized = true;
        if (changedCount && mode !== 'picker') render();
        updateFieldValues();
      } else {
        synchronized = false;
        explicitArgumentCount = null;
        preserveExplicitArgumentCount = false;
      }
    } else {
      synchronized = true;
    }
    updateDataset();
    renderPreview();
  };

  const finishDraft = (binding: DraftBinding, outcome: 'committed' | 'cancelled'): void => {
    if (activeDraft !== binding) return;
    const committedName = selectedName;
    const committedRaw = currentRaw;
    const committedArgs = [...args];
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
    const shouldRecord = parsed !== null && assembled === committedRaw.trim();
    activeDraft = null;
    binding.unsubscribe();
    if (outcome === 'committed' && committedName) {
      postCommit = {
        name: committedName,
        args: committedArgs,
        raw: committedRaw,
        result: committedResult,
      };
      mode = 'arguments-committed';
      synchronized = parsed !== null;
      explicitArgumentCount = parsed?.args.length ?? null;
      preserveExplicitArgumentCount = parsed !== null;
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

  const startDraft = (seed: string, lease?: FormulaEditLease): boolean => {
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
    currentRaw = handle.value();
    binding.unsubscribe = handle.subscribe((raw) => handleRawUpdate(binding, raw));
    notifyMirror(currentRaw);
    return true;
  };

  const ensurePostCommitDraft = (): boolean => {
    if (activeDraft) return true;
    if ((mode !== 'arguments-committed' && mode !== 'picker') || !postCommit) return false;
    if (!startDraft(postCommit.raw)) {
      args = [...postCommit.args];
      currentRaw = postCommit.raw;
      const parsed = selectedName ? parseOuterCall(currentRaw, selectedName) : null;
      explicitArgumentCount = parsed?.args.length ?? null;
      preserveExplicitArgumentCount = parsed !== null;
      synchronized = parsed !== null;
      guardMessage = labels.draftConflict;
      notifyMirror(null);
      render();
      return false;
    }
    mode = 'arguments-editing';
    const parsed = parseOuterCall(currentRaw, postCommit.name);
    synchronized = parsed !== null;
    explicitArgumentCount = parsed?.args.length ?? null;
    preserveExplicitArgumentCount = parsed !== null;
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

  const onArgumentInput = (index: number, value: string): void => {
    focusedArgument = index;
    if (mode === 'arguments-committed' && !ensurePostCommitDraft()) return;
    if (mode !== 'arguments-editing' || !activeDraft || !synchronized) return;
    args[index] = value;
    if (preserveExplicitArgumentCount) {
      explicitArgumentCount = Math.max(explicitArgumentCount ?? 0, index + 1);
    }
    setDraftRaw(assembledCurrentFormula(selectedName ?? ''));
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
    args.push('');
    preserveExplicitArgumentCount = true;
    explicitArgumentCount = Math.max(explicitArgumentCount ?? 0, args.length);
    const raw = assembledCurrentFormula(selectedName);
    setDraftRaw(raw, true);
    renderArguments();
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
    // A leased draft stays open so Close can still restore the suspended edit.
    const keptDraft = activeDraft?.leased ? activeDraft : null;
    if (activeDraft && !keptDraft) {
      closing = true;
      activeDraft.handle.cancel();
      closing = false;
    }
    resetFunctionSelection();
    resetPickerSelection();
    guardMessage = '';
    mode = 'picker';
    if (keptDraft) setDraftRaw('=');
    else startDraft('=');
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

    if (mode === 'picker' && activeDraft) {
      const raw = activeDraft.handle.value();
      const initialDraft = selectedName === null && raw.trim() === '=';
      const completeDraft =
        selectedName !== null && synchronized && parseOuterCall(raw, selectedName) !== null;
      if (!initialDraft && !completeDraft) {
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
      showGuardedPicker(labels.unavailable);
      return false;
    }
    if (mode !== 'picker' && selectedEntryWithdrawn()) {
      showGuardedPicker(labels.unavailable);
      return false;
    }
    selectedEntry = selectedName ? (catalog.entries.get(selectedName) ?? null) : null;
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
      const raw = activeDraft?.handle.value() ?? currentRaw;
      if (
        !activeDraft ||
        !selectedName ||
        !synchronized ||
        parseOuterCall(raw, selectedName) === null
      ) {
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
    if (activeDraft) activeDraft.handle.cancel();
    closing = false;
    mode = 'closed';
    resetFunctionSelection();
    resetPickerSelection();
    postCommit = null;
    notifyMirror(null);
    render();
    const target = connectedElement(restoredFocusTarget) ?? connectedElement(opener) ?? deps.host;
    target.focus();
    restoredFocusTarget = null;
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
    mode = 'picker';
    const lease = deps.suspendActiveEdit?.(leaseContext) ?? undefined;
    if (!startDraft('=', lease) && lease) {
      // The draft refused the lease; hand the edit straight back to its owner.
      lease.userCancel()?.focus();
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
    catalog = readCatalog();
    reconcilePickerSelection();
    if (mode !== 'picker' && selectedEntryWithdrawn()) {
      showGuardedPicker(labels.unavailable);
      return;
    }
    selectedEntry = selectedName ? (catalog.entries.get(selectedName) ?? null) : null;
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
      isFormulaEdit: () => Boolean((activeDraft || postCommit) && synchronized),
      insertRefAtCaret: (ref: string) => {
        if (mode === 'arguments-committed' && !ensurePostCommitDraft()) return;
        if (!activeDraft || !synchronized) return;
        args[focusedArgument] = ref;
        if (preserveExplicitArgumentCount) {
          explicitArgumentCount = Math.max(explicitArgumentCount ?? 0, focusedArgument + 1);
        }
        updateFieldValues();
        setDraftRaw(assembledCurrentFormula(selectedName ?? ''));
        renderPreview();
      },
    };
  };

  const detach = (): void => {
    if (detached) return;
    detached = true;
    closing = true;
    if (activeDraft) activeDraft.handle.cancel();
    closing = false;
    recentUnsubscribe?.();
    recentUnsubscribe = null;
    notifyMirror(null);
    deps.dock.hidden = true;
    root.remove();
    mode = 'closed';
    opener = null;
    restoredFocusTarget = null;
    anchor = null;
    sessionWb = null;
  };

  const api: MacFormulaPaletteHandle = {
    open,
    close,
    refresh,
    detach,
    isOpen: () => mode !== 'closed',
    rangeInsertTarget,
    setStrings,
  };
  return api;
}
