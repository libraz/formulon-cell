import {
  planSelectionFormat,
  type SelectionFormatAction,
  type SelectionFormatPlan,
} from '../commands/format.js';
import type { History } from '../commands/history.js';
import { interactionControllerFor } from '../commands/interaction-controller.js';
import { expandRangeWithMerges, mergeAt, mergeWillLoseData } from '../commands/merge.js';
import { addrKey } from '../engine/address.js';
import type { CellValue, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import { sameRange } from '../store/selection-geometry.js';
import type { CellFormat, SpreadsheetStore, State } from '../store/store.js';
import { confirmMergeLoseData } from '../toolbar/dialogs/merge-confirm.js';
import { formatA1Range } from '../wrappers/toolbar-a1.js';
import { syncCustomSelects } from './custom-select.js';
import {
  type DraftState,
  defaultPatternForLocale,
  type NumberCategory,
  normalizeFormatLocale,
  type TabId,
} from './format-dialog-model.js';
import type { PaletteFlyout } from './format-dialog-palette-flyout.js';
import { createFormatPreview } from './format-dialog-preview.js';
import {
  buildDialogDxf,
  buildTouchedDialogPatch,
  type FormatDialogField,
  hydrateDraftFromFormat,
  makeEmptyDraft,
  summarizeDialogFormats,
} from './format-dialog-state.js';
import { attachAlignTab, type MergeSelectionState } from './format-dialog-tabs/align-controller.js';
import { attachBorderTab } from './format-dialog-tabs/border-controller.js';
import type { FormatTabContext, FormatTabController } from './format-dialog-tabs/controller.js';
import { attachFillTab } from './format-dialog-tabs/fill-controller.js';
import { attachFontTab } from './format-dialog-tabs/font-controller.js';
import { attachMoreTab } from './format-dialog-tabs/more-controller.js';
import {
  attachNumberTab,
  numberCategoryDescription,
} from './format-dialog-tabs/number-controller.js';
import { attachProtectionTab } from './format-dialog-tabs/protection-controller.js';
import { runFormatDialogTransaction } from './format-dialog-transaction.js';
import { createFormatDialogView } from './format-dialog-view.js';
import { attachRangePickerButton } from './range-picker-control.js';

export interface FormatDialogDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  strings?: Strings;
  /** Shared history. When provided the OK click pushes one format-snapshot
   *  entry that reverts the entire dialog apply on undo. */
  history?: History | null;
  /** Workbook getter. When provided, format mutations that affect engine
   *  state (data validation rules, cell-XF entries, hyperlinks) are flushed
   *  to the engine on OK so xlsx round-trip is complete. Lazy so the dialog
   *  stays in lockstep with `setWorkbook` swaps. */
  getWb?: () => WorkbookHandle | null;
  /** Locale used for number/date previews and locale-specific format presets. */
  getLocale?: () => string;
}

export interface FormatDialogOpenOptions {
  mode?: 'format' | 'dataValidation' | 'dxf';
  focus?: 'activeTab' | 'validation';
  /** Initial differential format when editing a conditional-format rule. */
  initialFormat?: Partial<CellFormat>;
  /** Receives the editable differential-format dimensions instead of writing
   *  to the active cell. Used by conditional-format Custom Format. */
  onApplyDxf?: (format: Partial<CellFormat>) => void;
}

export interface FormatDialogHandle {
  open(tab?: TabId, options?: FormatDialogOpenOptions): void;
  close(): void;
  detach(): void;
}

const mergeSelectionState = (state: State, range: Range): MergeSelectionState => {
  const touching = [...state.merges.byAnchor.values()].filter(
    (merge) =>
      merge.sheet === range.sheet &&
      merge.r0 <= range.r1 &&
      merge.r1 >= range.r0 &&
      merge.c0 <= range.c1 &&
      merge.c1 >= range.c0,
  );
  if (touching.length === 0) return 'none';
  if (touching.length === 1 && touching[0] && sameRange(touching[0], range)) return 'merged';
  return 'mixed';
};

const mergeHasNonAnchorContent = (state: State, range: Range): boolean => {
  const effective = expandRangeWithMerges(state, range);
  for (const [key, cell] of state.data.cells) {
    const parts = key.split(':');
    const sheet = Number(parts[0]);
    const row = Number(parts[1]);
    const col = Number(parts[2]);
    if (
      !Number.isInteger(sheet) ||
      !Number.isInteger(row) ||
      !Number.isInteger(col) ||
      sheet !== effective.sheet ||
      row < effective.r0 ||
      row > effective.r1 ||
      col < effective.c0 ||
      col > effective.c1 ||
      (row === effective.r0 && col === effective.c0)
    ) {
      continue;
    }
    if (cell.formula || cell.value.kind !== 'blank') return true;
  }
  return false;
};

export function attachFormatDialog(deps: FormatDialogDeps): FormatDialogHandle {
  const { host, store } = deps;
  const history = deps.history ?? null;
  const getWb = deps.getWb ?? ((): WorkbookHandle | null => null);
  const getFormatLocale = (): string => normalizeFormatLocale(deps.getLocale?.() ?? 'en-US');
  const strings = deps.strings ?? defaultStrings;
  const t = strings.formatDialog;

  const view = createFormatDialogView({
    host,
    strings,
    t,
    fontLocale: getFormatLocale().startsWith('ja') ? 'ja' : 'en',
  });
  const {
    shell,
    overlay,
    headerTitle,
    preview,
    tabsStrip,
    tabButtons,
    tabPanels,
    hyperlinkSection,
    commentSection,
    validationSection,
    validationKindSelect,
    validationListRangeInput,
    closeBtn,
    okBtn,
    cancelBtn,
    hintBar,
  } = view;
  attachRangePickerButton(validationListRangeInput, {
    label: strings.pivotTableDialog.rangePickerSelect,
    getValue: () => formatA1Range(store.getState().selection.range),
    subscribeToRangeChanges: (listener) => store.subscribe(listener),
    kind: 'format-validation-list-range',
  });

  // ── State ──────────────────────────────────────────────────────────────
  let activeTab: TabId = 'number';
  let applyDxf: ((format: Partial<CellFormat>) => void) | null = null;
  let submitting = false;
  let waitingForConfirmation = false;
  let previewValue: CellValue = { kind: 'blank' };
  let selectionPlanForDialog: SelectionFormatPlan | null = null;
  const mixedFields = new Set<FormatDialogField>();
  const touchedFields = new Set<FormatDialogField>();
  const draft: DraftState = makeEmptyDraft(getFormatLocale());

  const touch = (...fields: FormatDialogField[]): void => {
    for (const field of fields) {
      touchedFields.add(field);
      for (const element of overlay.querySelectorAll<HTMLElement>('[data-fc-mixed-field]')) {
        if (element.dataset.fcMixedField !== field) continue;
        delete element.dataset.fcMixed;
        if (element instanceof HTMLInputElement) element.indeterminate = false;
      }
    }
  };
  const isMixedField = (field: FormatDialogField): boolean =>
    mixedFields.has(field) && !touchedFields.has(field);
  const closePaletteFlyouts = (): void => {
    for (const flyout of paletteFlyouts) flyout.setOpen(false);
  };

  // ── Hydration ──────────────────────────────────────────────────────────
  const hydrateFromActive = (initialFormat?: Partial<CellFormat>): void => {
    const state = store.getState();
    const active = state.selection.active;
    const activeCell = state.data.cells.get(addrKey(active));
    previewValue = activeCell?.value ?? getWb()?.getValue(active) ?? { kind: 'blank' };
    touchedFields.clear();
    mixedFields.clear();
    selectionPlanForDialog = applyDxf ? null : planSelectionFormat(state);
    const summary = applyDxf
      ? { activeFormat: initialFormat ?? {}, mixed: new Set<FormatDialogField>() }
      : summarizeDialogFormats(state, selectionPlanForDialog);
    for (const field of summary.mixed) mixedFields.add(field);
    const fmt = summary.activeFormat;
    hydrateDraftFromFormat(draft, fmt, getFormatLocale());
    borderTab.clearPendingPreset();

    syncControlsFromDraft();
    const range = state.selection.range;
    const activeMerge = mergeAt(state, state.selection.active);
    const multiCell = range.r0 !== range.r1 || range.c0 !== range.c1;
    const mergeRange = expandRangeWithMerges(state, range);
    const hasExtraRanges = (state.selection.extraRanges?.length ?? 0) > 0;
    const mergeDisabled =
      hasExtraRanges ||
      (!multiCell && activeMerge === null) ||
      (!getWb() && mergeHasNonAnchorContent(state, mergeRange));
    alignTab.hydrateMerge(mergeSelectionState(state, range), mergeDisabled);
    renderPreview();
    setActiveTab('number');
  };

  const syncMixedControls = (): void => {
    for (const tab of tabControllers) tab.syncMixed();
  };

  const syncControlsFromDraft = (): void => {
    for (const tab of tabControllers) tab.sync();
    syncMixedControls();
  };

  // ── Preview rendering ──────────────────────────────────────────────────
  const renderPreviewBody = createFormatPreview({
    refs: view,
    draft,
    getLocale: getFormatLocale,
    defaultPatternFor: (cat) => defaultPatternFor(cat),
    previewValue: () => (applyDxf !== null ? null : previewValue),
  });
  const renderPreview = (): void => {
    alignTab.syncRotationDial();
    renderPreviewBody();
  };

  const defaultPatternFor = (cat: NumberCategory): string =>
    defaultPatternForLocale(cat, getFormatLocale());

  // ── Tab switch ─────────────────────────────────────────────────────────
  const tabOrder = Array.from(tabButtons.keys());
  const setActiveTab = (id: TabId): void => {
    activeTab = id;
    closePaletteFlyouts();
    for (const [tabId, btn] of tabButtons) {
      btn.setAttribute('aria-selected', tabId === id ? 'true' : 'false');
      btn.tabIndex = tabId === id ? 0 : -1;
    }
    for (const [tabId, p] of tabPanels) {
      p.hidden = tabId !== id;
    }
    syncHintBar();
  };

  const syncHintBar = (): void => {
    // Only the Number tab carries a per-category description in the hint bar.
    // Other tabs collapse the bar so the body keeps its space.
    if (activeTab === 'number') {
      hintBar.textContent = numberCategoryDescription(draft.numberCategory, t);
    } else {
      hintBar.textContent = '';
    }
  };

  const setDialogMode = (mode: FormatDialogOpenOptions['mode'] = 'format'): void => {
    const dataValidationMode = mode === 'dataValidation';
    const dxfMode = mode === 'dxf';
    overlay.classList.toggle('fc-fmtdlg--data-validation', dataValidationMode);
    // `role="dialog"` sits on the overlay, so the accessible name has to be set
    // there — labelling the panel leaves the dialog announcing the format title
    // while the header reads Data Validation.
    shell.setAriaLabel(dataValidationMode ? t.validationLegend : t.title);
    headerTitle.textContent = dataValidationMode ? t.validationLegend : t.title;
    tabsStrip.hidden = dataValidationMode;
    preview.hidden = dataValidationMode;
    hyperlinkSection.hidden = dataValidationMode;
    commentSection.hidden = dataValidationMode;
    validationSection.classList.toggle('fc-fmtdlg__section--standalone', dataValidationMode);
    for (const [tabId, button] of tabButtons) {
      button.hidden = dxfMode && (tabId === 'align' || tabId === 'protection' || tabId === 'more');
    }
    if (dxfMode && (activeTab === 'align' || activeTab === 'protection' || activeTab === 'more')) {
      setActiveTab('number');
    }
  };

  // ── Apply OK ───────────────────────────────────────────────────────────
  const applyAndClose = async (): Promise<void> => {
    if (applyDxf) {
      applyDxf(buildDialogDxf(draft, defaultPatternFor));
      api.close();
      return;
    }

    const state = store.getState();
    const range = state.selection.range;
    const liveWb = getWb();
    const plan = planSelectionFormat(state);
    if (!plan) return;

    const mergeRange = expandRangeWithMerges(state, range);
    const mergeAction = alignTab.mergeAction();

    // A multi-area selection has no single merge result. The control is also
    // disabled during hydration, but retain this guard for stale event state.
    if ((state.selection.extraRanges?.length ?? 0) > 0 && mergeAction) return;
    // Structural merge authorization has no safe dialog intent yet. A
    // registered restricted controller must reject the complete composite
    // action before formatting or confirmation can mutate anything.
    if (mergeAction && interactionControllerFor(store)?.policy) return;

    // Ask before any format, value, or merge mutation. The no-loss path does
    // not await, preserving the synchronous behavior of existing callers.
    if (mergeAction === 'merge' && !liveWb && mergeHasNonAnchorContent(state, mergeRange)) return;
    if (mergeAction === 'merge' && mergeWillLoseData(state, mergeRange)) {
      waitingForConfirmation = true;
      if (!(await confirmMergeLoseData(strings, state, mergeRange))) return;
    }

    const touchedPatch = buildTouchedDialogPatch(draft, touchedFields, defaultPatternFor);
    let actionPatch: Partial<CellFormat> = touchedPatch;
    let borderAction: SelectionFormatAction['border'];
    const pendingBorderPreset = borderTab.pendingPreset();
    if (pendingBorderPreset) {
      const { borders: _borders, ...withoutBorders } = touchedPatch;
      actionPatch = withoutBorders;
      borderAction = {
        preset: pendingBorderPreset,
        style: draft.borderStyle,
        ...(draft.borderColor !== undefined ? { color: draft.borderColor } : {}),
      };
    }
    const hasFormatAction = Object.keys(actionPatch).length > 0 || borderAction !== undefined;
    if (!hasFormatAction && !mergeAction) {
      api.close();
      return;
    }

    const action: SelectionFormatAction = {
      patch: actionPatch,
      ...(borderAction ? { border: borderAction } : {}),
    };
    const completed = runFormatDialogTransaction({
      store,
      history,
      getWb,
      state,
      liveWb,
      plan,
      action,
      merge: { action: mergeAction, range: mergeRange },
    });
    if (completed) api.close();
  };

  // ── Event handlers ─────────────────────────────────────────────────────
  const onTabClick = (e: MouseEvent): void => {
    const target = e.target as HTMLElement;
    const btn = target.closest('button[data-fc-tab]') as HTMLButtonElement | null;
    if (!btn) return;
    const id = btn.dataset.fcTab as TabId | undefined;
    if (id) {
      setActiveTab(id);
      btn.focus();
    }
  };

  const focusTabByIndex = (idx: number): void => {
    const next = tabOrder[(idx + tabOrder.length) % tabOrder.length];
    if (!next) return;
    setActiveTab(next);
    tabButtons.get(next)?.focus();
  };

  const onTabKeyDown = (e: KeyboardEvent): void => {
    const btn = (e.target as HTMLElement).closest<HTMLButtonElement>('button[data-fc-tab]');
    if (!btn) return;
    const id = btn.dataset.fcTab as TabId | undefined;
    const idx = id ? tabOrder.indexOf(id) : -1;
    if (idx < 0) return;
    if (e.key === 'ArrowRight' || e.key === 'ArrowDown') {
      e.preventDefault();
      focusTabByIndex(idx + 1);
    } else if (e.key === 'ArrowLeft' || e.key === 'ArrowUp') {
      e.preventDefault();
      focusTabByIndex(idx - 1);
    } else if (e.key === 'Home') {
      e.preventDefault();
      focusTabByIndex(0);
    } else if (e.key === 'End') {
      e.preventDefault();
      focusTabByIndex(tabOrder.length - 1);
    }
  };

  const onOverlayPointerDown = (e: Event): void => {
    const target = e.target as Node | null;
    if (!target) return;
    for (const flyout of paletteFlyouts) {
      if (flyout.isOpen() && !flyout.owns(target)) flyout.setOpen(false);
    }
  };

  const onOk = (): void => {
    if (submitting) return;
    waitingForConfirmation = false;
    submitting = true;
    const operation = applyAndClose();
    if (waitingForConfirmation) {
      void operation.finally(() => {
        waitingForConfirmation = false;
        submitting = false;
      });
    } else {
      submitting = false;
    }
  };
  const onCancel = (): void => api.close();

  const onOverlayKey = (e: KeyboardEvent): void => {
    e.stopPropagation();
    if (e.key === 'Escape') {
      e.preventDefault();
      // A flyout swallows the first Escape so the dialog itself survives it.
      const open = paletteFlyouts.find((flyout) => flyout.isOpen());
      if (open) {
        open.setOpen(false);
        open.toggle.focus();
        return;
      }
      api.close();
      return;
    }
    if (e.key === 'Enter') {
      const target = e.target as HTMLElement;
      const tag = target.tagName;
      // Don't intercept Enter inside textarea or buttons that should activate.
      if (tag === 'BUTTON' || tag === 'TEXTAREA') return;
      e.preventDefault();
      onOk();
    }
  };

  // ── Wire up ────────────────────────────────────────────────────────────
  const tabContext: FormatTabContext = {
    draft,
    t,
    on: shell.on,
    touch,
    isMixed: isMixedField,
    syncControls: () => syncControlsFromDraft(),
    renderPreview: () => renderPreview(),
    getLocale: getFormatLocale,
  };
  shell.on(tabsStrip, 'click', onTabClick as EventListener);
  shell.on(tabsStrip, 'keydown', onTabKeyDown as EventListener);
  const numberTab = attachNumberTab(tabContext, view, syncHintBar);
  const alignTab = attachAlignTab(tabContext, view);
  const fontTab = attachFontTab(tabContext, view);
  const borderTab = attachBorderTab(tabContext, view);
  const fillTab = attachFillTab(tabContext, view);
  const protectionTab = attachProtectionTab(tabContext, view);
  const moreTab = attachMoreTab(tabContext, view);
  const tabControllers: readonly FormatTabController[] = [
    numberTab,
    alignTab,
    fontTab,
    borderTab,
    fillTab,
    protectionTab,
    moreTab,
  ];
  const paletteFlyouts: readonly PaletteFlyout[] = [fontTab.palette, borderTab.palette];
  shell.on(closeBtn, 'click', onCancel);
  shell.on(okBtn, 'click', onOk);
  shell.on(cancelBtn, 'click', onCancel);
  shell.on(overlay, 'click', (e) => {
    if ((e as MouseEvent).target === overlay) api.close();
  });
  shell.on(overlay, 'keydown', onOverlayKey as EventListener);
  shell.on(overlay, 'mousedown', onOverlayPointerDown);

  const api: FormatDialogHandle = {
    open(tab?: TabId, options?: FormatDialogOpenOptions): void {
      applyDxf = options?.mode === 'dxf' ? (options.onApplyDxf ?? null) : null;
      hydrateFromActive(options?.initialFormat);
      setDialogMode(options?.mode);
      if (tab && tabButtons.has(tab)) setActiveTab(tab);
      if (options?.mode === 'dataValidation') setActiveTab('more');
      shell.open();
      if (mixedFields.size > 0) {
        syncMixedControls();
        syncCustomSelects(overlay);
      }
      requestAnimationFrame(() => {
        if (options?.focus === 'validation' || options?.mode === 'dataValidation') {
          validationKindSelect.focus();
          validationKindSelect.scrollIntoView({ block: 'nearest' });
          return;
        }
        tabButtons.get(activeTab)?.focus();
      });
    },
    close(): void {
      closePaletteFlyouts();
      shell.close();
      host.focus();
    },
    detach(): void {
      shell.dispose();
    },
  };

  return api;
}
