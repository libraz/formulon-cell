import { authorizeWrite } from '../commands/authorized-write.js';
import {
  canonicalizeLocalizedFormulaForInput,
  inputLocaleForWorkbook,
} from '../commands/coerce-input.js';
import { type SelectionFormatAction, setFont } from '../commands/format.js';
import type { History } from '../commands/history.js';
import {
  expandRangeWithMerges,
  mergeAt,
  mergeIntents,
  mergeWillLoseData,
} from '../commands/merge.js';
import { deferForAllowedRangeUnlock, selectionEditRanges } from '../commands/protection.js';
import {
  planSelectionFormat,
  type SelectionFormatPlan,
} from '../commands/selection-format-plan.js';
import { recordFormatChange } from '../commands/slice-history.js';
import { syncColorPaletteTheme, themeColorColumns } from '../components/color-palette.js';
import { addrKey, parseAddrKey } from '../engine/address.js';
import { displayCellFormatFromXf, effectiveCellFormat } from '../engine/cell-format-projection.js';
import type { Addr, CellValue, Range } from '../engine/types.js';
import { isWholeSheetSelection } from '../engine/whole-sheet-font-sync.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { isMacHost } from '../extensions/ui-options.js';
import { bundledLocaleForStrings } from '../i18n/locale-for-strings.js';
import { defaultStrings, dictionaryLocaleFor, type Strings } from '../i18n/strings.js';
import { rangeValidationAt } from '../store/range-validations.js';
import { rangeContainsAddr, sameRange } from '../store/selection-geometry.js';
import type { CellFormat, NumFmt, SpreadsheetStore, State } from '../store/store.js';
import { confirmMergeLoseData } from '../toolbar/dialogs/merge-confirm.js';
import { createExcelRibbonSvg } from '../toolbar/excel-ribbon-icons.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import { syncCustomSelects } from './custom-select.js';
import { type DialogEntry, INSTANCE_API_FORMAT_ENTRY } from './dialog-entry.js';
import { deferInitialFocus } from './dialog-shell.js';
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
  padsNumberRows,
  summarizeDialogFormats,
} from './format-dialog-state.js';
import { attachAlignTab, type MergeSelectionState } from './format-dialog-tabs/align-controller.js';
import { attachBorderTab } from './format-dialog-tabs/border-controller.js';
import type {
  FormatTabContext,
  FormatTabController,
  MacDxfForm,
} from './format-dialog-tabs/controller.js';
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
import { isSubmitEnter } from './mac-dialog-keys.js';
import { createRangePickFormatter } from './range-pick-format.js';
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
  /** Who opened the dialog; the instance API when absent. */
  entry?: DialogEntry;
  /** `numberFormat` shows only the number tab and hands its result to
   *  `onApplyNumberFormat`, writing nothing to the sheet. */
  mode?: 'format' | 'dataValidation' | 'dxf' | 'numberFormat';
  focus?: 'activeTab' | 'validation';
  /** Initial differential format when editing a conditional-format rule. */
  initialFormat?: Partial<CellFormat>;
  /** Receives the editable differential-format dimensions instead of writing
   *  to the active cell. Used by conditional-format Custom Format. */
  onApplyDxf?: (format: Partial<CellFormat>) => void;
  /** Which differential format `dxf` mode edits. On the Mac profile both open on Font with the
   *  font name, size and super/subscript fixed, a tri-state strikethrough and a Clear button: a
   *  table or PivotTable style element offers only Font, Border and Fill, and a conditional
   *  format keeps Number and also fixes the underline. */
  dxfForm?: MacDxfForm;
  /** Receives the chosen number format in `numberFormat` mode. */
  onApplyNumberFormat?: (numFmt: NumFmt) => void;
  /** The value the sample shows in `numberFormat` mode; the active cell's by default. */
  previewValue?: CellValue;
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
    const addr = parseAddrKey(key);
    if (!addr) continue;
    const { sheet, row, col } = addr;
    if (
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

/** Split `range` into rectangles that cover everything except `cells`
 *  (all of which lie inside `range`). */
const rangeMinusCells = (range: Range, cells: readonly Addr[]): Range[] => {
  const colsByRow = new Map<number, number[]>();
  for (const { row, col } of cells) {
    const cols = colsByRow.get(row);
    if (cols) cols.push(col);
    else colsByRow.set(row, [col]);
  }
  const { sheet } = range;
  const out: Range[] = [];
  let nextRow = range.r0;
  for (const row of [...colsByRow.keys()].sort((a, b) => a - b)) {
    if (row > nextRow) out.push({ ...range, r0: nextRow, r1: row - 1 });
    let nextCol = range.c0;
    for (const col of [...new Set(colsByRow.get(row))].sort((a, b) => a - b)) {
      if (col > nextCol) out.push({ sheet, r0: row, r1: row, c0: nextCol, c1: col - 1 });
      nextCol = col + 1;
    }
    if (nextCol <= range.c1) out.push({ sheet, r0: row, r1: row, c0: nextCol, c1: range.c1 });
    nextRow = row + 1;
  }
  if (nextRow <= range.r1) out.push({ ...range, r0: nextRow });
  return out;
};

/** Read compact native XF bases only for an axis selection whose result can
 *  be compared directly with the active cell's display font. Store-formatted
 *  cells are excluded from the sampled ranges; sheet overrides or merges make
 *  the regular resolver authoritative. */
export function nativeFontSamplesForDialog(
  state: State,
  plan: SelectionFormatPlan,
  wb: WorkbookHandle,
): Partial<CellFormat>[] {
  if (plan.axisBands.length === 0 || !wb.capabilities.rangeStylePatch) return [];
  if (
    plan.ranges.some(
      ({ sheet }) =>
        state.format.sheetFontOverrides?.has(sheet) || state.format.sheetAxisOverrides?.has(sheet),
    )
  ) {
    return [];
  }
  if (
    [...state.merges.byAnchor.values()].some((merge) =>
      plan.ranges.some(
        (range) =>
          range.sheet === merge.sheet &&
          range.r0 <= merge.r1 &&
          range.r1 >= merge.r0 &&
          range.c0 <= merge.c1 &&
          range.c1 >= merge.c0,
      ),
    )
  ) {
    return [];
  }

  // Addresses with store entries are resolved by the store resolver, so
  // their native bases are left out of the sampled ranges.
  const storeAddrs: Addr[] = [];
  for (const key of state.format.formats.keys()) {
    const addr = parseAddrKey(key);
    if (addr) storeAddrs.push(addr);
  }

  try {
    const xfIndices = new Set<number>();
    for (const range of plan.ranges) {
      const inside = storeAddrs.filter((addr) => rangeContainsAddr(range, addr));
      for (const part of rangeMinusCells(range, inside)) {
        const bases = wb.getRangeStyleBases(part);
        if (!bases) return [];
        for (const xfIndex of bases) {
          if (!Number.isInteger(xfIndex) || xfIndex < 0) return [];
          xfIndices.add(xfIndex);
        }
      }
    }
    const samples: Partial<CellFormat>[] = [];
    for (const xfIndex of xfIndices) {
      const sample = displayCellFormatFromXf(wb, xfIndex);
      if (!sample) return [];
      samples.push(sample);
    }
    return samples;
  } catch {
    return [];
  }
}

export function attachFormatDialog(deps: FormatDialogDeps): FormatDialogHandle {
  const { host, store } = deps;
  const history = deps.history ?? null;
  const getWb = deps.getWb ?? ((): WorkbookHandle | null => null);
  const strings = deps.strings ?? defaultStrings;
  const getFormatLocale = (): string => {
    const wb = getWb();
    const wbLocale = wb ? inputLocaleForWorkbook(wb) : undefined;
    return normalizeFormatLocale(
      deps.getLocale?.() ?? wbLocale ?? bundledLocaleForStrings(strings) ?? 'en-US',
    );
  };
  const getDate1904 = (): boolean => store.getState().ui.date1904 === true;
  const t = isMacHost(host)
    ? { ...strings.formatDialog, ...strings.mac.formatDialog }
    : strings.formatDialog;

  const localizeDialogFormula = (formula: string): string => {
    const wb = getWb();
    if (!wb || typeof wb.getFormulaForDisplay !== 'function') return formula;
    try {
      return wb.getFormulaForDisplay(formula);
    } catch {
      return formula;
    }
  };

  const canonicalizeDialogFormula = (formula: string): string => {
    const wb = getWb();
    if (!wb) return formula;
    try {
      return canonicalizeLocalizedFormulaForInput(wb, formula, getFormatLocale());
    } catch {
      return formula;
    }
  };

  const view = createFormatDialogView({
    host,
    strings,
    t,
    fontLocale: dictionaryLocaleFor(getFormatLocale()),
    mac: isMacHost(host),
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
    clearBtn,
    hintBar,
  } = view;
  overlay.classList.toggle('fc-fmtdlg--mac', isMacHost(host));
  // The Mac dialog has no text-direction control.
  if (isMacHost(host)) {
    overlay
      .querySelector<HTMLElement>('.fc-fmtdlg__text-direction-row')
      ?.setAttribute('hidden', '');
  }
  const pickFormat = createRangePickFormatter({
    store,
    sheetName: (index) => getWb()?.sheetName(index) ?? '',
  });
  attachRangePickerButton(validationListRangeInput, {
    label: strings.pivotTableDialog.rangePickerSelect,
    getValue: () => `=${pickFormat.range()}`,
    subscribeToRangeChanges: pickFormat.subscribe,
    kind: 'format-validation-list-range',
  });
  attachRangePickerButton(view.validationFormulaInput, {
    label: strings.pivotTableDialog.rangePickerSelect,
    getValue: () => `=${pickFormat.range()}`,
    subscribeToRangeChanges: pickFormat.subscribe,
    kind: 'format-validation-formula',
  });

  // ── State ──────────────────────────────────────────────────────────────
  let activeTab: TabId = 'number';
  let applyDxf: ((format: Partial<CellFormat>) => void) | null = null;
  // The Mac form of the differential format being edited, if any.
  let macDxf: MacDxfForm | null = null;
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
    const dialogWb = getWb();
    const effectiveFormatCache = new Map<string, Partial<CellFormat> | null>();
    const nativeFontSamples =
      !applyDxf && dialogWb && selectionPlanForDialog
        ? nativeFontSamplesForDialog(state, selectionPlanForDialog, dialogWb)
        : [];
    const summary = applyDxf
      ? { activeFormat: initialFormat ?? {}, mixed: new Set<FormatDialogField>() }
      : summarizeDialogFormats(
          state,
          selectionPlanForDialog,
          dialogWb
            ? (addr) => effectiveCellFormat(dialogWb, state, addr, effectiveFormatCache)
            : undefined,
          nativeFontSamples,
        );
    for (const field of summary.mixed) mixedFields.add(field);
    // Strikethrough is tri-state there: unset until the user picks a side.
    if (macDxf && initialFormat?.strike === undefined) mixedFields.add('strike');
    const displayMerge = mergeAt(state, active);
    const displayAddr = displayMerge
      ? { sheet: displayMerge.sheet, row: displayMerge.r0, col: displayMerge.c0 }
      : active;
    const displayFormat =
      !applyDxf && dialogWb
        ? effectiveCellFormat(dialogWb, state, displayAddr, effectiveFormatCache)
        : null;
    const merged = displayFormat
      ? { ...displayFormat, ...summary.activeFormat }
      : summary.activeFormat;
    const rangeValidation = applyDxf ? undefined : rangeValidationAt(state.format, displayAddr);
    const fmt =
      merged.validation || !rangeValidation ? merged : { ...merged, validation: rangeValidation };
    hydrateDraftFromFormat(
      draft,
      fmt,
      getFormatLocale(),
      localizeDialogFormula,
      isMacHost(host) ? 0 : 2,
    );
    borderTab.clearPendingPreset();

    syncControlsFromDraft();
    const range = state.selection.range;
    const multiCell = range.r0 !== range.r1 || range.c0 !== range.c1;
    const mergeRange = expandRangeWithMerges(state, range);
    const hasExtraRanges = (state.selection.extraRanges?.length ?? 0) > 0;
    const mergeDisabled =
      hasExtraRanges || (!getWb() && mergeHasNonAnchorContent(state, mergeRange));
    alignTab.hydrateMerge(mergeSelectionState(state, range), mergeDisabled);
    borderTab.setInsideAvailable(multiCell || hasExtraRanges);
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
    getDate1904,
    defaultPatternFor: (cat) => defaultPatternFor(cat),
    padNumber: () => padsNumberRows(getWb()),
    // A rule's differential format has no cell to preview; a number format has.
    previewValue: () => (dialogMode === 'dxf' ? null : previewValue),
    fontEffects: () => isMacHost(host),
  });
  // The dialog's differential format; the Mac form leaves a mixed strikethrough unset.
  const buildCurrentDxf = () => {
    const dxf = buildDialogDxf(draft, defaultPatternFor, padsNumberRows(getWb()));
    if (macDxf && isMixedField('strike')) dxf.strike = undefined;
    return dxf;
  };
  const macDxfHasContent = (): boolean => {
    const { numFmt: _numFmt, ...dxf } = buildCurrentDxf();
    return Object.values(dxf).some((v) => v !== undefined && v !== false && v !== '');
  };
  const renderPreview = (): void => {
    alignTab.syncRotationDial();
    renderPreviewBody();
    if (macDxf) projectDisabledState(clearBtn, !macDxfHasContent(), null);
  };

  const defaultPatternFor = (cat: NumberCategory): string =>
    defaultPatternForLocale(cat, getFormatLocale(), isMacHost(host));

  // ── Tab switch ─────────────────────────────────────────────────────────
  const tabOrder = Array.from(tabButtons.keys());
  let lastUsedTab: TabId = 'number';
  const setActiveTab = (id: TabId): void => {
    activeTab = id;
    if (id !== 'more') lastUsedTab = id;
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
      hintBar.dataset.fcCat = draft.numberCategory;
    } else {
      hintBar.textContent = '';
      delete hintBar.dataset.fcCat;
    }
  };

  // The Mac message tabs caption their fields; the caption follows the
  // show-message checkbox of each pane.
  const validationCaptions = new Map<'input' | 'error', HTMLDivElement>();
  const syncValidationCaptions = (show: boolean): void => {
    const text = strings.mac.dataValidation;
    for (const pane of ['input', 'error'] as const) {
      let caption = validationCaptions.get(pane);
      if (!caption && show) {
        const host = validationSection.querySelector<HTMLElement>(
          `.fc-fmtdlg__vpane[data-fc-pane="${pane}"]`,
        );
        const check = host?.querySelector(':scope > .fc-fmtdlg__check');
        if (!host || !check) continue;
        caption = document.createElement('div');
        caption.className = 'fc-fmtdlg__vmsg-caption';
        check.after(caption);
        validationCaptions.set(pane, caption);
      }
      if (!caption) continue;
      caption.hidden = !show;
      caption.textContent = pane === 'input' ? text.inputCaption : text.errorCaption;
    }
    syncValidationAlertIcon(show);
  };

  // The Mac error tab shows the chosen style's alert mark below the pop-up.
  const ALERT_GLYPHS = {
    stop: 'alertStop',
    warning: 'alertWarning',
    information: 'alertInformation',
  } as const;
  let validationAlertIcon: HTMLDivElement | null = null;
  const renderValidationAlertIcon = (): void => {
    if (!validationAlertIcon) return;
    const select = validationAlertIcon.parentElement?.querySelector('select');
    const style = (select?.value ?? 'stop') as keyof typeof ALERT_GLYPHS;
    validationAlertIcon.dataset.style = style;
    const svg = createExcelRibbonSvg(
      ALERT_GLYPHS[style] ?? ALERT_GLYPHS.stop,
      'fc-fmtdlg__valert-svg',
    );
    validationAlertIcon.replaceChildren(...(svg ? [svg] : []));
  };
  const syncValidationAlertIcon = (show: boolean): void => {
    if (!validationAlertIcon && show) {
      const row = validationSection.querySelector<HTMLElement>(
        '.fc-fmtdlg__vpane[data-fc-pane="error"] > .fc-fmtdlg__row:has(select)',
      );
      const select = row?.querySelector('select');
      if (!row || !select) return;
      validationAlertIcon = document.createElement('div');
      validationAlertIcon.className = 'fc-fmtdlg__valert-icon';
      validationAlertIcon.setAttribute('aria-hidden', 'true');
      row.appendChild(validationAlertIcon);
      select.addEventListener('change', renderValidationAlertIcon);
    }
    if (!validationAlertIcon) return;
    validationAlertIcon.hidden = !show;
    renderValidationAlertIcon();
  };

  let showMoreTab = false;
  let dialogMode: NonNullable<FormatDialogOpenOptions['mode']> = 'format';
  let dialogEntry: DialogEntry = INSTANCE_API_FORMAT_ENTRY;
  const setDialogMode = (mode: FormatDialogOpenOptions['mode'] = 'format'): void => {
    const dataValidationMode = mode === 'dataValidation';
    const dxfMode = mode === 'dxf';
    clearBtn.hidden = macDxf === null;
    const numberMode = mode === 'numberFormat';
    overlay.classList.toggle('fc-fmtdlg--data-validation', dataValidationMode);
    overlay.classList.toggle('fc-fmtdlg--number-format', numberMode);
    overlay.classList.toggle('fc-fmtdlg--mac-dxf', macDxf !== null);
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
    syncValidationCaptions(dataValidationMode && isMacHost(host));
    for (const [tabId, button] of tabButtons) {
      // The Mac app has no seventh tab; it opens only when asked for by name.
      const macHidden = isMacHost(host) && tabId === 'more' && !showMoreTab;
      button.hidden =
        macHidden ||
        (macDxf === 'styleElement' && tabId === 'number') ||
        (numberMode && tabId !== 'number') ||
        (dxfMode && (tabId === 'align' || tabId === 'protection' || tabId === 'more'));
    }
    if (macDxf) {
      setActiveTab('font');
    } else if (
      numberMode ||
      (dxfMode && (activeTab === 'align' || activeTab === 'protection' || activeTab === 'more'))
    ) {
      setActiveTab('number');
    }
  };

  // ── Apply OK ───────────────────────────────────────────────────────────
  const applyAndClose = async (): Promise<void> => {
    if (applyDxf) {
      applyDxf(buildCurrentDxf());
      api.close();
      return;
    }

    const state = store.getState();
    // The dialog stays open while a password-guarded range under the selection is unlocked.
    if (deferForAllowedRangeUnlock(store, selectionEditRanges(state), () => void applyAndClose())) {
      return;
    }
    const range = state.selection.range;
    const liveWb = getWb();
    const plan = planSelectionFormat(state);
    const selectedRanges = [range, ...(state.selection.extraRanges ?? [])];

    const mergeRange = expandRangeWithMerges(state, range);
    // Merging a lone cell is a no-op, as in the spreadsheet app.
    const loneCell = mergeRange.r0 === mergeRange.r1 && mergeRange.c0 === mergeRange.c1;
    const mergeAction =
      loneCell && alignTab.mergeAction() === 'merge' ? null : alignTab.mergeAction();

    // A multi-area selection has no single merge result. The control is also
    // disabled during hydration, but retain this guard for stale event state.
    if ((state.selection.extraRanges?.length ?? 0) > 0 && mergeAction) return;
    // A refused merge is reported before the data-loss prompt, not after it.
    if (
      mergeAction &&
      !authorizeWrite(store, mergeIntents(mergeRange, dialogEntry.commandId), dialogEntry.origin)
        .allowed
    )
      return;
    // Ask before any format, value, or merge mutation. The no-loss path does
    // not await, preserving the synchronous behavior of existing callers.
    if (mergeAction === 'merge' && !liveWb && mergeHasNonAnchorContent(state, mergeRange)) return;
    if (mergeAction === 'merge' && mergeWillLoseData(state, mergeRange)) {
      waitingForConfirmation = true;
      if (!(await confirmMergeLoseData(strings, state, mergeRange, false, liveWb))) return;
    }

    const touchedPatch = buildTouchedDialogPatch(
      draft,
      touchedFields,
      padsNumberRows(liveWb),
      defaultPatternFor,
      canonicalizeDialogFormula,
    );
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

    // Keep worksheet-wide font-only changes on their compact native font
    // path even when the selection planner can now resolve complete axes.
    // Keep the dialog path narrow: only the compact native font axes may be
    // applied without enumerating the worksheet, and only when no structural
    // merge action is part of the dialog submission. The format snapshot then
    // captures both sparse cell formats and the sheet-level override so one
    // undo restores the exact pre-dialog state.
    const wholeSheetFontOnly =
      mergeAction === null &&
      borderAction === undefined &&
      isWholeSheetSelection(selectedRanges) &&
      Object.keys(actionPatch).length > 0 &&
      Object.keys(actionPatch).every((key) => key === 'fontFamily' || key === 'fontSize');
    if (wholeSheetFontOnly) {
      if (liveWb?.capabilities.sheetFontStyles !== true) return;
      const fontPatch: { fontFamily?: string | null; fontSize?: number | null } = {};
      if (Object.hasOwn(actionPatch, 'fontFamily')) {
        fontPatch.fontFamily = actionPatch.fontFamily ?? null;
      }
      if (Object.hasOwn(actionPatch, 'fontSize')) {
        fontPatch.fontSize = actionPatch.fontSize ?? null;
      }
      const repeat = (): void => {
        const current = store.getState();
        const currentRanges = [current.selection.range, ...(current.selection.extraRanges ?? [])];
        if (!isWholeSheetSelection(currentRanges)) return;
        let applied = false;
        recordFormatChange(
          history,
          store,
          () => {
            applied = setFont(current, store, fontPatch);
          },
          { repeat },
        );
        if (applied) history?.setRepeat(repeat);
      };
      let applied = false;
      recordFormatChange(
        history,
        store,
        () => {
          applied = setFont(state, store, fontPatch);
        },
        { repeat },
      );
      if (applied) {
        history?.setRepeat(repeat);
        api.close();
      }
      return;
    }
    if (!plan) return;

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
      entry: dialogEntry,
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
  const onClear = (): void => {
    const tab = activeTab;
    hydrateFromActive({});
    setActiveTab(tab);
  };

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
    if (isSubmitEnter(e)) {
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
    getDate1904,
    isDxfMode: () => dialogMode === 'dxf',
    macDxfForm: () => macDxf,
    isMacPlatform: () => isMacHost(host),
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
  shell.on(clearBtn, 'click', onClear);
  shell.on(overlay, 'click', (e) => {
    if ((e as MouseEvent).target === overlay) e.stopPropagation();
  });
  shell.on(overlay, 'keydown', onOverlayKey as EventListener);
  shell.on(overlay, 'mousedown', onOverlayPointerDown);

  const api: FormatDialogHandle = {
    open(tab?: TabId, options?: FormatDialogOpenOptions): void {
      pickFormat.begin();
      dialogMode = options?.mode ?? 'format';
      dialogEntry = options?.entry ?? INSTANCE_API_FORMAT_ENTRY;
      macDxf =
        options?.mode === 'dxf' && options.dxfForm && isMacHost(host) ? options.dxfForm : null;
      const onNumberFormat = options?.onApplyNumberFormat;
      applyDxf =
        options?.mode === 'dxf'
          ? (options.onApplyDxf ?? null)
          : options?.mode === 'numberFormat' && onNumberFormat
            ? (format) => {
                if (format.numFmt) onNumberFormat(format.numFmt);
              }
            : null;
      syncColorPaletteTheme(overlay, themeColorColumns(getWb()?.getTheme()?.colors));
      showMoreTab = tab === 'more';
      const remembered = lastUsedTab;
      hydrateFromActive(options?.initialFormat);
      if (options?.mode === 'numberFormat' && options.previewValue) {
        previewValue = options.previewValue;
        renderPreview();
      }
      setDialogMode(options?.mode);
      // The Mac app reopens on the tab used last.
      const requested =
        options?.mode === 'numberFormat'
          ? 'number'
          : (tab ?? (isMacHost(host) && options?.mode !== 'dxf' ? remembered : undefined));
      // The Mac differential-format forms always open on Font.
      if (requested && tabButtons.has(requested) && !macDxf) setActiveTab(requested);
      if (options?.mode === 'dataValidation') setActiveTab('more');
      shell.open();
      if (mixedFields.size > 0) {
        syncMixedControls();
        syncCustomSelects(overlay);
      }
      deferInitialFocus(overlay, () => {
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
