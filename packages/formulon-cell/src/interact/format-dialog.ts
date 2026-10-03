import { coerceInput } from '../commands/coerce-input.js';
import { recordDialogFormatChange } from '../commands/dialog-format-history.js';
import {
  applySelectionFormatAction,
  formatNumber,
  planSelectionFormat,
  type SelectionFormatAction,
  type SelectionFormatPlan,
} from '../commands/format.js';
import { History, recordMergesChangeWithEngine } from '../commands/history.js';
import { interactionControllerFor } from '../commands/interaction-controller.js';
import {
  applyMerge,
  applyUnmerge,
  expandRangeWithMerges,
  mergeAt,
  mergeWillLoseData,
} from '../commands/merge.js';
import { addrKey } from '../engine/address.js';
import type { CellValue, Range } from '../engine/types.js';
import { formatCell, formatGeneralNumber } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import {
  type CellAlign,
  type CellBorderSide,
  type CellFormat,
  type CellVAlign,
  type FillPattern,
  mutators,
  type NegativeStyle,
  type SpreadsheetStore,
  type State,
  type TextDirection,
  type ValidationErrorStyle,
  type ValidationOp,
} from '../store/store.js';
import { appendDialogSelectOptions } from '../toolbar/dialogs/form-controls.js';
import { confirmMergeLoseData } from '../toolbar/dialogs/merge-confirm.js';
import { projectDisabledReason, projectDisabledState } from '../toolbar/menu-a11y.js';
import { formatA1Range } from '../wrappers/toolbar-a1.js';
import { syncCustomSelects } from './custom-select.js';
import { appendDialogOptionButton } from './dialog-shell.js';
import {
  type BorderStyleKey,
  type DraftState,
  defaultCurrencySymbolFor,
  isHexColor,
  type NumberCategory,
  normalizeFormatLocale,
  patternPresetsFor,
  type SideKey,
  type TabId,
  type ValidationKind,
} from './format-dialog-model.js';
import {
  activeDraftSide,
  buildTouchedDialogPatch,
  computeDialogNumFmt,
  computeDialogValidation,
  explicitDraftBorders,
  type FormatDialogField,
  hydrateDraftFromFormat,
  makeEmptyDraft,
  setDraftSide,
  summarizeDialogFormats,
} from './format-dialog-state.js';
import { createFormatDialogView } from './format-dialog-view.js';
import { clampPanelToViewport } from './overlay-position.js';
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

/** A swatch palette that hangs off a color control instead of sitting inline. */
interface PaletteFlyout {
  readonly toggle: HTMLButtonElement;
  setOpen(open: boolean): void;
  isOpen(): boolean;
  /** True when `node` is inside the flyout or its trigger. */
  owns(node: Node): boolean;
}

/** Wire a chevron trigger to a palette flyout.
 *
 * The font and border tabs have no room left for a palette inline, so theirs
 * hang off the color control. The flyout is `position: fixed`, escaping the
 * panel's clip rect, and is placed against the viewport like the grid's own
 * menus. */
function createPaletteFlyout(
  toggle: HTMLButtonElement,
  flyout: HTMLElement,
  palette: { focus(): void },
): PaletteFlyout {
  return {
    toggle,
    setOpen(open: boolean): void {
      flyout.hidden = !open;
      toggle.setAttribute('aria-expanded', open ? 'true' : 'false');
      if (!open) return;
      flyout.style.left = '-9999px';
      flyout.style.top = '-9999px';
      const anchor = toggle.getBoundingClientRect();
      const { x, y } = clampPanelToViewport(flyout, anchor.left, anchor.bottom + 4, { pad: 8 });
      flyout.style.left = `${x}px`;
      flyout.style.top = `${y}px`;
      palette.focus();
    },
    isOpen: () => !flyout.hidden,
    owns: (node: Node) => flyout.contains(node) || toggle.contains(node),
  };
}

type MergeSelectionState = 'none' | 'merged' | 'mixed';

const sameRange = (a: Range, b: Range): boolean =>
  a.sheet === b.sheet && a.r0 === b.r0 && a.c0 === b.c0 && a.r1 === b.r1 && a.c1 === b.c1;

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

/** Convert a spreadsheet date serial to a native `<input type="date">` value
 *  (`yyyy-mm-dd`, UTC). Returns '' for non-finite serials. */
const serialToDateInputValue = (serial: number): string => {
  if (!Number.isFinite(serial)) return '';
  const ms = Math.round((serial - 25569) * 86_400_000);
  const d = new Date(ms);
  const y = String(d.getUTCFullYear()).padStart(4, '0');
  const m = String(d.getUTCMonth() + 1).padStart(2, '0');
  const day = String(d.getUTCDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
};

/** Convert a day-fraction time serial to a native `<input type="time">` value
 *  (`HH:mm`, or `HH:mm:ss` when the serial carries seconds). */
const serialToTimeInputValue = (serial: number): string => {
  if (!Number.isFinite(serial)) return '';
  let total = Math.round((serial % 1) * 86_400);
  total = ((total % 86_400) + 86_400) % 86_400;
  const hh = String(Math.floor(total / 3600)).padStart(2, '0');
  const mm = String(Math.floor((total % 3600) / 60)).padStart(2, '0');
  const ss = total % 60;
  return ss ? `${hh}:${mm}:${String(ss).padStart(2, '0')}` : `${hh}:${mm}`;
};

/** Format a stored bound (serial or plain number) for the bound `<input>` value,
 *  matching the input type chosen for the validation kind. */
const boundInputValue = (kind: ValidationKind, value: number): string => {
  if (kind === 'date') return serialToDateInputValue(value);
  if (kind === 'time') return serialToTimeInputValue(value);
  return String(value);
};

/** Parse a bound `<input>` value back into a stored number. Date/time kinds
 *  route the string through `coerceInput` so `yyyy-mm-dd` / `HH:mm` become the
 *  matching spreadsheet serial; other kinds parse a plain number. Returns null
 *  when the field is empty or unparseable so the previous bound is kept. */
const parseBoundInputValue = (kind: ValidationKind, raw: string): number | null => {
  if (kind === 'date' || kind === 'time') {
    const coerced = coerceInput(raw);
    return coerced.kind === 'number' ? coerced.value : null;
  }
  const n = Number.parseFloat(raw);
  return Number.isFinite(n) ? n : null;
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
    previewCell,
    tabsStrip,
    tabButtons,
    tabPanels,
    catList,
    catDefs,
    catButtons,
    decimalsRow,
    decimalsInput,
    thousandsCk,
    symbolRow,
    symbolSelect,
    patternPresetRow,
    patternPresetSelect,
    patternListWrap,
    patternList,
    patternRow,
    patternInput,
    localeRow,
    localeSelect,
    calendarRow,
    negativeList,
    negativeOptions,
    numberSummaryTitle,
    numberSummaryDesc,
    hAlignRadios,
    hAlignSelect,
    vAlignRadios,
    vAlignSelect,
    wrapCk,
    justifyLastLineCk,
    shrinkCk,
    mergeCk,
    indentInput,
    textDirectionSelect,
    rotationInput,
    alignPreviewDial,
    alignPreviewDialDots,
    alignPreviewDialPointer,
    alignPreviewDialText,
    boldCk,
    italicCk,
    underlineSelect,
    strikeCk,
    superscriptCk,
    subscriptCk,
    normalFontCk,
    fontStyleList,
    familyInput,
    sizeInput,
    colorInput,
    colorReset,
    fontSwatches,
    fontSwatchesToggle,
    fontSwatchesFlyout,
    fontPreviewBox,
    syncFontFamilyOptions,
    borderStyleSelect,
    borderStyleButtons,
    borderStyleGallery,
    borderColorInput,
    borderColorReset,
    borderSwatches,
    borderSwatchesToggle,
    borderSwatchesFlyout,
    presetNone,
    presetOutline,
    presetAll,
    topCk,
    bottomCk,
    leftCk,
    rightCk,
    diagDownCk,
    diagUpCk,
    borderVisualStage,
    borderVisualPreview,
    visualSideButtons,
    fillInput,
    fillReset,
    fillSwatches,
    fillPatternSelect,
    fillPatternGallery,
    fillPatternColorInput,
    fillSample,
    lockedCk,
    hiddenFormulaCk,
    hyperlinkSection,
    commentSection,
    validationSection,
    hlInput,
    hlClear,
    commentArea,
    commentClear,
    validationKindSelect,
    validationOpRow,
    validationOpSelect,
    validationARow,
    validationAInput,
    validationBRow,
    validationBInput,
    validationFormulaRow,
    validationFormulaInput,
    validationListSourceKindRow,
    validationListLiteralRadio,
    validationListRangeRadio,
    validationRow,
    validationArea,
    validationClear,
    validationListRangeRow,
    validationListRangeInput,
    validationShowDropdownRow,
    validationShowDropdownInput,
    validationAllowBlankRow,
    validationAllowBlankInput,
    validationErrorStyleRow,
    validationErrorStyleSelect,
    validationShowInputMessageRow,
    validationShowInputMessageInput,
    validationPromptTitleRow,
    validationPromptTitleInput,
    validationPromptMessageRow,
    validationPromptMessageArea,
    validationShowErrorMessageRow,
    validationShowErrorMessageInput,
    validationErrorTitleRow,
    validationErrorTitleInput,
    validationErrorMessageRow,
    validationErrorMessageArea,
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
  let pendingBorderPreset: 'none' | 'outline' | 'all' | null = null;
  let applyDxf: ((format: Partial<CellFormat>) => void) | null = null;
  let mergeTouched = false;
  let initialMergeChecked = false;
  let initialMergeIndeterminate = false;
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
  const syncFontSizeOptions = (current: number | undefined): void => {
    for (const item of overlay.querySelectorAll<HTMLButtonElement>('[data-fc-font-size]')) {
      item.setAttribute(
        'aria-selected',
        current !== undefined && Number(item.dataset.fcFontSize) === current ? 'true' : 'false',
      );
    }
  };

  // ── Color palette flyouts ──────────────────────────────────────────────
  const fontPalette = createPaletteFlyout(fontSwatchesToggle, fontSwatchesFlyout, fontSwatches);
  const borderPalette = createPaletteFlyout(
    borderSwatchesToggle,
    borderSwatchesFlyout,
    borderSwatches,
  );
  const paletteFlyouts: readonly PaletteFlyout[] = [fontPalette, borderPalette];
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
    pendingBorderPreset = null;

    syncControlsFromDraft();
    const range = state.selection.range;
    const activeMerge = mergeAt(state, state.selection.active);
    const multiCell = range.r0 !== range.r1 || range.c0 !== range.c1;
    const selectedMergeState = mergeSelectionState(state, range);
    mergeCk.input.checked = selectedMergeState === 'merged';
    mergeCk.input.indeterminate = selectedMergeState === 'mixed';
    mergeTouched = false;
    initialMergeChecked = mergeCk.input.checked;
    initialMergeIndeterminate = mergeCk.input.indeterminate;
    const mergeRange = expandRangeWithMerges(state, range);
    const hasExtraRanges = (state.selection.extraRanges?.length ?? 0) > 0;
    const mergeDisabled =
      hasExtraRanges ||
      (!multiCell && activeMerge === null) ||
      (!getWb() && mergeHasNonAnchorContent(state, mergeRange));
    const mergeReason = mergeDisabled ? strings.formatDialog.mergeCellsRequiresMultiCell : null;
    projectDisabledState(mergeCk.input, mergeDisabled, mergeReason, {
      datasetKey: 'disabledReason',
      titlePrefix: strings.formatDialog.mergeCells,
    });
    projectDisabledReason(mergeCk.wrap, mergeReason, {
      datasetKey: 'disabledReason',
      titlePrefix: strings.formatDialog.mergeCells,
    });
    mergeCk.wrap.classList.toggle('fc-fmtdlg__check--muted', mergeCk.input.disabled);
    renderPreview();
    setActiveTab('number');
  };

  const underlineMixedOption = ((): HTMLOptionElement => {
    const scratch = document.createElement('select');
    appendDialogSelectOptions(scratch, [{ value: 'mixed', label: '' }]);
    const option = scratch.options[0] as HTMLOptionElement;
    projectDisabledState(option, true, null);
    return option;
  })();

  const syncMixedControls = (): void => {
    const isMixed = isMixedField;
    const mark = (element: HTMLElement, mixed: boolean, field?: FormatDialogField): void => {
      if (field !== undefined) element.dataset.fcMixedField = field;
      if (mixed) element.dataset.fcMixed = 'true';
      else delete element.dataset.fcMixed;
    };
    const setMixedCheck = (input: HTMLInputElement, field: FormatDialogField): void => {
      input.indeterminate = isMixed(field);
      mark(input, input.indeterminate, field);
    };
    setMixedCheck(thousandsCk.input, 'numFmt');
    setMixedCheck(wrapCk.input, 'wrap');
    setMixedCheck(justifyLastLineCk.input, 'justifyLastLine');
    setMixedCheck(shrinkCk.input, 'shrinkToFit');
    setMixedCheck(boldCk.input, 'bold');
    setMixedCheck(italicCk.input, 'italic');
    setMixedCheck(strikeCk.input, 'strike');
    setMixedCheck(superscriptCk.input, 'fontVertAlign');
    setMixedCheck(subscriptCk.input, 'fontVertAlign');
    const normalFontMixed = [
      'bold',
      'italic',
      'underline',
      'strike',
      'fontVertAlign',
      'fontFamily',
      'fontSize',
      'color',
    ].some((field) => isMixed(field as FormatDialogField));
    normalFontCk.input.indeterminate = normalFontMixed;
    mark(normalFontCk.input, normalFontMixed);
    setMixedCheck(lockedCk.input, 'locked');
    setMixedCheck(hiddenFormulaCk.input, 'formulaHidden');

    if (isMixed('numFmt')) {
      for (const button of catButtons.values()) button.setAttribute('aria-selected', 'false');
      mark(catList, true, 'numFmt');
      decimalsInput.value = '';
      mark(decimalsInput, true, 'numFmt');
      symbolSelect.value = '';
      symbolSelect.selectedIndex = -1;
      mark(symbolSelect, true, 'numFmt');
      for (const button of negativeOptions.querySelectorAll<HTMLButtonElement>(
        '[data-fc-negative-style]',
      )) {
        button.setAttribute('aria-selected', 'false');
      }
      mark(negativeOptions, true, 'numFmt');
      patternInput.value = '';
      patternPresetSelect.value = '';
      for (const button of patternList.querySelectorAll<HTMLButtonElement>('[data-fc-pattern]')) {
        button.setAttribute('aria-selected', 'false');
      }
      mark(patternList, true, 'numFmt');
      mark(patternInput, true, 'numFmt');
      mark(patternPresetSelect, true, 'numFmt');
    }
    if (isMixed('align')) {
      for (const radio of hAlignRadios.values()) radio.checked = false;
      hAlignSelect.value = '';
      mark(hAlignSelect, true, 'align');
    }
    if (isMixed('vAlign')) {
      for (const radio of vAlignRadios.values()) radio.checked = false;
      vAlignSelect.value = '';
      mark(vAlignSelect, true, 'vAlign');
    }
    if (isMixed('underline')) {
      // A placeholder keeps the real "None" option distinct from the mixed state,
      // so choosing None still reports a change.
      underlineSelect.append(underlineMixedOption);
      underlineSelect.value = underlineMixedOption.value;
      mark(underlineSelect, true, 'underline');
    }
    if (isMixed('fontFamily')) {
      familyInput.value = '';
      mark(familyInput, true, 'fontFamily');
    }
    if (isMixed('fontSize')) {
      sizeInput.value = '';
      mark(sizeInput, true, 'fontSize');
    }
    syncFontStyleList();
    syncFontFamilyOptions(isMixed('fontFamily') ? '' : draft.fontFamily, !isMixed('fontFamily'));
    syncFontSizeOptions(isMixed('fontSize') ? undefined : draft.fontSize);
    if (isMixed('color')) {
      fontSwatches.setValue(null);
      mark(colorInput, true, 'color');
    }
    if (isMixed('fill')) {
      fillSwatches.setValue(null);
      mark(fillInput, true, 'fill');
    }
    if (isMixed('fillPattern')) {
      fillPatternSelect.value = '';
      for (const button of fillPatternGallery.querySelectorAll<HTMLButtonElement>(
        '[data-fc-fill-pattern]',
      )) {
        button.setAttribute('aria-pressed', 'false');
      }
      mark(fillPatternGallery, true, 'fillPattern');
      mark(fillPatternSelect, true, 'fillPattern');
    }
    if (isMixed('fillPatternColor')) {
      mark(fillPatternColorInput, true, 'fillPatternColor');
    }
    if (isMixed('hyperlink')) {
      hlInput.value = '';
      mark(hlInput, true, 'hyperlink');
    }
    if (isMixed('comment')) {
      commentArea.value = '';
      mark(commentArea, true, 'comment');
    }
    if (isMixed('validation')) {
      validationKindSelect.value = '';
      validationOpSelect.value = '';
      validationAInput.value = '';
      validationBInput.value = '';
      validationFormulaInput.value = '';
      validationListSourceKindRow.querySelectorAll<HTMLInputElement>('input').forEach((input) => {
        input.checked = false;
        input.indeterminate = false;
      });
      validationArea.value = '';
      validationListRangeInput.value = '';
      validationShowDropdownInput.checked = false;
      validationShowDropdownInput.indeterminate = true;
      validationAllowBlankInput.checked = false;
      validationAllowBlankInput.indeterminate = true;
      validationErrorStyleSelect.value = '';
      validationShowInputMessageInput.checked = false;
      validationShowInputMessageInput.indeterminate = true;
      validationPromptTitleInput.value = '';
      validationPromptMessageArea.value = '';
      validationShowErrorMessageInput.checked = false;
      validationShowErrorMessageInput.indeterminate = true;
      validationErrorTitleInput.value = '';
      validationErrorMessageArea.value = '';
      for (const control of [
        validationKindSelect,
        validationOpSelect,
        validationAInput,
        validationBInput,
        validationFormulaInput,
        validationArea,
        validationListRangeInput,
        validationShowDropdownInput,
        validationAllowBlankInput,
        validationErrorStyleSelect,
        validationShowInputMessageInput,
        validationPromptTitleInput,
        validationPromptMessageArea,
        validationShowErrorMessageInput,
        validationErrorTitleInput,
        validationErrorMessageArea,
      ]) {
        mark(control, true, 'validation');
      }
      validationListSourceKindRow.querySelectorAll<HTMLInputElement>('input').forEach((input) => {
        mark(input, true, 'validation');
      });
      mark(validationKindSelect, true, 'validation');
    }
    for (const [field, input] of [
      ['indent', indentInput],
      ['rotation', rotationInput],
      ['textDirection', textDirectionSelect],
    ] as const) {
      if (isMixed(field)) {
        input.value = '';
        mark(input, true, field);
      }
    }
    for (const side of ['top', 'right', 'bottom', 'left', 'diagonalDown', 'diagonalUp'] as const) {
      const field = `border.${side}` as FormatDialogField;
      if (isMixed(field)) {
        const control =
          side === 'top'
            ? topCk.input
            : side === 'right'
              ? rightCk.input
              : side === 'bottom'
                ? bottomCk.input
                : side === 'left'
                  ? leftCk.input
                  : side === 'diagonalDown'
                    ? diagDownCk.input
                    : diagUpCk.input;
        control.indeterminate = true;
        mark(control, true, field);
      }
    }
  };

  const syncJustifyLastLineAvailability = (): void => {
    projectDisabledState(justifyLastLineCk.input, draft.align !== 'distributed', null);
    justifyLastLineCk.wrap.classList.toggle(
      'fc-fmtdlg__check--muted',
      justifyLastLineCk.input.disabled,
    );
  };

  const syncControlsFromDraft = (): void => {
    // Number
    for (const [id, btn] of catButtons) {
      btn.setAttribute('aria-selected', id === draft.numberCategory ? 'true' : 'false');
      btn.tabIndex = id === draft.numberCategory ? 0 : -1;
    }
    decimalsInput.value = String(draft.decimals);
    thousandsCk.input.checked = draft.thousands;
    for (const item of negativeOptions.querySelectorAll<HTMLButtonElement>(
      '[data-fc-negative-style]',
    )) {
      item.setAttribute(
        'aria-selected',
        item.dataset.fcNegativeStyle === draft.negativeStyle ? 'true' : 'false',
      );
    }
    symbolSelect.value = draft.currencySymbol;
    patternInput.value = draft.pattern;
    if (!draft.pattern) {
      patternInput.placeholder = defaultPatternFor(draft.numberCategory) || t.patternPlaceholder;
    } else {
      patternInput.placeholder = t.patternPlaceholder;
    }
    syncPatternPresetOptions();
    syncNumberControlsVisibility();

    // Alignment
    const hKey: 'default' | CellAlign = draft.align ?? 'default';
    for (const [id, r] of hAlignRadios) r.checked = id === hKey;
    hAlignSelect.value = hKey;
    const vKey: 'default' | CellVAlign = draft.vAlign ?? 'default';
    for (const [id, r] of vAlignRadios) r.checked = id === vKey;
    vAlignSelect.value = vKey;
    wrapCk.input.checked = draft.wrap;
    justifyLastLineCk.input.checked = draft.justifyLastLine;
    syncJustifyLastLineAvailability();
    shrinkCk.input.checked = draft.shrinkToFit;
    indentInput.value = String(draft.indent);
    textDirectionSelect.value = draft.textDirection;
    rotationInput.value = String(draft.rotation);
    syncRotationDial(draft.rotation);

    // Font
    boldCk.input.checked = draft.bold;
    italicCk.input.checked = draft.italic;
    underlineMixedOption.remove();
    underlineSelect.value = draft.underline === true ? 'single' : draft.underline || '';
    strikeCk.input.checked = draft.strike;
    superscriptCk.input.checked = draft.fontVertAlign === 'superscript';
    subscriptCk.input.checked = draft.fontVertAlign === 'subscript';
    normalFontCk.input.checked =
      !draft.bold &&
      !draft.italic &&
      !draft.underline &&
      !draft.strike &&
      !draft.fontVertAlign &&
      !draft.fontFamily &&
      draft.fontSize === undefined &&
      draft.color === undefined;
    syncFontStyleList();
    familyInput.value = draft.fontFamily;
    syncFontFamilyOptions(draft.fontFamily);
    sizeInput.value = draft.fontSize !== undefined ? String(draft.fontSize) : '';
    syncFontSizeOptions(draft.fontSize);
    colorInput.value = draft.color && isHexColor(draft.color) ? draft.color : '#000000';
    fontSwatches.setValue(draft.color && isHexColor(draft.color) ? draft.color : null);

    // Borders
    borderStyleSelect.value = draft.borderStyle;
    for (const [id, btn] of borderStyleButtons) {
      btn.setAttribute('aria-pressed', id === draft.borderStyle ? 'true' : 'false');
    }
    borderColorInput.value =
      draft.borderColor && isHexColor(draft.borderColor) ? draft.borderColor : '#000000';
    borderSwatches.setValue(
      draft.borderColor && isHexColor(draft.borderColor) ? draft.borderColor : null,
    );
    topCk.input.checked = !!draft.borders.top;
    bottomCk.input.checked = !!draft.borders.bottom;
    leftCk.input.checked = !!draft.borders.left;
    rightCk.input.checked = !!draft.borders.right;
    diagDownCk.input.checked = !!draft.borders.diagonalDown;
    diagUpCk.input.checked = !!draft.borders.diagonalUp;

    // Fill
    fillInput.value = draft.fill && isHexColor(draft.fill) ? draft.fill : '#ffffff';
    fillSwatches.setValue(draft.fill && isHexColor(draft.fill) ? draft.fill : null);
    fillPatternSelect.value = draft.fillPattern ?? '';
    for (const button of fillPatternGallery.querySelectorAll<HTMLButtonElement>(
      '[data-fc-fill-pattern]',
    )) {
      button.setAttribute(
        'aria-pressed',
        button.dataset.fcFillPattern === (draft.fillPattern ?? '') ? 'true' : 'false',
      );
    }
    fillPatternColorInput.value =
      draft.fillPatternColor && isHexColor(draft.fillPatternColor)
        ? draft.fillPatternColor
        : '#000000';

    // Protection
    lockedCk.input.checked = draft.locked;
    hiddenFormulaCk.input.checked = draft.formulaHidden;

    // More
    hlInput.value = draft.hyperlink;
    commentArea.value = draft.comment;
    validationArea.value = draft.validationList;
    validationListRangeInput.value = draft.validationListRange;
    validationListLiteralRadio.input.checked = draft.validationListSourceKind === 'literal';
    validationListRangeRadio.input.checked = draft.validationListSourceKind === 'range';
    validationShowDropdownInput.checked = draft.validationShowDropdown;
    validationKindSelect.value = draft.validationKind;
    validationOpSelect.value = draft.validationOp;
    applyBoundInputMode(draft.validationKind);
    validationAInput.value = boundInputValue(draft.validationKind, draft.validationA);
    validationBInput.value = boundInputValue(draft.validationKind, draft.validationB);
    validationFormulaInput.value = draft.validationFormula;
    validationAllowBlankInput.checked = draft.validationAllowBlank;
    validationErrorStyleSelect.value = draft.validationErrorStyle;
    validationShowInputMessageInput.checked = draft.validationShowInputMessage;
    validationPromptTitleInput.value = draft.validationPromptTitle;
    validationPromptMessageArea.value = draft.validationPromptMessage;
    validationShowErrorMessageInput.checked = draft.validationShowErrorMessage;
    validationErrorTitleInput.value = draft.validationErrorTitle;
    validationErrorMessageArea.value = draft.validationErrorMessage;
    syncValidationVisibility();
    syncMixedControls();
  };

  /** Switch the A/B bound inputs to a native date/time picker for date/time
   *  validation (so bounds are pickable instead of raw serials) and back to a
   *  number field otherwise. */
  const applyBoundInputMode = (kind: ValidationKind): void => {
    const type = kind === 'date' ? 'date' : kind === 'time' ? 'time' : 'number';
    for (const input of [validationAInput, validationBInput]) {
      if (input.type !== type) input.type = type;
      if (type === 'time') input.step = '1';
      else if (type === 'number') input.step = 'any';
      else input.removeAttribute('step');
    }
  };

  const syncValidationVisibility = (): void => {
    const k = draft.validationKind;
    applyBoundInputMode(k);
    const isBounded =
      k === 'whole' || k === 'decimal' || k === 'date' || k === 'time' || k === 'textLength';
    const isListLike = k === 'list';
    const isCustom = k === 'custom';
    const isActive = k !== 'none';
    validationOpRow.hidden = !isBounded;
    validationARow.hidden = !isBounded;
    validationBRow.hidden =
      !isBounded || (draft.validationOp !== 'between' && draft.validationOp !== 'notBetween');
    validationFormulaRow.hidden = !isCustom;
    validationListSourceKindRow.hidden = !isListLike;
    validationRow.hidden = !isListLike || draft.validationListSourceKind !== 'literal';
    validationListRangeRow.hidden = !isListLike || draft.validationListSourceKind !== 'range';
    validationShowDropdownRow.hidden = !isListLike;
    validationAllowBlankRow.hidden = !isActive;
    validationErrorStyleRow.hidden = !isActive;
    validationShowInputMessageRow.hidden = !isActive;
    validationPromptTitleRow.hidden = !isActive || !draft.validationShowInputMessage;
    validationPromptMessageRow.hidden = !isActive || !draft.validationShowInputMessage;
    validationShowErrorMessageRow.hidden = !isActive;
    validationErrorTitleRow.hidden = !isActive || !draft.validationShowErrorMessage;
    validationErrorMessageRow.hidden = !isActive || !draft.validationShowErrorMessage;
  };

  const syncNumberControlsVisibility = (): void => {
    const cat = draft.numberCategory;
    tabPanels.get('number')?.setAttribute('data-number-category', cat);
    const decimalsCats = new Set<NumberCategory>([
      'fixed',
      'currency',
      'percent',
      'scientific',
      'accounting',
    ]);
    const symbolCats = new Set<NumberCategory>(['currency', 'accounting']);
    const listboxCats = new Set<NumberCategory>(['date', 'time', 'fraction', 'special']);
    decimalsRow.hidden = !decimalsCats.has(cat);
    thousandsCk.wrap.hidden = cat !== 'fixed';
    symbolRow.hidden = !symbolCats.has(cat);
    // For date/time-like categories use the Office365-style clickable
    // listbox; only the custom category keeps the dropdown of code-style
    // presets.
    patternPresetRow.hidden = cat !== 'custom';
    patternListWrap.hidden = !listboxCats.has(cat);
    patternRow.hidden = cat !== 'custom';
    localeRow.hidden = cat !== 'date' && cat !== 'time' && cat !== 'special';
    localeSelect.value = normalizeFormatLocale(getFormatLocale()).startsWith('ja') ? 'ja' : 'en';
    calendarRow.hidden = cat !== 'date';
    negativeList.hidden = cat !== 'fixed' && cat !== 'currency';
    const active = catDefs.find((c) => c.id === cat);
    numberSummaryTitle.textContent = active?.label ?? t.catGeneral;
    // Description moved to the hint bar; keep the in-controls slot empty so
    // it does not push the layout.
    numberSummaryDesc.textContent = '';
    syncNegativeSamples();
    syncHintBar();
  };

  const syncNegativeSamples = (): void => {
    const cat = draft.numberCategory;
    const items = negativeOptions.querySelectorAll<HTMLButtonElement>('[data-fc-negative-style]');
    for (const item of items) {
      const style = item.dataset.fcNegativeStyle as NegativeStyle | undefined;
      if (!style) continue;
      const sampleFmt = computeDialogNumFmt(
        { ...draft, numberCategory: cat, negativeStyle: style },
        defaultPatternFor,
      );
      item.textContent = formatNumber(-1234, sampleFmt, getFormatLocale());
    }
  };

  const currentFontStyleId = (): 'regular' | 'italic' | 'bold' | 'boldItalic' => {
    if (draft.bold && draft.italic) return 'boldItalic';
    if (draft.bold) return 'bold';
    if (draft.italic) return 'italic';
    return 'regular';
  };

  const syncFontStyleList = (): void => {
    const id = isMixedField('bold') || isMixedField('italic') ? null : currentFontStyleId();
    for (const item of fontStyleList.querySelectorAll<HTMLButtonElement>('[data-fc-font-style]')) {
      item.setAttribute(
        'aria-selected',
        id !== null && item.dataset.fcFontStyle === id ? 'true' : 'false',
      );
    }
  };

  const fillPatternImage = (pattern: FillPattern | undefined, color = '#000000'): string => {
    switch (pattern) {
      case 'gray0625':
        return `radial-gradient(${color} 0.4px, transparent 0.4px)`;
      case 'gray125':
        return `radial-gradient(${color} 0.6px, transparent 0.6px)`;
      case 'gray25':
        return `radial-gradient(${color} 1px, transparent 1px)`;
      case 'gray50':
        return `repeating-linear-gradient(45deg, ${color} 0 2px, transparent 2px 4px)`;
      case 'gray75':
        return `repeating-linear-gradient(45deg, ${color} 0 3px, transparent 3px 4px)`;
      case 'horizontal':
      case 'darkHorizontal':
        return `repeating-linear-gradient(0deg, ${color} 0 1px, transparent 1px 4px)`;
      case 'lightHorizontal':
        return `repeating-linear-gradient(0deg, ${color} 0 1px, transparent 1px 7px)`;
      case 'vertical':
      case 'darkVertical':
        return `repeating-linear-gradient(90deg, ${color} 0 1px, transparent 1px 4px)`;
      case 'lightVertical':
        return `repeating-linear-gradient(90deg, ${color} 0 1px, transparent 1px 7px)`;
      case 'diagonalDown':
      case 'darkDown':
        return `repeating-linear-gradient(45deg, ${color} 0 1px, transparent 1px 5px)`;
      case 'diagonalUp':
      case 'darkUp':
        return `repeating-linear-gradient(135deg, ${color} 0 1px, transparent 1px 5px)`;
      case 'lightDown':
        return `repeating-linear-gradient(45deg, ${color} 0 1px, transparent 1px 9px)`;
      case 'lightUp':
        return `repeating-linear-gradient(135deg, ${color} 0 1px, transparent 1px 9px)`;
      case 'darkGrid':
      case 'lightGrid': {
        const step = pattern === 'darkGrid' ? 4 : 7;
        return `repeating-linear-gradient(0deg, ${color} 0 1px, transparent 1px ${step}px), repeating-linear-gradient(90deg, ${color} 0 1px, transparent 1px ${step}px)`;
      }
      case 'darkTrellis':
      case 'lightTrellis': {
        const step = pattern === 'darkTrellis' ? 6 : 9;
        return `repeating-linear-gradient(45deg, ${color} 0 1px, transparent 1px ${step}px), repeating-linear-gradient(135deg, ${color} 0 1px, transparent 1px ${step}px)`;
      }
      default:
        return '';
    }
  };

  // ── Border helpers ─────────────────────────────────────────────────────
  const activeSide = (): CellBorderSide => activeDraftSide(draft);
  const setSide = (key: SideKey, on: boolean): void => {
    draft.borders = setDraftSide(draft, key, on);
  };
  // ── Preview rendering ──────────────────────────────────────────────────
  const cssHorizontalAlign = (align: CellAlign | undefined): CSSStyleDeclaration['textAlign'] => {
    switch (align) {
      case 'center':
      case 'centerContinuous':
        return 'center';
      case 'right':
        return 'right';
      case 'justify':
      case 'distributed':
        return 'justify';
      default:
        return 'left';
    }
  };

  const cssVerticalJustify = (
    align: CellVAlign | undefined,
  ): CSSStyleDeclaration['justifyContent'] => {
    switch (align) {
      case 'top':
        return 'flex-start';
      case 'bottom':
        return 'flex-end';
      default:
        return 'center';
    }
  };

  /** Move the dial's marker onto the selected angle and tilt its sample text to
   *  match. Driven from `renderPreview` so typing a degree, clicking a dot and
   *  reopening the dialog all land on the same dial state. */
  const syncRotationDial = (rotation: number): void => {
    for (const dot of alignPreviewDialDots) {
      const angle = Number.parseInt(dot.dataset.fcAngle ?? '0', 10);
      const active = angle === rotation;
      dot.classList.toggle('fc-fmtdlg__align-preview-dot--active', active);
      dot.setAttribute('aria-pressed', active ? 'true' : 'false');
    }
    const rad = (rotation * Math.PI) / 180;
    const cx = 12;
    const cy = 66;
    const radius = 56;
    const px = cx + radius * Math.cos(rad);
    const py = cy - radius * Math.sin(rad);
    alignPreviewDialPointer.style.left = `${px}px`;
    alignPreviewDialPointer.style.top = `${py}px`;
    alignPreviewDialText.style.transform = `translate(0, -50%) rotate(${-rotation}deg)`;
  };

  const renderPreview = (): void => {
    syncRotationDial(draft.rotation);
    const cssFontVertAlign =
      draft.fontVertAlign === 'superscript'
        ? 'super'
        : draft.fontVertAlign === 'subscript'
          ? 'sub'
          : '';
    const applyFontPreview = (el: HTMLElement): void => {
      el.style.fontWeight = draft.bold ? 'bold' : 'normal';
      el.style.fontStyle = draft.italic ? 'italic' : 'normal';
      const decos: string[] = [];
      if (draft.underline) decos.push('underline');
      if (draft.strike) decos.push('line-through');
      el.style.textDecoration = decos.length > 0 ? decos.join(' ') : 'none';
      el.style.textDecorationStyle =
        draft.underline === 'double' || draft.underline === 'doubleAccounting' ? 'double' : '';
      el.style.fontFamily = draft.fontFamily || '';
      el.style.fontSize = draft.fontSize !== undefined ? `${draft.fontSize}px` : '';
      el.style.verticalAlign = cssFontVertAlign;
      el.style.color = draft.color ?? '';
    };
    const applyFillPreview = (el: HTMLElement): void => {
      el.style.backgroundColor = draft.fill ?? '';
      el.style.backgroundImage = fillPatternImage(draft.fillPattern, draft.fillPatternColor);
      el.style.backgroundSize =
        draft.fillPattern === 'gray125' || draft.fillPattern === 'gray25' ? '4px 4px' : '';
    };

    preview.style.fontWeight = draft.bold ? 'bold' : 'normal';
    preview.style.fontStyle = draft.italic ? 'italic' : 'normal';
    const decos: string[] = [];
    if (draft.underline) decos.push('underline');
    if (draft.strike) decos.push('line-through');
    preview.style.textDecoration = decos.length > 0 ? decos.join(' ') : 'none';
    preview.style.textDecorationStyle =
      draft.underline === 'double' || draft.underline === 'doubleAccounting' ? 'double' : '';
    preview.style.verticalAlign = cssFontVertAlign;
    preview.style.textAlign = cssHorizontalAlign(draft.align);
    applyFontPreview(previewCell);
    applyFontPreview(fontPreviewBox);
    previewCell.style.textAlign = cssHorizontalAlign(draft.align);
    previewCell.style.direction = draft.textDirection === 'context' ? '' : draft.textDirection;
    applyFillPreview(previewCell);
    applyFillPreview(fillSample);
    previewCell.style.whiteSpace = draft.wrap ? 'pre-wrap' : 'nowrap';
    previewCell.style.fontSize = draft.shrinkToFit
      ? `${Math.max(8, Math.round((draft.fontSize ?? 13) * 0.85))}px`
      : draft.fontSize !== undefined
        ? `${draft.fontSize}px`
        : '';
    previewCell.style.justifyContent = cssVerticalJustify(draft.vAlign);
    const cssBorder = (s: CellBorderSide | undefined): string => {
      if (!s) return '0 solid transparent';
      const cfg = typeof s === 'object' ? s : { style: 'thin' as const };
      const widthPx = cfg.style === 'thick' ? 3 : cfg.style === 'medium' ? 2 : 1;
      const cssStyle =
        cfg.style === 'dashed'
          ? 'dashed'
          : cfg.style === 'dotted'
            ? 'dotted'
            : cfg.style === 'double'
              ? 'double'
              : 'solid';
      const cssColor = (typeof s === 'object' && s.color) || 'currentColor';
      const w = cfg.style === 'double' ? Math.max(widthPx, 3) : widthPx;
      return `${w}px ${cssStyle} ${cssColor}`;
    };
    previewCell.style.borderTop = cssBorder(draft.borders.top);
    previewCell.style.borderRight = cssBorder(draft.borders.right);
    previewCell.style.borderBottom = cssBorder(draft.borders.bottom);
    previewCell.style.borderLeft = cssBorder(draft.borders.left);
    borderVisualPreview.style.borderTop = cssBorder(draft.borders.top);
    borderVisualPreview.style.borderRight = cssBorder(draft.borders.right);
    borderVisualPreview.style.borderBottom = cssBorder(draft.borders.bottom);
    borderVisualPreview.style.borderLeft = cssBorder(draft.borders.left);
    borderVisualPreview.classList.toggle(
      'fc-fmtdlg__border-preview--diag-down',
      !!draft.borders.diagonalDown,
    );
    borderVisualPreview.classList.toggle(
      'fc-fmtdlg__border-preview--diag-up',
      !!draft.borders.diagonalUp,
    );
    const diagSide = draft.borders.diagonalDown || draft.borders.diagonalUp || activeSide();
    const diagCfg = typeof diagSide === 'object' ? diagSide : { style: 'thin' as const };
    const diagColor = (typeof diagSide === 'object' && diagSide.color) || 'currentColor';
    const diagWidth = diagCfg.style === 'thick' ? 3 : diagCfg.style === 'medium' ? 2 : 1;
    borderVisualPreview.style.setProperty('--fc-fmtdlg-border-diag-color', diagColor);
    borderVisualPreview.style.setProperty('--fc-fmtdlg-border-diag-width', `${diagWidth}px`);
    for (const [key, buttons] of visualSideButtons) {
      for (const btn of buttons) {
        btn.setAttribute('aria-pressed', draft.borders[key] ? 'true' : 'false');
      }
    }

    const numFmt = computeDialogNumFmt(draft, defaultPatternFor);
    // Differential-format editing has no cell value to preview, so retain a
    // representative sample there. Normal Format Cells previews the active
    // cell, including text, booleans, errors, and blanks.
    const isDateLike =
      numFmt.kind === 'date' || numFmt.kind === 'time' || numFmt.kind === 'datetime';
    const syntheticSampleValue =
      draft.numberCategory === 'fraction'
        ? 1.25
        : (draft.numberCategory === 'fixed' || draft.numberCategory === 'currency') &&
            draft.negativeStyle !== 'minus'
          ? -1234
          : isDateLike || draft.numberCategory === 'currency' || draft.numberCategory === 'special'
            ? 10
            : 12345;
    const value =
      applyDxf !== null ? { kind: 'number' as const, value: syntheticSampleValue } : previewValue;
    const numericText =
      value.kind === 'number'
        ? applyDxf !== null || numFmt.kind !== 'general'
          ? formatNumber(value.value, numFmt, getFormatLocale())
          : formatGeneralNumber(value.value, getFormatLocale(), { useGrouping: false })
        : formatCell(value, getFormatLocale());
    previewCell.textContent = numericText;
    if (!draft.color && value.kind === 'number' && value.value < 0) {
      previewCell.style.color =
        draft.negativeStyle === 'red' || draft.negativeStyle === 'red-parens' ? '#c00000' : '';
    }
  };

  // ── Compute helpers ────────────────────────────────────────────────────
  const defaultPatternFor = (cat: NumberCategory): string => {
    const presets =
      cat === 'date' ||
      cat === 'time' ||
      cat === 'fraction' ||
      cat === 'special' ||
      cat === 'custom'
        ? patternPresetsFor(getFormatLocale())[cat]
        : [];
    if (presets[0]) return presets[0];
    switch (cat) {
      case 'date':
        return 'yyyy-mm-dd';
      case 'time':
        return 'HH:MM:SS';
      case 'fraction':
        return '# ?/?';
      case 'special':
        return '000';
      case 'custom':
        return '0.00';
      default:
        return '';
    }
  };

  const syncPatternPresetOptions = (): void => {
    const cat = draft.numberCategory;
    const specialLabels = t.specialFormatLabels.split('\n');
    const patterns =
      cat === 'date' ||
      cat === 'time' ||
      cat === 'fraction' ||
      cat === 'special' ||
      cat === 'custom'
        ? [...patternPresetsFor(getFormatLocale())[cat]]
        : [];
    const current = draft.pattern || defaultPatternFor(cat);
    if (current && !patterns.includes(current)) patterns.unshift(current);
    patternPresetSelect.replaceChildren();
    appendDialogSelectOptions(
      patternPresetSelect,
      patterns.map((pattern, index) => ({
        value: pattern,
        label: cat === 'special' ? (specialLabels[index] ?? pattern) : pattern,
      })),
    );
    patternPresetSelect.value = current;
    syncPatternListItems(patterns, current, specialLabels);
  };

  // Pattern preview values per category — pick a value that exercises the
  // formatting rules so users see day-of-week, AM/PM, etc.
  const patternSampleValue = (cat: NumberCategory): number => {
    switch (cat) {
      case 'date':
        return 41348.5625; // 2013-03-14 13:30
      case 'fraction':
        return 1.25;
      case 'time':
        return 0.5625; // 13:30:00
      case 'special':
        return 12345;
      default:
        return 12345;
    }
  };

  const syncPatternListItems = (
    patterns: string[],
    current: string,
    specialLabels: string[],
  ): void => {
    const cat = draft.numberCategory;
    const isListbox = cat === 'date' || cat === 'time' || cat === 'fraction' || cat === 'special';
    if (!isListbox) {
      patternList.replaceChildren();
      return;
    }
    const sample = patternSampleValue(cat);
    const locale = getFormatLocale();
    patternList.replaceChildren();
    for (const [index, pattern] of patterns.entries()) {
      let label = '';
      if (cat === 'special') {
        label = specialLabels[index] ?? pattern;
      } else {
        try {
          label = formatNumber(
            sample,
            cat === 'date'
              ? { kind: 'date', pattern }
              : cat === 'time'
                ? { kind: 'time', pattern }
                : { kind: 'custom', pattern },
            locale,
          );
        } catch {
          label = pattern;
        }
        label = label || pattern;
      }
      appendDialogOptionButton(patternList, {
        label,
        baseClass: 'fc-fmtdlg__pattern-item',
        datasetKey: 'fcPattern',
        value: pattern,
        selected: pattern === current,
      });
    }
  };

  const numberCategoryDescription = (cat: NumberCategory): string => {
    switch (cat) {
      case 'fixed':
        return t.descFixed;
      case 'currency':
        return t.descCurrency;
      case 'accounting':
        return t.descAccounting;
      case 'percent':
        return t.descPercent;
      case 'scientific':
        return t.descScientific;
      case 'date':
        return t.descDate;
      case 'time':
        return t.descTime;
      case 'fraction':
        return t.descFraction;
      case 'text':
        return t.descText;
      case 'special':
        return t.descOther;
      case 'custom':
        return t.descCustom;
      default:
        return t.descGeneral;
    }
  };
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
      hintBar.textContent = numberCategoryDescription(draft.numberCategory);
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
    const state = store.getState();
    const range = state.selection.range;

    const validationLines = draft.validationList
      .split(/\r?\n/)
      .map((s) => s.trim())
      .filter((s) => s.length > 0);

    const explicitBorders = explicitDraftBorders(draft);

    const validation = computeDialogValidation(draft, validationLines);
    const hyperlink = draft.hyperlink.trim();
    const preserveHyperlinkMetadata = hyperlink.length > 0 && hyperlink === draft.originalHyperlink;

    const patch: Partial<CellFormat> = {
      numFmt: computeDialogNumFmt(draft, defaultPatternFor),
      align: draft.align,
      vAlign: draft.vAlign,
      wrap: draft.wrap,
      justifyLastLine: draft.align === 'distributed' && draft.justifyLastLine,
      shrinkToFit: draft.shrinkToFit,
      indent: draft.indent > 0 ? draft.indent : undefined,
      rotation: draft.rotation !== 0 ? draft.rotation : undefined,
      textDirection: draft.textDirection !== 'context' ? draft.textDirection : undefined,
      bold: draft.bold,
      italic: draft.italic,
      underline: draft.underline,
      strike: draft.strike,
      fontVertAlign: draft.fontVertAlign,
      fontFamily: draft.fontFamily ? draft.fontFamily : undefined,
      fontSize: draft.fontSize,
      color: draft.color,
      fill: draft.fill,
      fillPattern: draft.fillPattern,
      fillPatternColor: draft.fillPattern ? draft.fillPatternColor : undefined,
      borders: explicitBorders,
      hyperlink: hyperlink ? hyperlink : undefined,
      hyperlinkDisplay: preserveHyperlinkMetadata ? draft.hyperlinkDisplay : undefined,
      hyperlinkTooltip: preserveHyperlinkMetadata ? draft.hyperlinkTooltip : undefined,
      comment: draft.comment ? draft.comment : undefined,
      ...(draft.comment ? {} : { commentAuthor: undefined }),
      validation,
      locked: draft.locked,
      formulaHidden: draft.formulaHidden ? true : undefined,
    };

    if (applyDxf) {
      applyDxf({
        numFmt: patch.numFmt,
        bold: patch.bold,
        italic: patch.italic,
        underline: patch.underline,
        strike: patch.strike,
        fontVertAlign: patch.fontVertAlign,
        fontFamily: patch.fontFamily,
        fontSize: patch.fontSize,
        color: patch.color,
        fill: patch.fill,
        fillPattern: patch.fillPattern,
        fillPatternColor: patch.fillPatternColor,
        ...(Object.values(explicitBorders).some((side) => side !== false)
          ? { borders: explicitBorders }
          : {}),
      });
      api.close();
      return;
    }

    const liveWb = getWb();
    const plan = planSelectionFormat(state);
    if (!plan) return;

    const mergeRange = expandRangeWithMerges(state, range);
    const mergeChanged =
      !mergeCk.input.disabled &&
      (mergeCk.input.checked !== initialMergeChecked ||
        (mergeTouched && mergeCk.input.indeterminate !== initialMergeIndeterminate));
    const mergeAction: 'merge' | 'unmerge' | null = mergeChanged
      ? mergeCk.input.checked
        ? 'merge'
        : 'unmerge'
      : null;

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

    // F4 repeats only formatting, never cell-identity metadata. It resolves a
    // fresh selection, workbook, policy, and plan at invocation time.
    const {
      hyperlink: _hyperlink,
      hyperlinkDisplay: _hyperlinkDisplay,
      hyperlinkTooltip: _hyperlinkTooltip,
      comment: _comment,
      commentAuthor: _commentAuthor,
      validation: _validation,
      ...repeatablePatch
    } = actionPatch;
    const repeatAction: SelectionFormatAction = {
      patch: structuredClone(repeatablePatch),
      ...(borderAction ? { border: structuredClone(borderAction) } : {}),
    };
    const hasRepeatAction = Object.keys(repeatablePatch).length > 0 || borderAction !== undefined;
    let repeatFormatting: (() => void) | undefined;
    if (hasRepeatAction) {
      repeatFormatting = (): void => {
        const current = store.getState();
        const currentPlan = planSelectionFormat(current);
        if (!currentPlan) return;
        const currentWb = getWb();
        recordDialogFormatChange({
          history,
          store,
          workbook: currentWb,
          sheet: current.selection.range.sheet,
          targets: currentPlan.cells,
          pendingBefore: current.ui.pendingFormat,
          mutate: () =>
            applySelectionFormatAction(current, store, repeatAction, {
              allowPending: false,
              origin: 'instanceApi',
              commandId: 'formatCells',
            }),
          repeat: repeatFormatting,
        });
      };
    }

    // A merge is a legacy multi-child history action. Use an ephemeral history
    // when the caller did not provide one so a later merge failure can abort
    // the already-applied format child atomically.
    const actionHistory = mergeAction && !history ? new History() : history;
    const transaction = mergeAction && actionHistory ? actionHistory.begin() : undefined;
    let transactionOpen = transaction !== undefined;
    const abortTransaction = (): void => {
      if (!transaction || !actionHistory || !transactionOpen) return;
      transactionOpen = false;
      try {
        actionHistory.abort(transaction);
      } finally {
        // Scoped material replay intentionally preserves a later pending
        // format. A failed composite dialog action instead restores the exact
        // pre-action pending snapshot.
        mutators.setPendingFormat(store, state.ui.pendingFormat);
      }
    };
    let completed = false;
    try {
      if (hasFormatAction) {
        const wrote = recordDialogFormatChange({
          history: actionHistory,
          store,
          workbook: liveWb,
          sheet: range.sheet,
          targets: plan.cells,
          pendingBefore: state.ui.pendingFormat,
          mutate: () =>
            applySelectionFormatAction(store.getState(), store, action, {
              allowPending: false,
              origin: 'instanceApi',
              commandId: 'formatCells',
            }),
          repeat: repeatFormatting,
          registerRepeat: !transaction,
        });
        if (!wrote) {
          abortTransaction();
          return;
        }
      }
      if (mergeAction === 'merge') {
        if (mergeRange.r0 !== mergeRange.r1 || mergeRange.c0 !== mergeRange.c1) {
          const merged = liveWb
            ? applyMerge(store, liveWb, actionHistory, mergeRange)
            : (() => {
                recordMergesChangeWithEngine(actionHistory, store, null, mergeRange.sheet, () => {
                  mutators.mergeRange(store, mergeRange);
                });
                return true;
              })();
          if (!merged) {
            abortTransaction();
            return;
          }
        }
      } else if (mergeAction === 'unmerge') {
        if (!applyUnmerge(store, liveWb, actionHistory, range)) {
          abortTransaction();
          return;
        }
      }
      if (transaction && actionHistory) {
        actionHistory.end(transaction);
        transactionOpen = false;
        if (hasFormatAction && repeatFormatting) actionHistory.setRepeat(repeatFormatting);
      }
      completed = true;
    } catch (error) {
      try {
        abortTransaction();
      } catch (abortError) {
        throw new AggregateError(
          [error, abortError],
          'Format dialog transaction failed and its rollback failed',
          { cause: error },
        );
      }
      throw error;
    }
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

  const onCatClick = (e: MouseEvent): void => {
    const target = e.target as HTMLElement;
    const btn = target.closest('button[data-fc-cat]') as HTMLButtonElement | null;
    if (!btn) return;
    const id = btn.dataset.fcCat as NumberCategory | undefined;
    if (!id) return;
    setNumberCategory(id);
    btn.focus();
  };

  const setNumberCategory = (id: NumberCategory): void => {
    touch('numFmt');
    const previous = draft.numberCategory;
    draft.numberCategory = id;
    if (previous !== id) {
      const fallback = defaultPatternFor(id);
      if (fallback) draft.pattern = fallback;
      if (
        (id === 'currency' || id === 'accounting') &&
        previous !== 'currency' &&
        previous !== 'accounting'
      ) {
        draft.currencySymbol = defaultCurrencySymbolFor(getFormatLocale());
      }
    }
    syncControlsFromDraft();
    renderPreview();
  };

  const focusCategoryByIndex = (idx: number): void => {
    const categories = Array.from(catButtons.keys());
    const next = categories[(idx + categories.length) % categories.length];
    if (!next) return;
    setNumberCategory(next);
    catButtons.get(next)?.focus();
  };

  const onCatKeyDown = (e: KeyboardEvent): void => {
    const btn = (e.target as HTMLElement).closest<HTMLButtonElement>('button[data-fc-cat]');
    if (!btn) return;
    const id = btn.dataset.fcCat as NumberCategory | undefined;
    const categories = Array.from(catButtons.keys());
    const idx = id ? categories.indexOf(id) : -1;
    if (idx < 0) return;
    if (e.key === 'ArrowDown' || e.key === 'ArrowRight') {
      e.preventDefault();
      focusCategoryByIndex(idx + 1);
    } else if (e.key === 'ArrowUp' || e.key === 'ArrowLeft') {
      e.preventDefault();
      focusCategoryByIndex(idx - 1);
    } else if (e.key === 'Home') {
      e.preventDefault();
      focusCategoryByIndex(0);
    } else if (e.key === 'End') {
      e.preventDefault();
      focusCategoryByIndex(categories.length - 1);
    }
  };

  const onDecimalsInput = (): void => {
    touch('numFmt');
    const n = Number.parseInt(decimalsInput.value, 10);
    if (Number.isFinite(n)) draft.decimals = Math.max(0, Math.min(10, n));
    syncControlsFromDraft();
    syncNegativeSamples();
    renderPreview();
  };

  const onThousandsChange = (): void => {
    touch('numFmt');
    draft.thousands = thousandsCk.input.checked;
    syncControlsFromDraft();
    syncNegativeSamples();
    renderPreview();
  };

  const onNegativeStyleClick = (e: Event): void => {
    const item = (e.target as HTMLElement).closest<HTMLButtonElement>('[data-fc-negative-style]');
    const style = item?.dataset.fcNegativeStyle as NegativeStyle | undefined;
    if (!style) return;
    touch('numFmt');
    draft.negativeStyle = style;
    syncControlsFromDraft();
    renderPreview();
  };

  const onSymbolChange = (): void => {
    touch('numFmt');
    draft.currencySymbol = symbolSelect.value;
    syncControlsFromDraft();
    syncNegativeSamples();
    renderPreview();
  };

  const onPatternInput = (): void => {
    touch('numFmt');
    draft.pattern = patternInput.value;
    syncControlsFromDraft();
    syncPatternPresetOptions();
    renderPreview();
  };

  const onPatternPresetChange = (): void => {
    touch('numFmt');
    draft.pattern = patternPresetSelect.value;
    patternInput.value = draft.pattern;
    syncControlsFromDraft();
    renderPreview();
  };

  const onPatternListClick = (e: Event): void => {
    const target = (e.target as HTMLElement | null)?.closest<HTMLButtonElement>(
      '[data-fc-pattern]',
    );
    if (!target) return;
    const pattern = target.dataset.fcPattern;
    if (!pattern) return;
    touch('numFmt');
    draft.pattern = pattern;
    patternInput.value = pattern;
    syncControlsFromDraft();
    syncPatternPresetOptions();
    renderPreview();
  };

  const onHAlignChange = (e: Event): void => {
    const r = e.target as HTMLInputElement;
    if (!r.checked) return;
    touch('align');
    if (draft.align === 'distributed' && r.value !== 'distributed') touch('justifyLastLine');
    draft.align = r.value === 'default' ? undefined : (r.value as CellAlign);
    if (draft.align !== 'distributed') draft.justifyLastLine = false;
    syncJustifyLastLineAvailability();
    hAlignSelect.value = r.value;
    renderPreview();
  };
  const onVAlignChange = (e: Event): void => {
    const r = e.target as HTMLInputElement;
    if (!r.checked) return;
    touch('vAlign');
    draft.vAlign = r.value === 'default' ? undefined : (r.value as CellVAlign);
    vAlignSelect.value = r.value;
    renderPreview();
  };
  const onHAlignSelectChange = (): void => {
    const value = hAlignSelect.value as 'default' | CellAlign;
    touch('align');
    if (draft.align === 'distributed' && value !== 'distributed') touch('justifyLastLine');
    draft.align = value === 'default' ? undefined : value;
    if (draft.align !== 'distributed') draft.justifyLastLine = false;
    syncJustifyLastLineAvailability();
    for (const [id, r] of hAlignRadios) r.checked = id === value;
    renderPreview();
  };
  const onVAlignSelectChange = (): void => {
    const value = vAlignSelect.value as 'default' | CellVAlign;
    touch('vAlign');
    draft.vAlign = value === 'default' ? undefined : value;
    for (const [id, r] of vAlignRadios) r.checked = id === value;
    renderPreview();
  };
  const onWrapChange = (): void => {
    touch('wrap');
    draft.wrap = wrapCk.input.checked;
    renderPreview();
  };
  const onJustifyLastLineChange = (): void => {
    touch('justifyLastLine');
    draft.justifyLastLine = justifyLastLineCk.input.checked;
    renderPreview();
  };
  const onShrinkToFitChange = (): void => {
    touch('shrinkToFit');
    draft.shrinkToFit = shrinkCk.input.checked;
    renderPreview();
  };
  const onIndentInput = (): void => {
    touch('indent');
    const n = Number.parseInt(indentInput.value, 10);
    if (Number.isFinite(n)) draft.indent = Math.max(0, Math.min(15, n));
    renderPreview();
  };
  const onTextDirectionChange = (): void => {
    touch('textDirection');
    draft.textDirection = textDirectionSelect.value as TextDirection;
    renderPreview();
  };
  const onRotationInput = (): void => {
    touch('rotation');
    const n = Number.parseInt(rotationInput.value, 10);
    if (Number.isFinite(n)) draft.rotation = Math.max(-90, Math.min(90, n));
    renderPreview();
  };

  const onMergeChange = (): void => {
    mergeTouched = true;
    mergeCk.input.indeterminate = false;
  };

  const onDialClick = (event: Event): void => {
    const target = event.target as Element | null;
    const dot = target?.closest<HTMLButtonElement>('[data-fc-angle]');
    if (!dot) return;
    const angle = Number.parseInt(dot.dataset.fcAngle ?? '0', 10);
    if (!Number.isFinite(angle)) return;
    touch('rotation');
    draft.rotation = Math.max(-90, Math.min(90, angle));
    rotationInput.value = String(draft.rotation);
    renderPreview();
  };

  const onBoldChange = (): void => {
    touch('bold');
    draft.bold = boldCk.input.checked;
    normalFontCk.input.checked = false;
    syncControlsFromDraft();
    renderPreview();
  };
  const onItalicChange = (): void => {
    touch('italic');
    draft.italic = italicCk.input.checked;
    normalFontCk.input.checked = false;
    syncControlsFromDraft();
    renderPreview();
  };
  const onUnderlineChange = (): void => {
    touch('underline');
    switch (underlineSelect.value) {
      case 'single':
      case 'double':
      case 'singleAccounting':
      case 'doubleAccounting':
        draft.underline = underlineSelect.value;
        break;
      default:
        draft.underline = false;
    }
    normalFontCk.input.checked = false;
    syncControlsFromDraft();
    renderPreview();
  };
  const onStrikeChange = (): void => {
    touch('strike');
    draft.strike = strikeCk.input.checked;
    normalFontCk.input.checked = false;
    syncControlsFromDraft();
    renderPreview();
  };
  const onSuperscriptChange = (): void => {
    touch('fontVertAlign');
    if (superscriptCk.input.checked) {
      draft.fontVertAlign = 'superscript';
      subscriptCk.input.checked = false;
    } else if (draft.fontVertAlign === 'superscript') {
      draft.fontVertAlign = undefined;
    }
    normalFontCk.input.checked = false;
    syncControlsFromDraft();
    renderPreview();
  };
  const onSubscriptChange = (): void => {
    touch('fontVertAlign');
    if (subscriptCk.input.checked) {
      draft.fontVertAlign = 'subscript';
      superscriptCk.input.checked = false;
    } else if (draft.fontVertAlign === 'subscript') {
      draft.fontVertAlign = undefined;
    }
    normalFontCk.input.checked = false;
    syncControlsFromDraft();
    renderPreview();
  };
  const onNormalFontChange = (): void => {
    if (!normalFontCk.input.checked) return;
    touch(
      'bold',
      'italic',
      'underline',
      'strike',
      'fontVertAlign',
      'fontFamily',
      'fontSize',
      'color',
    );
    draft.bold = false;
    draft.italic = false;
    draft.underline = false;
    draft.strike = false;
    draft.fontVertAlign = undefined;
    draft.fontFamily = '';
    draft.fontSize = undefined;
    draft.color = undefined;
    syncControlsFromDraft();
    renderPreview();
  };

  const onFontStyleListClick = (e: Event): void => {
    const item = (e.target as HTMLElement).closest<HTMLButtonElement>('[data-fc-font-style]');
    const style = item?.dataset.fcFontStyle;
    if (!style) return;
    touch('bold', 'italic');
    draft.bold = style === 'bold' || style === 'boldItalic';
    draft.italic = style === 'italic' || style === 'boldItalic';
    boldCk.input.checked = draft.bold;
    italicCk.input.checked = draft.italic;
    normalFontCk.input.checked = false;
    syncControlsFromDraft();
    renderPreview();
  };

  const onFamilyInput = (): void => {
    touch('fontFamily');
    draft.fontFamily = familyInput.value;
    normalFontCk.input.checked = false;
    syncControlsFromDraft();
    renderPreview();
  };

  const onSizeInput = (): void => {
    touch('fontSize');
    if (sizeInput.value === '') {
      draft.fontSize = undefined;
    } else {
      const n = Number.parseInt(sizeInput.value, 10);
      if (Number.isFinite(n)) draft.fontSize = Math.max(1, Math.min(409, n));
    }
    normalFontCk.input.checked = false;
    syncControlsFromDraft();
    renderPreview();
  };

  const onColorInput = (): void => {
    touch('color');
    draft.color = colorInput.value;
    normalFontCk.input.checked = false;
    syncControlsFromDraft();
    renderPreview();
  };
  const onColorReset = (): void => {
    touch('color');
    draft.color = undefined;
    fontSwatches.setValue(null);
    syncControlsFromDraft();
    renderPreview();
  };

  const onFontSwatchesToggle = (): void => fontPalette.setOpen(!fontPalette.isOpen());
  const onFontSwatchClick = (e: Event): void => {
    const btn = (e.target as HTMLElement).closest<HTMLButtonElement>('[data-color]');
    const color = btn?.dataset.color;
    if (!color) return;
    touch('color');
    draft.color = color;
    colorInput.value = color;
    normalFontCk.input.checked = false;
    syncControlsFromDraft();
    renderPreview();
    fontPalette.setOpen(false);
    fontSwatchesToggle.focus();
  };
  const onOverlayPointerDown = (e: Event): void => {
    const target = e.target as Node | null;
    if (!target) return;
    for (const flyout of paletteFlyouts) {
      if (flyout.isOpen() && !flyout.owns(target)) flyout.setOpen(false);
    }
  };

  // Border events
  const onBorderStyleChange = (): void => {
    draft.borderStyle = borderStyleSelect.value as BorderStyleKey;
    pendingBorderPreset = null;
    syncControlsFromDraft();
    renderPreview();
  };
  const onBorderStyleGalleryClick = (e: Event): void => {
    const btn = (e.target as HTMLElement).closest<HTMLButtonElement>('[data-border-style]');
    const style = btn?.dataset.borderStyle as BorderStyleKey | undefined;
    if (!style) return;
    draft.borderStyle = style;
    borderStyleSelect.value = style;
    pendingBorderPreset = null;
    syncControlsFromDraft();
    renderPreview();
  };
  const onBorderColorInput = (): void => {
    draft.borderColor = borderColorInput.value;
    renderPreview();
  };
  const onBorderColorReset = (): void => {
    draft.borderColor = undefined;
    borderSwatches.setValue(null);
    renderPreview();
  };
  const onBorderSwatchesToggle = (): void => borderPalette.setOpen(!borderPalette.isOpen());
  const onBorderSwatchClick = (e: Event): void => {
    const btn = (e.target as HTMLElement).closest<HTMLButtonElement>('[data-color]');
    const color = btn?.dataset.color;
    if (!color) return;
    draft.borderColor = color;
    borderColorInput.value = color;
    renderPreview();
    borderPalette.setOpen(false);
    borderSwatchesToggle.focus();
  };

  const onPresetNone = (): void => {
    touch(
      'border.top',
      'border.right',
      'border.bottom',
      'border.left',
      'border.diagonalDown',
      'border.diagonalUp',
    );
    pendingBorderPreset = 'none';
    draft.borders = {};
    syncControlsFromDraft();
    renderPreview();
  };
  const onPresetOutline = (): void => {
    touch('border.top', 'border.right', 'border.bottom', 'border.left');
    pendingBorderPreset = 'outline';
    draft.borders = {
      top: activeSide(),
      right: activeSide(),
      bottom: activeSide(),
      left: activeSide(),
    };
    syncControlsFromDraft();
    renderPreview();
  };
  const onPresetAll = (): void => {
    touch('border.top', 'border.right', 'border.bottom', 'border.left');
    pendingBorderPreset = 'all';
    draft.borders = {
      top: activeSide(),
      right: activeSide(),
      bottom: activeSide(),
      left: activeSide(),
    };
    syncControlsFromDraft();
    renderPreview();
  };

  const onTopChange = (): void => {
    touch('border.top');
    pendingBorderPreset = null;
    setSide('top', topCk.input.checked);
    renderPreview();
  };
  const onBottomChange = (): void => {
    touch('border.bottom');
    pendingBorderPreset = null;
    setSide('bottom', bottomCk.input.checked);
    renderPreview();
  };
  const onLeftChange = (): void => {
    touch('border.left');
    pendingBorderPreset = null;
    setSide('left', leftCk.input.checked);
    renderPreview();
  };
  const onRightChange = (): void => {
    touch('border.right');
    pendingBorderPreset = null;
    setSide('right', rightCk.input.checked);
    renderPreview();
  };
  const onDiagDownChange = (): void => {
    touch('border.diagonalDown');
    pendingBorderPreset = null;
    setSide('diagonalDown', diagDownCk.input.checked);
    renderPreview();
  };
  const onDiagUpChange = (): void => {
    touch('border.diagonalUp');
    pendingBorderPreset = null;
    setSide('diagonalUp', diagUpCk.input.checked);
    renderPreview();
  };
  const onVisualSideClick = (e: Event): void => {
    const btn = (e.target as HTMLElement).closest<HTMLButtonElement>('[data-border-side]');
    if (!btn) return;
    const key = btn.dataset.borderSide as SideKey;
    touch(`border.${key}` as FormatDialogField);
    pendingBorderPreset = null;
    setSide(key, !draft.borders[key]);
    syncControlsFromDraft();
    renderPreview();
  };

  const onFillInput = (): void => {
    touch('fill');
    draft.fill = fillInput.value;
    syncControlsFromDraft();
    renderPreview();
  };
  const onFillReset = (): void => {
    touch('fill');
    draft.fill = undefined;
    fillSwatches.setValue(null);
    syncControlsFromDraft();
    renderPreview();
  };
  const onFillPatternChange = (): void => {
    touch('fillPattern');
    draft.fillPattern = (fillPatternSelect.value || undefined) as FillPattern | undefined;
    syncControlsFromDraft();
    renderPreview();
  };
  const onFillPatternGalleryClick = (e: Event): void => {
    const button = (e.target as HTMLElement).closest<HTMLButtonElement>('[data-fc-fill-pattern]');
    if (!button) return;
    fillPatternSelect.value = button.dataset.fcFillPattern ?? '';
    onFillPatternChange();
  };
  const onFillPatternColorInput = (): void => {
    touch('fillPatternColor');
    draft.fillPatternColor = fillPatternColorInput.value;
    syncControlsFromDraft();
    renderPreview();
  };
  const onFillSwatchClick = (e: Event): void => {
    const btn = (e.target as HTMLElement).closest<HTMLButtonElement>('[data-color]');
    const color = btn?.dataset.color;
    if (!color) return;
    touch('fill');
    draft.fill = color;
    fillInput.value = color;
    syncControlsFromDraft();
    renderPreview();
  };

  const onLockedChange = (): void => {
    touch('locked');
    draft.locked = lockedCk.input.checked;
  };
  const onHiddenFormulaChange = (): void => {
    touch('formulaHidden');
    draft.formulaHidden = hiddenFormulaCk.input.checked;
  };

  // More tab events
  const onHlInput = (): void => {
    touch('hyperlink');
    draft.hyperlink = hlInput.value;
  };
  const onHlClear = (): void => {
    touch('hyperlink');
    draft.hyperlink = '';
    hlInput.value = '';
  };
  const onCommentInput = (): void => {
    touch('comment');
    draft.comment = commentArea.value;
  };
  const onCommentClear = (): void => {
    touch('comment');
    draft.comment = '';
    commentArea.value = '';
  };
  const onValidationInput = (): void => {
    touch('validation');
    draft.validationList = validationArea.value;
    syncControlsFromDraft();
  };
  const onValidationClear = (): void => {
    touch('validation');
    draft.validationList = '';
    validationArea.value = '';
    syncControlsFromDraft();
  };
  const onValidationListRangeInput = (): void => {
    touch('validation');
    draft.validationListRange = validationListRangeInput.value;
    syncControlsFromDraft();
  };
  const onValidationListSourceKindChange = (): void => {
    touch('validation');
    if (validationListLiteralRadio.input.checked) draft.validationListSourceKind = 'literal';
    else if (validationListRangeRadio.input.checked) draft.validationListSourceKind = 'range';
    syncControlsFromDraft();
  };
  const onValidationShowDropdownChange = (): void => {
    touch('validation');
    draft.validationShowDropdown = validationShowDropdownInput.checked;
    syncControlsFromDraft();
  };
  const onValidationKindChange = (): void => {
    touch('validation');
    draft.validationKind = validationKindSelect.value as ValidationKind;
    // Switching between numeric / date / time kinds swaps the bound-input type,
    // so re-render the stored bounds in the new type's value format.
    applyBoundInputMode(draft.validationKind);
    validationAInput.value = boundInputValue(draft.validationKind, draft.validationA);
    validationBInput.value = boundInputValue(draft.validationKind, draft.validationB);
    syncControlsFromDraft();
  };
  const onValidationOpChange = (): void => {
    touch('validation');
    draft.validationOp = validationOpSelect.value as ValidationOp;
    syncControlsFromDraft();
  };
  const onValidationAInput = (): void => {
    touch('validation');
    const n = parseBoundInputValue(draft.validationKind, validationAInput.value);
    if (n !== null) draft.validationA = n;
    syncControlsFromDraft();
  };
  const onValidationBInput = (): void => {
    touch('validation');
    const n = parseBoundInputValue(draft.validationKind, validationBInput.value);
    if (n !== null) draft.validationB = n;
    syncControlsFromDraft();
  };
  const onValidationFormulaInput = (): void => {
    touch('validation');
    draft.validationFormula = validationFormulaInput.value;
    syncControlsFromDraft();
  };
  const onValidationAllowBlankChange = (): void => {
    touch('validation');
    draft.validationAllowBlank = validationAllowBlankInput.checked;
    syncControlsFromDraft();
  };
  const onValidationErrorStyleChange = (): void => {
    touch('validation');
    draft.validationErrorStyle = validationErrorStyleSelect.value as ValidationErrorStyle;
    syncControlsFromDraft();
  };
  const onValidationShowInputMessageChange = (): void => {
    touch('validation');
    draft.validationShowInputMessage = validationShowInputMessageInput.checked;
    syncControlsFromDraft();
  };
  const onValidationPromptTitleInput = (): void => {
    touch('validation');
    draft.validationPromptTitle = validationPromptTitleInput.value;
    syncControlsFromDraft();
  };
  const onValidationPromptMessageInput = (): void => {
    touch('validation');
    draft.validationPromptMessage = validationPromptMessageArea.value;
    syncControlsFromDraft();
  };
  const onValidationShowErrorMessageChange = (): void => {
    touch('validation');
    draft.validationShowErrorMessage = validationShowErrorMessageInput.checked;
    syncControlsFromDraft();
  };
  const onValidationErrorTitleInput = (): void => {
    touch('validation');
    draft.validationErrorTitle = validationErrorTitleInput.value;
    syncControlsFromDraft();
  };
  const onValidationErrorMessageInput = (): void => {
    touch('validation');
    draft.validationErrorMessage = validationErrorMessageArea.value;
    syncControlsFromDraft();
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
  shell.on(tabsStrip, 'click', onTabClick as EventListener);
  shell.on(tabsStrip, 'keydown', onTabKeyDown as EventListener);
  shell.on(catList, 'click', onCatClick as EventListener);
  shell.on(catList, 'keydown', onCatKeyDown as EventListener);
  shell.on(decimalsInput, 'input', onDecimalsInput);
  shell.on(thousandsCk.input, 'change', onThousandsChange);
  shell.on(negativeOptions, 'click', onNegativeStyleClick as EventListener);
  shell.on(symbolSelect, 'change', onSymbolChange);
  shell.on(patternInput, 'input', onPatternInput);
  shell.on(patternPresetSelect, 'change', onPatternPresetChange);
  shell.on(patternList, 'click', onPatternListClick);
  for (const r of hAlignRadios.values()) shell.on(r, 'change', onHAlignChange);
  for (const r of vAlignRadios.values()) shell.on(r, 'change', onVAlignChange);
  shell.on(hAlignSelect, 'change', onHAlignSelectChange);
  shell.on(vAlignSelect, 'change', onVAlignSelectChange);
  shell.on(wrapCk.input, 'change', onWrapChange);
  shell.on(justifyLastLineCk.input, 'change', onJustifyLastLineChange);
  shell.on(shrinkCk.input, 'change', onShrinkToFitChange);
  shell.on(mergeCk.input, 'change', onMergeChange);
  shell.on(indentInput, 'input', onIndentInput);
  shell.on(textDirectionSelect, 'change', onTextDirectionChange);
  shell.on(rotationInput, 'input', onRotationInput);
  shell.on(alignPreviewDial, 'click', onDialClick);
  shell.on(boldCk.input, 'change', onBoldChange);
  shell.on(italicCk.input, 'change', onItalicChange);
  shell.on(underlineSelect, 'change', onUnderlineChange);
  shell.on(strikeCk.input, 'change', onStrikeChange);
  shell.on(superscriptCk.input, 'change', onSuperscriptChange);
  shell.on(subscriptCk.input, 'change', onSubscriptChange);
  shell.on(normalFontCk.input, 'change', onNormalFontChange);
  shell.on(fontStyleList, 'click', onFontStyleListClick as EventListener);
  shell.on(familyInput, 'input', onFamilyInput);
  shell.on(sizeInput, 'input', onSizeInput);
  shell.on(colorInput, 'input', onColorInput);
  shell.on(colorReset, 'click', onColorReset);
  shell.on(fontSwatchesToggle, 'click', onFontSwatchesToggle);
  shell.on(fontSwatches.el, 'click', onFontSwatchClick);
  shell.on(borderStyleSelect, 'change', onBorderStyleChange);
  shell.on(borderStyleGallery, 'click', onBorderStyleGalleryClick);
  shell.on(borderColorInput, 'input', onBorderColorInput);
  shell.on(borderColorReset, 'click', onBorderColorReset);
  shell.on(borderSwatchesToggle, 'click', onBorderSwatchesToggle);
  shell.on(borderSwatches.el, 'click', onBorderSwatchClick);
  shell.on(presetNone, 'click', onPresetNone);
  shell.on(presetOutline, 'click', onPresetOutline);
  shell.on(presetAll, 'click', onPresetAll);
  shell.on(topCk.input, 'change', onTopChange);
  shell.on(bottomCk.input, 'change', onBottomChange);
  shell.on(leftCk.input, 'change', onLeftChange);
  shell.on(rightCk.input, 'change', onRightChange);
  shell.on(diagDownCk.input, 'change', onDiagDownChange);
  shell.on(diagUpCk.input, 'change', onDiagUpChange);
  shell.on(borderVisualStage, 'click', onVisualSideClick);
  shell.on(fillInput, 'input', onFillInput);
  shell.on(fillReset, 'click', onFillReset);
  shell.on(fillPatternSelect, 'change', onFillPatternChange);
  shell.on(fillPatternGallery, 'click', onFillPatternGalleryClick);
  shell.on(fillPatternColorInput, 'input', onFillPatternColorInput);
  shell.on(fillSwatches.el, 'click', onFillSwatchClick);
  shell.on(lockedCk.input, 'change', onLockedChange);
  shell.on(hiddenFormulaCk.input, 'change', onHiddenFormulaChange);
  shell.on(hlInput, 'input', onHlInput);
  shell.on(hlClear, 'click', onHlClear);
  shell.on(commentArea, 'input', onCommentInput);
  shell.on(commentClear, 'click', onCommentClear);
  shell.on(validationArea, 'input', onValidationInput);
  shell.on(validationClear, 'click', onValidationClear);
  shell.on(validationListRangeInput, 'input', onValidationListRangeInput);
  shell.on(validationListLiteralRadio.input, 'change', onValidationListSourceKindChange);
  shell.on(validationListRangeRadio.input, 'change', onValidationListSourceKindChange);
  shell.on(validationShowDropdownInput, 'change', onValidationShowDropdownChange);
  shell.on(validationKindSelect, 'change', onValidationKindChange);
  shell.on(validationOpSelect, 'change', onValidationOpChange);
  shell.on(validationAInput, 'input', onValidationAInput);
  shell.on(validationBInput, 'input', onValidationBInput);
  shell.on(validationFormulaInput, 'input', onValidationFormulaInput);
  shell.on(validationAllowBlankInput, 'change', onValidationAllowBlankChange);
  shell.on(validationErrorStyleSelect, 'change', onValidationErrorStyleChange);
  shell.on(validationShowInputMessageInput, 'change', onValidationShowInputMessageChange);
  shell.on(validationPromptTitleInput, 'input', onValidationPromptTitleInput);
  shell.on(validationPromptMessageArea, 'input', onValidationPromptMessageInput);
  shell.on(validationShowErrorMessageInput, 'change', onValidationShowErrorMessageChange);
  shell.on(validationErrorTitleInput, 'input', onValidationErrorTitleInput);
  shell.on(validationErrorMessageArea, 'input', onValidationErrorMessageInput);
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
