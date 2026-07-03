import { coerceInput } from '../commands/coerce-input.js';
import { applyFormatPatch, formatNumber } from '../commands/format.js';
import {
  type History,
  recordFormatChange,
  recordMergesChangeWithEngine,
} from '../commands/history.js';
import { applyMerge, applyUnmerge, mergeAt } from '../commands/merge.js';
import { addrKey } from '../engine/address.js';
import { flushFormatToEngine } from '../engine/cell-format-sync.js';
import type { Range } from '../engine/types.js';
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
  type TextDirection,
  type ValidationErrorStyle,
  type ValidationOp,
} from '../store/store.js';
import { appendDialogSelectOptions } from '../toolbar/dialogs/form-controls.js';
import { projectDisabledReason, projectDisabledState } from '../toolbar/menu-a11y.js';
import { formatA1Range } from '../wrappers/toolbar-a1.js';
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
  computeDialogNumFmt,
  computeDialogValidation,
  explicitDraftBorders,
  hydrateDraftFromFormat,
  makeEmptyDraft,
  restyleDraftBorders,
  setDraftSide,
} from './format-dialog-state.js';
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
  mode?: 'format' | 'dataValidation';
  focus?: 'activeTab' | 'validation';
}

export interface FormatDialogHandle {
  open(tab?: TabId, options?: FormatDialogOpenOptions): void;
  close(): void;
  detach(): void;
}

const MAX_OUTLINE_BORDER_CELLS = 100_000;

const rangeArea = (range: Range): number => (range.r1 - range.r0 + 1) * (range.c1 - range.c0 + 1);

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

  const view = createFormatDialogView({ host, strings, t });
  const {
    shell,
    overlay,
    panel,
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
    underlineCk,
    strikeCk,
    normalFontCk,
    fontStyleList,
    familyInput,
    sizeInput,
    colorInput,
    colorReset,
    fontSwatches,
    fontPreviewBox,
    borderStyleSelect,
    borderStyleButtons,
    borderStyleGallery,
    borderColorInput,
    borderColorReset,
    borderSwatches,
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
  const draft: DraftState = makeEmptyDraft(getFormatLocale());

  // ── Hydration ──────────────────────────────────────────────────────────
  const hydrateFromActive = (): void => {
    const state = store.getState();
    const fmt = state.format.formats.get(addrKey(state.selection.active)) ?? {};
    hydrateDraftFromFormat(draft, fmt, getFormatLocale());
    pendingBorderPreset = null;

    syncControlsFromDraft();
    const range = state.selection.range;
    const activeMerge = mergeAt(state, state.selection.active);
    const multiCell = range.r0 !== range.r1 || range.c0 !== range.c1;
    mergeCk.input.checked = activeMerge !== null;
    const mergeDisabled = !multiCell && activeMerge === null;
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
    shrinkCk.input.checked = draft.shrinkToFit;
    indentInput.value = String(draft.indent);
    textDirectionSelect.value = draft.textDirection;
    rotationInput.value = String(draft.rotation);
    syncRotationDial(draft.rotation);

    // Font
    boldCk.input.checked = draft.bold;
    italicCk.input.checked = draft.italic;
    underlineCk.input.checked = draft.underline;
    strikeCk.input.checked = draft.strike;
    normalFontCk.input.checked =
      !draft.bold &&
      !draft.italic &&
      !draft.underline &&
      !draft.strike &&
      !draft.fontFamily &&
      draft.fontSize === undefined &&
      draft.color === undefined;
    syncFontStyleList();
    familyInput.value = draft.fontFamily;
    sizeInput.value = draft.fontSize !== undefined ? String(draft.fontSize) : '';
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
    const listboxCats = new Set<NumberCategory>(['date', 'time', 'datetime', 'special']);
    decimalsRow.hidden = !decimalsCats.has(cat);
    thousandsCk.wrap.hidden = cat !== 'fixed';
    symbolRow.hidden = !symbolCats.has(cat);
    // For date/time-like categories use the Office365-style clickable
    // listbox; only the custom category keeps the dropdown of code-style
    // presets.
    patternPresetRow.hidden = cat !== 'custom';
    patternListWrap.hidden = !listboxCats.has(cat);
    patternRow.hidden = cat !== 'custom';
    localeRow.hidden = cat !== 'date' && cat !== 'time' && cat !== 'datetime' && cat !== 'special';
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
    const symbol = cat === 'currency' ? (draft.currencySymbol ?? '') : '';
    const formatSample = (value: number, style: NegativeStyle): string => {
      const abs = Math.abs(value);
      const grouped = abs.toLocaleString('en-US');
      const body = `${symbol}${grouped}`;
      switch (style) {
        case 'parens':
        case 'red-parens':
          return `(${body})`;
        case 'red':
          return `${symbol}${grouped}`;
        default:
          return `${symbol}-${grouped}`;
      }
    };
    for (const item of items) {
      const style = item.dataset.fcNegativeStyle as NegativeStyle | undefined;
      if (!style) continue;
      item.textContent = formatSample(-1234, style);
    }
  };

  const currentFontStyleId = (): 'regular' | 'italic' | 'bold' | 'boldItalic' => {
    if (draft.bold && draft.italic) return 'boldItalic';
    if (draft.bold) return 'bold';
    if (draft.italic) return 'italic';
    return 'regular';
  };

  const syncFontStyleList = (): void => {
    const id = currentFontStyleId();
    for (const item of fontStyleList.querySelectorAll<HTMLButtonElement>('[data-fc-font-style]')) {
      item.setAttribute('aria-selected', item.dataset.fcFontStyle === id ? 'true' : 'false');
    }
  };

  const fillPatternImage = (pattern: FillPattern | undefined, color = '#000000'): string => {
    switch (pattern) {
      case 'gray125':
        return `radial-gradient(${color} 0.6px, transparent 0.6px)`;
      case 'gray25':
        return `radial-gradient(${color} 1px, transparent 1px)`;
      case 'gray50':
        return `repeating-linear-gradient(45deg, ${color} 0 2px, transparent 2px 4px)`;
      case 'horizontal':
        return `repeating-linear-gradient(0deg, ${color} 0 1px, transparent 1px 4px)`;
      case 'vertical':
        return `repeating-linear-gradient(90deg, ${color} 0 1px, transparent 1px 4px)`;
      case 'diagonalDown':
        return `repeating-linear-gradient(45deg, ${color} 0 1px, transparent 1px 5px)`;
      case 'diagonalUp':
        return `repeating-linear-gradient(135deg, ${color} 0 1px, transparent 1px 5px)`;
      default:
        return '';
    }
  };

  // ── Border helpers ─────────────────────────────────────────────────────
  const activeSide = (): CellBorderSide => activeDraftSide(draft);
  const setSide = (key: SideKey, on: boolean): void => {
    draft.borders = setDraftSide(draft, key, on);
  };
  const restyleExistingSides = (): void => {
    draft.borders = restyleDraftBorders(draft);
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

  const renderPreview = (): void => {
    const applyFontPreview = (el: HTMLElement): void => {
      el.style.fontWeight = draft.bold ? 'bold' : 'normal';
      el.style.fontStyle = draft.italic ? 'italic' : 'normal';
      const decos: string[] = [];
      if (draft.underline) decos.push('underline');
      if (draft.strike) decos.push('line-through');
      el.style.textDecoration = decos.length > 0 ? decos.join(' ') : 'none';
      el.style.fontFamily = draft.fontFamily || '';
      el.style.fontSize = draft.fontSize !== undefined ? `${draft.fontSize}px` : '';
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
    // Pick a sample value that exercises the active category. Date/time
    //  categories use a serial near the present (45123 ≈ 2023-07-16).
    const isDateLike =
      numFmt.kind === 'date' || numFmt.kind === 'time' || numFmt.kind === 'datetime';
    const sampleValue =
      (draft.numberCategory === 'fixed' || draft.numberCategory === 'currency') &&
      draft.negativeStyle !== 'minus'
        ? -1234
        : isDateLike || draft.numberCategory === 'currency' || draft.numberCategory === 'special'
          ? 10
          : 12345;
    const numericText = formatNumber(sampleValue, numFmt, getFormatLocale());
    previewCell.textContent = numericText;
    if (!draft.color && sampleValue < 0) {
      previewCell.style.color =
        draft.negativeStyle === 'red' || draft.negativeStyle === 'red-parens' ? '#c00000' : '';
    }
  };

  // ── Compute helpers ────────────────────────────────────────────────────
  const defaultPatternFor = (cat: NumberCategory): string => {
    const presets =
      cat === 'date' ||
      cat === 'time' ||
      cat === 'datetime' ||
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
      case 'datetime':
        return 'yyyy-mm-dd HH:MM';
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
      cat === 'datetime' ||
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
      case 'datetime':
        return 41348.5625; // 2013-03-14 13:30
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
    const isListbox = cat === 'date' || cat === 'time' || cat === 'datetime' || cat === 'special';
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
                : { kind: 'datetime', pattern },
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
      case 'datetime':
        return t.descDateTime;
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
    overlay.classList.toggle('fc-fmtdlg--data-validation', dataValidationMode);
    panel.setAttribute('aria-label', dataValidationMode ? t.validationLegend : t.title);
    headerTitle.textContent = dataValidationMode ? t.validationLegend : t.title;
    tabsStrip.hidden = dataValidationMode;
    preview.hidden = dataValidationMode;
    hyperlinkSection.hidden = dataValidationMode;
    commentSection.hidden = dataValidationMode;
    validationSection.classList.toggle('fc-fmtdlg__section--standalone', dataValidationMode);
  };

  // ── Apply OK ───────────────────────────────────────────────────────────
  const applyAndClose = (): void => {
    const state = store.getState();
    const range = state.selection.range;

    const validationLines = draft.validationList
      .split(/\r?\n/)
      .map((s) => s.trim())
      .filter((s) => s.length > 0);

    const explicitBorders = explicitDraftBorders(draft);
    const useRangeOutline =
      pendingBorderPreset === 'outline' && (range.r0 !== range.r1 || range.c0 !== range.c1);

    const validation = computeDialogValidation(draft, validationLines);
    const hyperlink = draft.hyperlink.trim();
    const preserveHyperlinkMetadata = hyperlink.length > 0 && hyperlink === draft.originalHyperlink;

    const patch: Partial<CellFormat> = {
      numFmt: computeDialogNumFmt(draft, defaultPatternFor),
      align: draft.align,
      vAlign: draft.vAlign,
      wrap: draft.wrap,
      shrinkToFit: draft.shrinkToFit,
      indent: draft.indent > 0 ? draft.indent : undefined,
      rotation: draft.rotation !== 0 ? draft.rotation : undefined,
      textDirection: draft.textDirection !== 'context' ? draft.textDirection : undefined,
      bold: draft.bold,
      italic: draft.italic,
      underline: draft.underline,
      strike: draft.strike,
      fontFamily: draft.fontFamily ? draft.fontFamily : undefined,
      fontSize: draft.fontSize,
      color: draft.color,
      fill: draft.fill,
      fillPattern: draft.fillPattern,
      fillPatternColor: draft.fillPattern ? draft.fillPatternColor : undefined,
      ...(useRangeOutline ? {} : { borders: explicitBorders }),
      hyperlink: hyperlink ? hyperlink : undefined,
      hyperlinkDisplay: preserveHyperlinkMetadata ? draft.hyperlinkDisplay : undefined,
      hyperlinkTooltip: preserveHyperlinkMetadata ? draft.hyperlinkTooltip : undefined,
      comment: draft.comment ? draft.comment : undefined,
      ...(draft.comment ? {} : { commentAuthor: undefined }),
      validation,
      locked: draft.locked,
      formulaHidden: draft.formulaHidden ? true : undefined,
    };

    const liveWb = getWb();
    const applyOutlineToRange = (): void => {
      if (rangeArea(range) > MAX_OUTLINE_BORDER_CELLS) return;
      const side = activeSide();
      for (let row = range.r0; row <= range.r1; row += 1) {
        for (let col = range.c0; col <= range.c1; col += 1) {
          const borders: CellFormat['borders'] = {};
          if (row === range.r0) borders.top = side;
          if (row === range.r1) borders.bottom = side;
          if (col === range.c0) borders.left = side;
          if (col === range.c1) borders.right = side;
          if (Object.keys(borders).length > 0) {
            applyFormatPatch(
              state,
              store,
              { sheet: range.sheet, r0: row, c0: col, r1: row, c1: col },
              { borders },
              { allowPending: false },
            );
          }
        }
      }
    };
    if (history) history.begin();
    try {
      recordFormatChange(history, store, () => {
        const wroteFormat = applyFormatPatch(state, store, range, patch, { allowPending: false });
        if (wroteFormat && useRangeOutline) applyOutlineToRange();
      });
      if (mergeCk.input.checked) {
        if (range.r0 !== range.r1 || range.c0 !== range.c1) {
          if (liveWb) {
            applyMerge(store, liveWb, history, range);
          } else {
            recordMergesChangeWithEngine(history, store, null, range.sheet, () => {
              mutators.mergeRange(store, range);
            });
          }
        }
      } else {
        applyUnmerge(store, liveWb, history, range);
      }
    } finally {
      if (history) history.end();
    }
    if (liveWb) flushFormatToEngine(liveWb, store, range.sheet);
    api.close();
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
    const n = Number.parseInt(decimalsInput.value, 10);
    if (Number.isFinite(n)) draft.decimals = Math.max(0, Math.min(10, n));
    renderPreview();
  };

  const onThousandsChange = (): void => {
    draft.thousands = thousandsCk.input.checked;
    renderPreview();
  };

  const onNegativeStyleClick = (e: Event): void => {
    const item = (e.target as HTMLElement).closest<HTMLButtonElement>('[data-fc-negative-style]');
    const style = item?.dataset.fcNegativeStyle as NegativeStyle | undefined;
    if (!style) return;
    draft.negativeStyle = style;
    syncControlsFromDraft();
    renderPreview();
  };

  const onSymbolChange = (): void => {
    draft.currencySymbol = symbolSelect.value;
    syncNegativeSamples();
    renderPreview();
  };

  const onPatternInput = (): void => {
    draft.pattern = patternInput.value;
    syncPatternPresetOptions();
    renderPreview();
  };

  const onPatternPresetChange = (): void => {
    draft.pattern = patternPresetSelect.value;
    patternInput.value = draft.pattern;
    renderPreview();
  };

  const onPatternListClick = (e: Event): void => {
    const target = (e.target as HTMLElement | null)?.closest<HTMLButtonElement>(
      '[data-fc-pattern]',
    );
    if (!target) return;
    const pattern = target.dataset.fcPattern;
    if (!pattern) return;
    draft.pattern = pattern;
    patternInput.value = pattern;
    syncPatternPresetOptions();
    renderPreview();
  };

  const onHAlignChange = (e: Event): void => {
    const r = e.target as HTMLInputElement;
    if (!r.checked) return;
    draft.align = r.value === 'default' ? undefined : (r.value as CellAlign);
    hAlignSelect.value = r.value;
    renderPreview();
  };
  const onVAlignChange = (e: Event): void => {
    const r = e.target as HTMLInputElement;
    if (!r.checked) return;
    draft.vAlign = r.value === 'default' ? undefined : (r.value as CellVAlign);
    vAlignSelect.value = r.value;
    renderPreview();
  };
  const onHAlignSelectChange = (): void => {
    const value = hAlignSelect.value as 'default' | CellAlign;
    draft.align = value === 'default' ? undefined : value;
    for (const [id, r] of hAlignRadios) r.checked = id === value;
    renderPreview();
  };
  const onVAlignSelectChange = (): void => {
    const value = vAlignSelect.value as 'default' | CellVAlign;
    draft.vAlign = value === 'default' ? undefined : value;
    for (const [id, r] of vAlignRadios) r.checked = id === value;
    renderPreview();
  };
  const onWrapChange = (): void => {
    draft.wrap = wrapCk.input.checked;
    renderPreview();
  };
  const onShrinkToFitChange = (): void => {
    draft.shrinkToFit = shrinkCk.input.checked;
    renderPreview();
  };
  const onIndentInput = (): void => {
    const n = Number.parseInt(indentInput.value, 10);
    if (Number.isFinite(n)) draft.indent = Math.max(0, Math.min(15, n));
    renderPreview();
  };
  const onTextDirectionChange = (): void => {
    draft.textDirection = textDirectionSelect.value as TextDirection;
    renderPreview();
  };
  const onRotationInput = (): void => {
    const n = Number.parseInt(rotationInput.value, 10);
    if (Number.isFinite(n)) draft.rotation = Math.max(-90, Math.min(90, n));
    renderPreview();
  };

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

  const onDialClick = (event: Event): void => {
    const target = event.target as Element | null;
    const dot = target?.closest<HTMLButtonElement>('[data-fc-angle]');
    if (!dot) return;
    const angle = Number.parseInt(dot.dataset.fcAngle ?? '0', 10);
    if (!Number.isFinite(angle)) return;
    draft.rotation = Math.max(-90, Math.min(90, angle));
    rotationInput.value = String(draft.rotation);
    renderPreview();
  };

  const onBoldChange = (): void => {
    draft.bold = boldCk.input.checked;
    normalFontCk.input.checked = false;
    syncFontStyleList();
    renderPreview();
  };
  const onItalicChange = (): void => {
    draft.italic = italicCk.input.checked;
    normalFontCk.input.checked = false;
    syncFontStyleList();
    renderPreview();
  };
  const onUnderlineChange = (): void => {
    draft.underline = underlineCk.input.checked;
    normalFontCk.input.checked = false;
    renderPreview();
  };
  const onStrikeChange = (): void => {
    draft.strike = strikeCk.input.checked;
    normalFontCk.input.checked = false;
    renderPreview();
  };
  const onNormalFontChange = (): void => {
    if (!normalFontCk.input.checked) return;
    draft.bold = false;
    draft.italic = false;
    draft.underline = false;
    draft.strike = false;
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
    draft.bold = style === 'bold' || style === 'boldItalic';
    draft.italic = style === 'italic' || style === 'boldItalic';
    boldCk.input.checked = draft.bold;
    italicCk.input.checked = draft.italic;
    normalFontCk.input.checked = false;
    syncFontStyleList();
    renderPreview();
  };

  const onFamilyInput = (): void => {
    draft.fontFamily = familyInput.value;
    normalFontCk.input.checked = false;
    renderPreview();
  };

  const onSizeInput = (): void => {
    if (sizeInput.value === '') {
      draft.fontSize = undefined;
    } else {
      const n = Number.parseInt(sizeInput.value, 10);
      if (Number.isFinite(n)) draft.fontSize = Math.max(8, Math.min(72, n));
    }
    normalFontCk.input.checked = false;
    renderPreview();
  };

  const onColorInput = (): void => {
    draft.color = colorInput.value;
    normalFontCk.input.checked = false;
    renderPreview();
  };
  const onColorReset = (): void => {
    draft.color = undefined;
    fontSwatches.setValue(null);
    renderPreview();
  };
  const onFontSwatchClick = (e: Event): void => {
    const btn = (e.target as HTMLElement).closest<HTMLButtonElement>('[data-color]');
    const color = btn?.dataset.color;
    if (!color) return;
    draft.color = color;
    colorInput.value = color;
    renderPreview();
  };

  // Border events
  const onBorderStyleChange = (): void => {
    draft.borderStyle = borderStyleSelect.value as BorderStyleKey;
    restyleExistingSides();
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
    restyleExistingSides();
    pendingBorderPreset = null;
    syncControlsFromDraft();
    renderPreview();
  };
  const onBorderColorInput = (): void => {
    draft.borderColor = borderColorInput.value;
    restyleExistingSides();
    renderPreview();
  };
  const onBorderColorReset = (): void => {
    draft.borderColor = undefined;
    borderSwatches.setValue(null);
    restyleExistingSides();
    renderPreview();
  };
  const onBorderSwatchClick = (e: Event): void => {
    const btn = (e.target as HTMLElement).closest<HTMLButtonElement>('[data-color]');
    const color = btn?.dataset.color;
    if (!color) return;
    draft.borderColor = color;
    borderColorInput.value = color;
    restyleExistingSides();
    renderPreview();
  };

  const onPresetNone = (): void => {
    pendingBorderPreset = 'none';
    draft.borders = {};
    syncControlsFromDraft();
    renderPreview();
  };
  const onPresetOutline = (): void => {
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
    pendingBorderPreset = null;
    setSide('top', topCk.input.checked);
    renderPreview();
  };
  const onBottomChange = (): void => {
    pendingBorderPreset = null;
    setSide('bottom', bottomCk.input.checked);
    renderPreview();
  };
  const onLeftChange = (): void => {
    pendingBorderPreset = null;
    setSide('left', leftCk.input.checked);
    renderPreview();
  };
  const onRightChange = (): void => {
    pendingBorderPreset = null;
    setSide('right', rightCk.input.checked);
    renderPreview();
  };
  const onDiagDownChange = (): void => {
    pendingBorderPreset = null;
    setSide('diagonalDown', diagDownCk.input.checked);
    renderPreview();
  };
  const onDiagUpChange = (): void => {
    pendingBorderPreset = null;
    setSide('diagonalUp', diagUpCk.input.checked);
    renderPreview();
  };
  const onVisualSideClick = (e: Event): void => {
    const btn = (e.target as HTMLElement).closest<HTMLButtonElement>('[data-border-side]');
    if (!btn) return;
    const key = btn.dataset.borderSide as SideKey;
    pendingBorderPreset = null;
    setSide(key, !draft.borders[key]);
    syncControlsFromDraft();
    renderPreview();
  };

  const onFillInput = (): void => {
    draft.fill = fillInput.value;
    renderPreview();
  };
  const onFillReset = (): void => {
    draft.fill = undefined;
    fillSwatches.setValue(null);
    renderPreview();
  };
  const onFillPatternChange = (): void => {
    draft.fillPattern = (fillPatternSelect.value || undefined) as FillPattern | undefined;
    renderPreview();
  };
  const onFillPatternColorInput = (): void => {
    draft.fillPatternColor = fillPatternColorInput.value;
    renderPreview();
  };
  const onFillSwatchClick = (e: Event): void => {
    const btn = (e.target as HTMLElement).closest<HTMLButtonElement>('[data-color]');
    const color = btn?.dataset.color;
    if (!color) return;
    draft.fill = color;
    fillInput.value = color;
    renderPreview();
  };

  const onLockedChange = (): void => {
    draft.locked = lockedCk.input.checked;
  };
  const onHiddenFormulaChange = (): void => {
    draft.formulaHidden = hiddenFormulaCk.input.checked;
  };

  // More tab events
  const onHlInput = (): void => {
    draft.hyperlink = hlInput.value;
  };
  const onHlClear = (): void => {
    draft.hyperlink = '';
    hlInput.value = '';
  };
  const onCommentInput = (): void => {
    draft.comment = commentArea.value;
  };
  const onCommentClear = (): void => {
    draft.comment = '';
    commentArea.value = '';
  };
  const onValidationInput = (): void => {
    draft.validationList = validationArea.value;
  };
  const onValidationClear = (): void => {
    draft.validationList = '';
    validationArea.value = '';
  };
  const onValidationListRangeInput = (): void => {
    draft.validationListRange = validationListRangeInput.value;
  };
  const onValidationListSourceKindChange = (): void => {
    if (validationListLiteralRadio.input.checked) draft.validationListSourceKind = 'literal';
    else if (validationListRangeRadio.input.checked) draft.validationListSourceKind = 'range';
    syncValidationVisibility();
  };
  const onValidationShowDropdownChange = (): void => {
    draft.validationShowDropdown = validationShowDropdownInput.checked;
  };
  const onValidationKindChange = (): void => {
    draft.validationKind = validationKindSelect.value as ValidationKind;
    // Switching between numeric / date / time kinds swaps the bound-input type,
    // so re-render the stored bounds in the new type's value format.
    applyBoundInputMode(draft.validationKind);
    validationAInput.value = boundInputValue(draft.validationKind, draft.validationA);
    validationBInput.value = boundInputValue(draft.validationKind, draft.validationB);
    syncValidationVisibility();
  };
  const onValidationOpChange = (): void => {
    draft.validationOp = validationOpSelect.value as ValidationOp;
    syncValidationVisibility();
  };
  const onValidationAInput = (): void => {
    const n = parseBoundInputValue(draft.validationKind, validationAInput.value);
    if (n !== null) draft.validationA = n;
  };
  const onValidationBInput = (): void => {
    const n = parseBoundInputValue(draft.validationKind, validationBInput.value);
    if (n !== null) draft.validationB = n;
  };
  const onValidationFormulaInput = (): void => {
    draft.validationFormula = validationFormulaInput.value;
  };
  const onValidationAllowBlankChange = (): void => {
    draft.validationAllowBlank = validationAllowBlankInput.checked;
  };
  const onValidationErrorStyleChange = (): void => {
    draft.validationErrorStyle = validationErrorStyleSelect.value as ValidationErrorStyle;
  };
  const onValidationShowInputMessageChange = (): void => {
    draft.validationShowInputMessage = validationShowInputMessageInput.checked;
    syncValidationVisibility();
  };
  const onValidationPromptTitleInput = (): void => {
    draft.validationPromptTitle = validationPromptTitleInput.value;
  };
  const onValidationPromptMessageInput = (): void => {
    draft.validationPromptMessage = validationPromptMessageArea.value;
  };
  const onValidationShowErrorMessageChange = (): void => {
    draft.validationShowErrorMessage = validationShowErrorMessageInput.checked;
    syncValidationVisibility();
  };
  const onValidationErrorTitleInput = (): void => {
    draft.validationErrorTitle = validationErrorTitleInput.value;
  };
  const onValidationErrorMessageInput = (): void => {
    draft.validationErrorMessage = validationErrorMessageArea.value;
  };

  const onOk = (): void => applyAndClose();
  const onCancel = (): void => api.close();

  const onOverlayKey = (e: KeyboardEvent): void => {
    e.stopPropagation();
    if (e.key === 'Escape') {
      e.preventDefault();
      api.close();
      return;
    }
    if (e.key === 'Enter') {
      const target = e.target as HTMLElement;
      const tag = target.tagName;
      // Don't intercept Enter inside textarea or buttons that should activate.
      if (tag === 'BUTTON' || tag === 'TEXTAREA') return;
      e.preventDefault();
      applyAndClose();
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
  shell.on(shrinkCk.input, 'change', onShrinkToFitChange);
  shell.on(indentInput, 'input', onIndentInput);
  shell.on(textDirectionSelect, 'change', onTextDirectionChange);
  shell.on(rotationInput, 'input', onRotationInput);
  shell.on(alignPreviewDial, 'click', onDialClick);
  shell.on(boldCk.input, 'change', onBoldChange);
  shell.on(italicCk.input, 'change', onItalicChange);
  shell.on(underlineCk.input, 'change', onUnderlineChange);
  shell.on(strikeCk.input, 'change', onStrikeChange);
  shell.on(normalFontCk.input, 'change', onNormalFontChange);
  shell.on(fontStyleList, 'click', onFontStyleListClick as EventListener);
  shell.on(familyInput, 'input', onFamilyInput);
  shell.on(sizeInput, 'input', onSizeInput);
  shell.on(colorInput, 'input', onColorInput);
  shell.on(colorReset, 'click', onColorReset);
  shell.on(fontSwatches.el, 'click', onFontSwatchClick);
  shell.on(borderStyleSelect, 'change', onBorderStyleChange);
  shell.on(borderStyleGallery, 'click', onBorderStyleGalleryClick);
  shell.on(borderColorInput, 'input', onBorderColorInput);
  shell.on(borderColorReset, 'click', onBorderColorReset);
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

  const api: FormatDialogHandle = {
    open(tab?: TabId, options?: FormatDialogOpenOptions): void {
      hydrateFromActive();
      setDialogMode(options?.mode);
      if (tab && tabButtons.has(tab)) setActiveTab(tab);
      if (options?.mode === 'dataValidation') setActiveTab('more');
      shell.open();
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
      shell.close();
      host.focus();
    },
    detach(): void {
      shell.dispose();
    },
  };

  return api;
}
