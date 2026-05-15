import { formatNumber } from '../commands/format.js';
import { type History, recordFormatChange } from '../commands/history.js';
import { addrKey } from '../engine/address.js';
import { flushFormatToEngine } from '../engine/cell-format-sync.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import {
  type CellAlign,
  type CellBorderSide,
  type CellFormat,
  type CellVAlign,
  mutators,
  type SpreadsheetStore,
  type ValidationErrorStyle,
  type ValidationOp,
} from '../store/store.js';
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

export interface FormatDialogHandle {
  open(): void;
  close(): void;
  detach(): void;
}

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
    symbolRow,
    symbolSelect,
    patternPresetRow,
    patternPresetSelect,
    patternRow,
    patternInput,
    localeRow,
    localeSelect,
    negativeList,
    numberSummaryTitle,
    numberSummaryDesc,
    hAlignRadios,
    hAlignSelect,
    vAlignRadios,
    vAlignSelect,
    wrapCk,
    indentInput,
    rotationInput,
    boldCk,
    italicCk,
    underlineCk,
    strikeCk,
    familyInput,
    sizeInput,
    colorInput,
    colorReset,
    fontSwatches,
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
    lockedCk,
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
    validationAllowBlankRow,
    validationAllowBlankInput,
    validationErrorStyleRow,
    validationErrorStyleSelect,
    okBtn,
    cancelBtn,
  } = view;

  // ── State ──────────────────────────────────────────────────────────────
  let activeTab: TabId = 'number';
  const draft: DraftState = makeEmptyDraft(getFormatLocale());

  // ── Hydration ──────────────────────────────────────────────────────────
  const hydrateFromActive = (): void => {
    const state = store.getState();
    const fmt = state.format.formats.get(addrKey(state.selection.active)) ?? {};
    hydrateDraftFromFormat(draft, fmt, getFormatLocale());

    syncControlsFromDraft();
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
    indentInput.value = String(draft.indent);
    rotationInput.value = String(draft.rotation);

    // Font
    boldCk.input.checked = draft.bold;
    italicCk.input.checked = draft.italic;
    underlineCk.input.checked = draft.underline;
    strikeCk.input.checked = draft.strike;
    familyInput.value = draft.fontFamily;
    sizeInput.value = draft.fontSize !== undefined ? String(draft.fontSize) : '';
    colorInput.value = draft.color && isHexColor(draft.color) ? draft.color : '#000000';

    // Borders
    borderStyleSelect.value = draft.borderStyle;
    for (const [id, btn] of borderStyleButtons) {
      btn.setAttribute('aria-pressed', id === draft.borderStyle ? 'true' : 'false');
    }
    borderColorInput.value =
      draft.borderColor && isHexColor(draft.borderColor) ? draft.borderColor : '#000000';
    topCk.input.checked = !!draft.borders.top;
    bottomCk.input.checked = !!draft.borders.bottom;
    leftCk.input.checked = !!draft.borders.left;
    rightCk.input.checked = !!draft.borders.right;
    diagDownCk.input.checked = !!draft.borders.diagonalDown;
    diagUpCk.input.checked = !!draft.borders.diagonalUp;

    // Fill
    fillInput.value = draft.fill && isHexColor(draft.fill) ? draft.fill : '#ffffff';

    // Protection
    lockedCk.input.checked = draft.locked;

    // More
    hlInput.value = draft.hyperlink;
    commentArea.value = draft.comment;
    validationArea.value = draft.validationList;
    validationListRangeInput.value = draft.validationListRange;
    validationListLiteralRadio.input.checked = draft.validationListSourceKind === 'literal';
    validationListRangeRadio.input.checked = draft.validationListSourceKind === 'range';
    validationKindSelect.value = draft.validationKind;
    validationOpSelect.value = draft.validationOp;
    validationAInput.value = String(draft.validationA);
    validationBInput.value = String(draft.validationB);
    validationFormulaInput.value = draft.validationFormula;
    validationAllowBlankInput.checked = draft.validationAllowBlank;
    validationErrorStyleSelect.value = draft.validationErrorStyle;
    syncValidationVisibility();
  };

  const syncValidationVisibility = (): void => {
    const k = draft.validationKind;
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
    validationAllowBlankRow.hidden = !isActive;
    validationErrorStyleRow.hidden = !isActive;
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
    const patternCats = new Set<NumberCategory>(['date', 'time', 'datetime', 'special', 'custom']);
    decimalsRow.hidden = !decimalsCats.has(cat);
    symbolRow.hidden = !symbolCats.has(cat);
    patternPresetRow.hidden = !patternCats.has(cat);
    patternRow.hidden = cat !== 'custom';
    localeRow.hidden = cat !== 'date' && cat !== 'time' && cat !== 'datetime' && cat !== 'special';
    localeSelect.value = normalizeFormatLocale(getFormatLocale()).startsWith('ja') ? 'ja' : 'en';
    negativeList.hidden = cat !== 'fixed' && cat !== 'currency';
    const active = catDefs.find((c) => c.id === cat);
    numberSummaryTitle.textContent = active?.label ?? t.catGeneral;
    numberSummaryDesc.textContent = numberCategoryDescription(cat);
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
  const renderPreview = (): void => {
    preview.style.fontWeight = draft.bold ? 'bold' : 'normal';
    preview.style.fontStyle = draft.italic ? 'italic' : 'normal';
    previewCell.style.fontWeight = draft.bold ? 'bold' : 'normal';
    previewCell.style.fontStyle = draft.italic ? 'italic' : 'normal';
    const decos: string[] = [];
    if (draft.underline) decos.push('underline');
    if (draft.strike) decos.push('line-through');
    preview.style.textDecoration = decos.length > 0 ? decos.join(' ') : 'none';
    preview.style.textAlign = draft.align ?? 'left';
    previewCell.style.textDecoration = decos.length > 0 ? decos.join(' ') : 'none';
    previewCell.style.textAlign = draft.align ?? 'left';
    previewCell.style.fontFamily = draft.fontFamily || '';
    previewCell.style.fontSize = draft.fontSize !== undefined ? `${draft.fontSize}px` : '';
    previewCell.style.color = draft.color ?? '';
    previewCell.style.backgroundColor = draft.fill ?? '';
    previewCell.style.whiteSpace = draft.wrap ? 'pre-wrap' : 'nowrap';
    previewCell.style.justifyContent =
      draft.vAlign === 'top' ? 'flex-start' : draft.vAlign === 'bottom' ? 'flex-end' : 'center';
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
      isDateLike || draft.numberCategory === 'currency' || draft.numberCategory === 'special'
        ? 10
        : 12345;
    const numericText = formatNumber(sampleValue, numFmt, getFormatLocale());
    previewCell.textContent = numericText;
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
    for (const [index, pattern] of patterns.entries()) {
      const opt = document.createElement('option');
      opt.value = pattern;
      opt.textContent = cat === 'special' ? (specialLabels[index] ?? pattern) : pattern;
      patternPresetSelect.appendChild(opt);
    }
    patternPresetSelect.value = current;
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

    const validation = computeDialogValidation(draft, validationLines);

    const patch: Partial<CellFormat> = {
      numFmt: computeDialogNumFmt(draft, defaultPatternFor),
      align: draft.align,
      vAlign: draft.vAlign,
      wrap: draft.wrap,
      indent: draft.indent > 0 ? draft.indent : undefined,
      rotation: draft.rotation !== 0 ? draft.rotation : undefined,
      bold: draft.bold,
      italic: draft.italic,
      underline: draft.underline,
      strike: draft.strike,
      fontFamily: draft.fontFamily ? draft.fontFamily : undefined,
      fontSize: draft.fontSize,
      color: draft.color,
      fill: draft.fill,
      borders: explicitBorders,
      hyperlink: draft.hyperlink.trim() ? draft.hyperlink.trim() : undefined,
      comment: draft.comment ? draft.comment : undefined,
      validation,
      locked: draft.locked,
    };

    recordFormatChange(history, store, () => {
      mutators.setRangeFormat(store, range, patch);
    });
    const liveWb = getWb();
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

  const onSymbolChange = (): void => {
    draft.currencySymbol = symbolSelect.value;
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
  const onIndentInput = (): void => {
    const n = Number.parseInt(indentInput.value, 10);
    if (Number.isFinite(n)) draft.indent = Math.max(0, Math.min(15, n));
    renderPreview();
  };
  const onRotationInput = (): void => {
    const n = Number.parseInt(rotationInput.value, 10);
    if (Number.isFinite(n)) draft.rotation = Math.max(-90, Math.min(90, n));
    renderPreview();
  };

  const onBoldChange = (): void => {
    draft.bold = boldCk.input.checked;
    renderPreview();
  };
  const onItalicChange = (): void => {
    draft.italic = italicCk.input.checked;
    renderPreview();
  };
  const onUnderlineChange = (): void => {
    draft.underline = underlineCk.input.checked;
    renderPreview();
  };
  const onStrikeChange = (): void => {
    draft.strike = strikeCk.input.checked;
    renderPreview();
  };

  const onFamilyInput = (): void => {
    draft.fontFamily = familyInput.value;
    renderPreview();
  };

  const onSizeInput = (): void => {
    if (sizeInput.value === '') {
      draft.fontSize = undefined;
    } else {
      const n = Number.parseInt(sizeInput.value, 10);
      if (Number.isFinite(n)) draft.fontSize = Math.max(8, Math.min(72, n));
    }
    renderPreview();
  };

  const onColorInput = (): void => {
    draft.color = colorInput.value;
    renderPreview();
  };
  const onColorReset = (): void => {
    draft.color = undefined;
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
    draft.borders = {};
    syncControlsFromDraft();
    renderPreview();
  };
  const onPresetOutline = (): void => {
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
    setSide('top', topCk.input.checked);
    renderPreview();
  };
  const onBottomChange = (): void => {
    setSide('bottom', bottomCk.input.checked);
    renderPreview();
  };
  const onLeftChange = (): void => {
    setSide('left', leftCk.input.checked);
    renderPreview();
  };
  const onRightChange = (): void => {
    setSide('right', rightCk.input.checked);
    renderPreview();
  };
  const onDiagDownChange = (): void => {
    setSide('diagonalDown', diagDownCk.input.checked);
    renderPreview();
  };
  const onDiagUpChange = (): void => {
    setSide('diagonalUp', diagUpCk.input.checked);
    renderPreview();
  };
  const onVisualSideClick = (e: Event): void => {
    const btn = (e.target as HTMLElement).closest<HTMLButtonElement>('[data-border-side]');
    if (!btn) return;
    const key = btn.dataset.borderSide as SideKey;
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
  const onValidationKindChange = (): void => {
    draft.validationKind = validationKindSelect.value as ValidationKind;
    syncValidationVisibility();
  };
  const onValidationOpChange = (): void => {
    draft.validationOp = validationOpSelect.value as ValidationOp;
    syncValidationVisibility();
  };
  const onValidationAInput = (): void => {
    const n = Number.parseFloat(validationAInput.value);
    if (Number.isFinite(n)) draft.validationA = n;
  };
  const onValidationBInput = (): void => {
    const n = Number.parseFloat(validationBInput.value);
    if (Number.isFinite(n)) draft.validationB = n;
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
  shell.on(symbolSelect, 'change', onSymbolChange);
  shell.on(patternInput, 'input', onPatternInput);
  shell.on(patternPresetSelect, 'change', onPatternPresetChange);
  for (const r of hAlignRadios.values()) shell.on(r, 'change', onHAlignChange);
  for (const r of vAlignRadios.values()) shell.on(r, 'change', onVAlignChange);
  shell.on(hAlignSelect, 'change', onHAlignSelectChange);
  shell.on(vAlignSelect, 'change', onVAlignSelectChange);
  shell.on(wrapCk.input, 'change', onWrapChange);
  shell.on(indentInput, 'input', onIndentInput);
  shell.on(rotationInput, 'input', onRotationInput);
  shell.on(boldCk.input, 'change', onBoldChange);
  shell.on(italicCk.input, 'change', onItalicChange);
  shell.on(underlineCk.input, 'change', onUnderlineChange);
  shell.on(strikeCk.input, 'change', onStrikeChange);
  shell.on(familyInput, 'input', onFamilyInput);
  shell.on(sizeInput, 'input', onSizeInput);
  shell.on(colorInput, 'input', onColorInput);
  shell.on(colorReset, 'click', onColorReset);
  shell.on(fontSwatches, 'click', onFontSwatchClick);
  shell.on(borderStyleSelect, 'change', onBorderStyleChange);
  shell.on(borderStyleGallery, 'click', onBorderStyleGalleryClick);
  shell.on(borderColorInput, 'input', onBorderColorInput);
  shell.on(borderColorReset, 'click', onBorderColorReset);
  shell.on(borderSwatches, 'click', onBorderSwatchClick);
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
  shell.on(fillSwatches, 'click', onFillSwatchClick);
  shell.on(lockedCk.input, 'change', onLockedChange);
  shell.on(hlInput, 'input', onHlInput);
  shell.on(hlClear, 'click', onHlClear);
  shell.on(commentArea, 'input', onCommentInput);
  shell.on(commentClear, 'click', onCommentClear);
  shell.on(validationArea, 'input', onValidationInput);
  shell.on(validationClear, 'click', onValidationClear);
  shell.on(validationListRangeInput, 'input', onValidationListRangeInput);
  shell.on(validationListLiteralRadio.input, 'change', onValidationListSourceKindChange);
  shell.on(validationListRangeRadio.input, 'change', onValidationListSourceKindChange);
  shell.on(validationKindSelect, 'change', onValidationKindChange);
  shell.on(validationOpSelect, 'change', onValidationOpChange);
  shell.on(validationAInput, 'input', onValidationAInput);
  shell.on(validationBInput, 'input', onValidationBInput);
  shell.on(validationFormulaInput, 'input', onValidationFormulaInput);
  shell.on(validationAllowBlankInput, 'change', onValidationAllowBlankChange);
  shell.on(validationErrorStyleSelect, 'change', onValidationErrorStyleChange);
  shell.on(okBtn, 'click', onOk);
  shell.on(cancelBtn, 'click', onCancel);
  shell.on(overlay, 'click', (e) => {
    if ((e as MouseEvent).target === overlay) api.close();
  });
  shell.on(overlay, 'keydown', onOverlayKey as EventListener);

  const api: FormatDialogHandle = {
    open(): void {
      hydrateFromActive();
      shell.open();
      requestAnimationFrame(() => {
        const first = tabButtons.get(activeTab);
        if (first) first.focus();
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
