import { formatNumber } from '../commands/format.js';
import { type History, recordFormatChange } from '../commands/history.js';
import { flushFormatToEngine } from '../engine/cell-format-sync.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { addrKey } from '../engine/workbook-handle.js';
import { type Strings, defaultStrings } from '../i18n/strings.js';
import {
  type CellAlign,
  type CellBorderSide,
  type CellBorders,
  type CellFormat,
  type CellVAlign,
  type NumFmt,
  type SpreadsheetStore,
  mutators,
} from '../store/store.js';

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
}

export interface FormatDialogHandle {
  open(): void;
  close(): void;
  detach(): void;
}

type TabId = 'number' | 'align' | 'font' | 'border' | 'fill' | 'more';
type NumberCategory = 'general' | 'fixed' | 'currency' | 'percent';
type BorderStyleKey = 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double';
type SideKey = 'top' | 'right' | 'bottom' | 'left' | 'diagonalDown' | 'diagonalUp';

interface DraftState {
  numFmt: NumFmt | undefined;
  numberCategory: NumberCategory;
  decimals: number;
  currencySymbol: string;
  align: CellAlign | undefined;
  vAlign: CellVAlign | undefined;
  wrap: boolean;
  indent: number;
  rotation: number;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  fontFamily: string;
  fontSize: number | undefined;
  color: string | undefined;
  fill: string | undefined;
  borders: CellBorders;
  /** "Active" line style — applied when a side checkbox is turned on. */
  borderStyle: BorderStyleKey;
  /** "Active" line color in #rrggbb form, or undefined for theme default. */
  borderColor: string | undefined;
  hyperlink: string;
  comment: string;
  validationList: string;
}

const COMMON_FONTS = [
  'system-ui',
  'Helvetica',
  'Arial',
  'Georgia',
  'Times New Roman',
  'Courier New',
  'monospace',
];
const CURRENCY_SYMBOLS = ['$', '¥', '€', '£'];

export function attachFormatDialog(deps: FormatDialogDeps): FormatDialogHandle {
  const { host, store } = deps;
  const history = deps.history ?? null;
  const getWb = deps.getWb ?? ((): WorkbookHandle | null => null);
  const strings = deps.strings ?? defaultStrings;
  const t = strings.formatDialog;

  // Root overlay (backdrop)
  const overlay = document.createElement('div');
  overlay.className = 'fc-fmtdlg';
  overlay.setAttribute('role', 'dialog');
  overlay.setAttribute('aria-modal', 'true');
  overlay.setAttribute('aria-label', t.title);
  overlay.hidden = true;

  // Panel
  const panel = document.createElement('div');
  panel.className = 'fc-fmtdlg__panel';
  overlay.appendChild(panel);

  // Header
  const header = document.createElement('div');
  header.className = 'fc-fmtdlg__header';
  header.textContent = t.title;
  panel.appendChild(header);

  // Preview
  const preview = document.createElement('div');
  preview.className = 'fc-fmtdlg__preview';
  panel.appendChild(preview);

  // Tabs strip
  const tabsStrip = document.createElement('div');
  tabsStrip.className = 'fc-fmtdlg__tabs';
  tabsStrip.setAttribute('role', 'tablist');
  panel.appendChild(tabsStrip);

  // Body
  const body = document.createElement('div');
  body.className = 'fc-fmtdlg__body';
  panel.appendChild(body);

  // Footer
  const footer = document.createElement('div');
  footer.className = 'fc-fmtdlg__footer';
  panel.appendChild(footer);

  // ── Tab buttons + panels ───────────────────────────────────────────────
  const tabDefs: { id: TabId; label: string }[] = [
    { id: 'number', label: t.tabNumber },
    { id: 'align', label: t.tabAlign },
    { id: 'font', label: t.tabFont },
    { id: 'border', label: t.tabBorder },
    { id: 'fill', label: t.tabFill },
    { id: 'more', label: t.tabMore },
  ];
  const tabButtons = new Map<TabId, HTMLButtonElement>();
  const tabPanels = new Map<TabId, HTMLDivElement>();

  for (const def of tabDefs) {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'fc-fmtdlg__tab';
    btn.textContent = def.label;
    btn.setAttribute('role', 'tab');
    btn.setAttribute('aria-selected', 'false');
    btn.dataset.fcTab = def.id;
    tabsStrip.appendChild(btn);
    tabButtons.set(def.id, btn);

    const panelEl = document.createElement('div');
    panelEl.className = 'fc-fmtdlg__panel-tab';
    panelEl.setAttribute('role', 'tabpanel');
    panelEl.dataset.fcTab = def.id;
    panelEl.hidden = true;
    body.appendChild(panelEl);
    tabPanels.set(def.id, panelEl);
  }

  const makeCheckbox = (label: string): { wrap: HTMLLabelElement; input: HTMLInputElement } => {
    const wrap = document.createElement('label');
    wrap.className = 'fc-fmtdlg__check';
    const input = document.createElement('input');
    input.type = 'checkbox';
    const span = document.createElement('span');
    span.textContent = label;
    wrap.append(input, span);
    return { wrap, input };
  };

  const makeBtn = (label: string, primary = false): HTMLButtonElement => {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = primary ? 'fc-fmtdlg__btn fc-fmtdlg__btn--primary' : 'fc-fmtdlg__btn';
    b.textContent = label;
    return b;
  };

  // ── Number tab ─────────────────────────────────────────────────────────
  const numberPanel = tabPanels.get('number') as HTMLDivElement;
  const numberLayout = document.createElement('div');
  numberLayout.className = 'fc-fmtdlg__row';
  numberPanel.appendChild(numberLayout);

  const catList = document.createElement('div');
  catList.className = 'fc-fmtdlg__cat';
  catList.setAttribute('role', 'listbox');
  numberLayout.appendChild(catList);

  const catDefs: { id: NumberCategory; label: string }[] = [
    { id: 'general', label: t.catGeneral },
    { id: 'fixed', label: t.catFixed },
    { id: 'currency', label: t.catCurrency },
    { id: 'percent', label: t.catPercent },
  ];
  const catButtons = new Map<NumberCategory, HTMLButtonElement>();
  for (const c of catDefs) {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'fc-fmtdlg__cat-item';
    b.textContent = c.label;
    b.setAttribute('role', 'option');
    b.dataset.fcCat = c.id;
    catList.appendChild(b);
    catButtons.set(c.id, b);
  }

  const numberControls = document.createElement('div');
  numberControls.className = 'fc-fmtdlg__cat-controls';
  numberLayout.appendChild(numberControls);

  const decimalsRow = document.createElement('label');
  decimalsRow.className = 'fc-fmtdlg__row';
  const decimalsLabel = document.createElement('span');
  decimalsLabel.textContent = t.decimals;
  const decimalsInput = document.createElement('input');
  decimalsInput.type = 'number';
  decimalsInput.min = '0';
  decimalsInput.max = '10';
  decimalsInput.step = '1';
  decimalsRow.append(decimalsLabel, decimalsInput);
  numberControls.appendChild(decimalsRow);

  const symbolRow = document.createElement('label');
  symbolRow.className = 'fc-fmtdlg__row';
  const symbolLabel = document.createElement('span');
  symbolLabel.textContent = t.symbol;
  const symbolSelect = document.createElement('select');
  for (const s of CURRENCY_SYMBOLS) {
    const opt = document.createElement('option');
    opt.value = s;
    opt.textContent = s;
    symbolSelect.appendChild(opt);
  }
  symbolRow.append(symbolLabel, symbolSelect);
  numberControls.appendChild(symbolRow);

  // ── Alignment tab ──────────────────────────────────────────────────────
  const alignPanel = tabPanels.get('align') as HTMLDivElement;

  // Horizontal
  const hAlignLegend = document.createElement('div');
  hAlignLegend.textContent = t.horizontalAlign;
  alignPanel.appendChild(hAlignLegend);
  const hAlignFieldset = document.createElement('div');
  hAlignFieldset.className = 'fc-fmtdlg__row';
  alignPanel.appendChild(hAlignFieldset);

  const hAlignName = `fc-fmtdlg-halign-${Math.random().toString(36).slice(2, 8)}`;
  const hAlignDefs: { id: 'default' | CellAlign; label: string }[] = [
    { id: 'default', label: t.alignDefault },
    { id: 'left', label: t.alignLeft },
    { id: 'center', label: t.alignCenter },
    { id: 'right', label: t.alignRight },
  ];
  const hAlignRadios = new Map<'default' | CellAlign, HTMLInputElement>();
  for (const a of hAlignDefs) {
    const wrap = document.createElement('label');
    wrap.className = 'fc-fmtdlg__radio';
    const radio = document.createElement('input');
    radio.type = 'radio';
    radio.name = hAlignName;
    radio.value = a.id;
    const txt = document.createElement('span');
    txt.textContent = a.label;
    wrap.append(radio, txt);
    hAlignFieldset.appendChild(wrap);
    hAlignRadios.set(a.id, radio);
  }

  // Vertical
  const vAlignLegend = document.createElement('div');
  vAlignLegend.textContent = t.verticalAlign;
  alignPanel.appendChild(vAlignLegend);
  const vAlignFieldset = document.createElement('div');
  vAlignFieldset.className = 'fc-fmtdlg__row';
  alignPanel.appendChild(vAlignFieldset);

  const vAlignName = `fc-fmtdlg-valign-${Math.random().toString(36).slice(2, 8)}`;
  const vAlignDefs: { id: 'default' | CellVAlign; label: string }[] = [
    { id: 'default', label: t.alignDefault },
    { id: 'top', label: t.vAlignTop },
    { id: 'middle', label: t.vAlignMiddle },
    { id: 'bottom', label: t.vAlignBottom },
  ];
  const vAlignRadios = new Map<'default' | CellVAlign, HTMLInputElement>();
  for (const a of vAlignDefs) {
    const wrap = document.createElement('label');
    wrap.className = 'fc-fmtdlg__radio';
    const radio = document.createElement('input');
    radio.type = 'radio';
    radio.name = vAlignName;
    radio.value = a.id;
    const txt = document.createElement('span');
    txt.textContent = a.label;
    wrap.append(radio, txt);
    vAlignFieldset.appendChild(wrap);
    vAlignRadios.set(a.id, radio);
  }

  // Wrap / Indent / Rotation
  const wrapRow = document.createElement('div');
  wrapRow.className = 'fc-fmtdlg__row';
  alignPanel.appendChild(wrapRow);
  const wrapCk = makeCheckbox(t.wrap);
  wrapCk.input.dataset.fcCheck = 'wrap';
  wrapRow.append(wrapCk.wrap);

  const indentRow = document.createElement('label');
  indentRow.className = 'fc-fmtdlg__row';
  const indentLabel = document.createElement('span');
  indentLabel.textContent = t.indent;
  const indentInput = document.createElement('input');
  indentInput.type = 'number';
  indentInput.min = '0';
  indentInput.max = '15';
  indentInput.step = '1';
  indentRow.append(indentLabel, indentInput);
  alignPanel.appendChild(indentRow);

  const rotationRow = document.createElement('label');
  rotationRow.className = 'fc-fmtdlg__row';
  const rotationLabel = document.createElement('span');
  rotationLabel.textContent = t.rotation;
  const rotationInput = document.createElement('input');
  rotationInput.type = 'number';
  rotationInput.min = '-90';
  rotationInput.max = '90';
  rotationInput.step = '1';
  rotationRow.append(rotationLabel, rotationInput);
  alignPanel.appendChild(rotationRow);

  // ── Font tab ───────────────────────────────────────────────────────────
  const fontPanel = tabPanels.get('font') as HTMLDivElement;

  const styleRow = document.createElement('div');
  styleRow.className = 'fc-fmtdlg__row';
  fontPanel.appendChild(styleRow);

  const boldCk = makeCheckbox(t.fontBold);
  boldCk.input.dataset.fcCheck = 'bold';
  const italicCk = makeCheckbox(t.fontItalic);
  italicCk.input.dataset.fcCheck = 'italic';
  const underlineCk = makeCheckbox(t.fontUnderline);
  underlineCk.input.dataset.fcCheck = 'underline';
  const strikeCk = makeCheckbox(t.fontStrike);
  strikeCk.input.dataset.fcCheck = 'strike';
  styleRow.append(boldCk.wrap, italicCk.wrap, underlineCk.wrap, strikeCk.wrap);

  // Font family
  const familyRow = document.createElement('label');
  familyRow.className = 'fc-fmtdlg__row';
  const familyLabel = document.createElement('span');
  familyLabel.textContent = t.fontFamily;
  const familyInput = document.createElement('input');
  familyInput.type = 'text';
  familyInput.spellcheck = false;
  familyInput.autocomplete = 'off';
  const familyListId = `fc-fmtdlg-fonts-${Math.random().toString(36).slice(2, 8)}`;
  familyInput.setAttribute('list', familyListId);
  const familyDatalist = document.createElement('datalist');
  familyDatalist.id = familyListId;
  for (const f of COMMON_FONTS) {
    const opt = document.createElement('option');
    opt.value = f;
    familyDatalist.appendChild(opt);
  }
  familyRow.append(familyLabel, familyInput, familyDatalist);
  fontPanel.appendChild(familyRow);

  // Font size
  const sizeRow = document.createElement('label');
  sizeRow.className = 'fc-fmtdlg__row';
  const sizeLabel = document.createElement('span');
  sizeLabel.textContent = t.fontSize;
  const sizeInput = document.createElement('input');
  sizeInput.type = 'number';
  sizeInput.min = '8';
  sizeInput.max = '72';
  sizeInput.step = '1';
  sizeRow.append(sizeLabel, sizeInput);
  fontPanel.appendChild(sizeRow);

  // Font color
  const colorRow = document.createElement('div');
  colorRow.className = 'fc-fmtdlg__row';
  const colorLabel = document.createElement('span');
  colorLabel.textContent = t.color;
  const colorInput = document.createElement('input');
  colorInput.type = 'color';
  colorInput.dataset.fcColor = 'font';
  const colorReset = makeBtn(t.resetToDefault);
  colorRow.append(colorLabel, colorInput, colorReset);
  fontPanel.appendChild(colorRow);

  // ── Border tab ─────────────────────────────────────────────────────────
  const borderPanel = tabPanels.get('border') as HTMLDivElement;

  // Active style + color row
  const borderStyleRow = document.createElement('label');
  borderStyleRow.className = 'fc-fmtdlg__row';
  const borderStyleLabel = document.createElement('span');
  borderStyleLabel.textContent = t.borderStyle;
  const borderStyleSelect = document.createElement('select');
  const styleOptions: { id: BorderStyleKey; label: string }[] = [
    { id: 'thin', label: t.borderStyleThin },
    { id: 'medium', label: t.borderStyleMedium },
    { id: 'thick', label: t.borderStyleThick },
    { id: 'dashed', label: t.borderStyleDashed },
    { id: 'dotted', label: t.borderStyleDotted },
    { id: 'double', label: t.borderStyleDouble },
  ];
  for (const s of styleOptions) {
    const opt = document.createElement('option');
    opt.value = s.id;
    opt.textContent = s.label;
    borderStyleSelect.appendChild(opt);
  }
  borderStyleRow.append(borderStyleLabel, borderStyleSelect);
  borderPanel.appendChild(borderStyleRow);

  const borderColorRow = document.createElement('div');
  borderColorRow.className = 'fc-fmtdlg__row';
  const borderColorLabel = document.createElement('span');
  borderColorLabel.textContent = t.borderColor;
  const borderColorInput = document.createElement('input');
  borderColorInput.type = 'color';
  borderColorInput.dataset.fcColor = 'border';
  const borderColorReset = makeBtn(t.resetToDefault);
  borderColorRow.append(borderColorLabel, borderColorInput, borderColorReset);
  borderPanel.appendChild(borderColorRow);

  // Presets
  const presetRow = document.createElement('div');
  presetRow.className = 'fc-fmtdlg__row';
  borderPanel.appendChild(presetRow);
  const presetNone = makeBtn(t.borderPresetNone);
  const presetOutline = makeBtn(t.borderPresetOutline);
  const presetAll = makeBtn(t.borderPresetAll);
  presetRow.append(presetNone, presetOutline, presetAll);

  // Per-side checkboxes
  const sideRow = document.createElement('div');
  sideRow.className = 'fc-fmtdlg__row';
  borderPanel.appendChild(sideRow);
  const topCk = makeCheckbox(t.borderTop);
  topCk.input.dataset.fcCheck = 'borderTop';
  const bottomCk = makeCheckbox(t.borderBottom);
  bottomCk.input.dataset.fcCheck = 'borderBottom';
  const leftCk = makeCheckbox(t.borderLeft);
  leftCk.input.dataset.fcCheck = 'borderLeft';
  const rightCk = makeCheckbox(t.borderRight);
  rightCk.input.dataset.fcCheck = 'borderRight';
  sideRow.append(topCk.wrap, bottomCk.wrap, leftCk.wrap, rightCk.wrap);

  const diagonalRow = document.createElement('div');
  diagonalRow.className = 'fc-fmtdlg__row';
  borderPanel.appendChild(diagonalRow);
  const diagDownCk = makeCheckbox(t.borderDiagonalDown);
  diagDownCk.input.dataset.fcCheck = 'borderDiagonalDown';
  const diagUpCk = makeCheckbox(t.borderDiagonalUp);
  diagUpCk.input.dataset.fcCheck = 'borderDiagonalUp';
  diagonalRow.append(diagDownCk.wrap, diagUpCk.wrap);

  // ── Fill tab ───────────────────────────────────────────────────────────
  const fillPanel = tabPanels.get('fill') as HTMLDivElement;
  const fillRow = document.createElement('div');
  fillRow.className = 'fc-fmtdlg__row';
  const fillLabel = document.createElement('span');
  fillLabel.textContent = t.fill;
  const fillInput = document.createElement('input');
  fillInput.type = 'color';
  fillInput.dataset.fcColor = 'fill';
  const fillReset = makeBtn(t.fillNone);
  fillRow.append(fillLabel, fillInput, fillReset);
  fillPanel.appendChild(fillRow);

  // ── More tab (hyperlink / comment / validation) ────────────────────────
  const morePanel = tabPanels.get('more') as HTMLDivElement;

  const hlRow = document.createElement('div');
  hlRow.className = 'fc-fmtdlg__row';
  morePanel.appendChild(hlRow);
  const hlLabel = document.createElement('span');
  hlLabel.textContent = t.hyperlink;
  const hlInput = document.createElement('input');
  hlInput.type = 'text';
  hlInput.spellcheck = false;
  hlInput.autocomplete = 'off';
  hlInput.placeholder = t.hyperlinkPlaceholder;
  const hlClear = makeBtn(t.clearField);
  hlRow.append(hlLabel, hlInput, hlClear);

  const commentLegend = document.createElement('div');
  commentLegend.textContent = t.comment;
  morePanel.appendChild(commentLegend);
  const commentRow = document.createElement('div');
  commentRow.className = 'fc-fmtdlg__row fc-fmtdlg__row--block';
  morePanel.appendChild(commentRow);
  const commentArea = document.createElement('textarea');
  commentArea.className = 'fc-fmtdlg__textarea';
  commentArea.rows = 3;
  commentArea.placeholder = t.commentPlaceholder;
  const commentClear = makeBtn(t.clearField);
  commentRow.append(commentArea, commentClear);

  const validationLegend = document.createElement('div');
  validationLegend.textContent = t.validationListSource;
  morePanel.appendChild(validationLegend);
  const validationRow = document.createElement('div');
  validationRow.className = 'fc-fmtdlg__row fc-fmtdlg__row--block';
  morePanel.appendChild(validationRow);
  const validationArea = document.createElement('textarea');
  validationArea.className = 'fc-fmtdlg__textarea';
  validationArea.rows = 4;
  validationArea.placeholder = t.validationListPlaceholder;
  const validationClear = makeBtn(t.clearField);
  validationRow.append(validationArea, validationClear);

  // ── Footer buttons ─────────────────────────────────────────────────────
  const cancelBtn = makeBtn(t.cancel);
  const okBtn = makeBtn(t.ok, true);
  footer.append(cancelBtn, okBtn);

  host.appendChild(overlay);

  // ── State ──────────────────────────────────────────────────────────────
  let activeTab: TabId = 'number';
  const draft: DraftState = makeEmptyDraft();

  function makeEmptyDraft(): DraftState {
    return {
      numFmt: undefined,
      numberCategory: 'general',
      decimals: 2,
      currencySymbol: '$',
      align: undefined,
      vAlign: undefined,
      wrap: false,
      indent: 0,
      rotation: 0,
      bold: false,
      italic: false,
      underline: false,
      strike: false,
      fontFamily: '',
      fontSize: undefined,
      color: undefined,
      fill: undefined,
      borders: {},
      borderStyle: 'thin',
      borderColor: undefined,
      hyperlink: '',
      comment: '',
      validationList: '',
    };
  }

  // ── Hydration ──────────────────────────────────────────────────────────
  const sideStyle = (s: CellBorderSide | undefined): BorderStyleKey | null => {
    if (!s) return null;
    if (typeof s === 'object') return s.style;
    return 'thin';
  };
  const sideColor = (s: CellBorderSide | undefined): string | undefined => {
    if (!s) return undefined;
    if (typeof s === 'object') return s.color;
    return undefined;
  };

  const hydrateFromActive = (): void => {
    const state = store.getState();
    const fmt = state.format.formats.get(addrKey(state.selection.active)) ?? {};

    if (
      fmt.numFmt &&
      (fmt.numFmt.kind === 'fixed' ||
        fmt.numFmt.kind === 'currency' ||
        fmt.numFmt.kind === 'percent')
    ) {
      draft.numFmt = fmt.numFmt;
      draft.numberCategory = fmt.numFmt.kind;
      if (fmt.numFmt.kind === 'fixed') {
        draft.decimals = fmt.numFmt.decimals;
      } else if (fmt.numFmt.kind === 'currency') {
        draft.decimals = fmt.numFmt.decimals;
        draft.currencySymbol = fmt.numFmt.symbol ?? '$';
      } else if (fmt.numFmt.kind === 'percent') {
        draft.decimals = fmt.numFmt.decimals;
      }
    } else {
      draft.numFmt = { kind: 'general' };
      draft.numberCategory = 'general';
      draft.decimals = 2;
      draft.currencySymbol = '$';
    }

    draft.align = fmt.align;
    draft.vAlign = fmt.vAlign;
    draft.wrap = !!fmt.wrap;
    draft.indent = fmt.indent ?? 0;
    draft.rotation = fmt.rotation ?? 0;
    draft.bold = !!fmt.bold;
    draft.italic = !!fmt.italic;
    draft.underline = !!fmt.underline;
    draft.strike = !!fmt.strike;
    draft.fontFamily = fmt.fontFamily ?? '';
    draft.fontSize = fmt.fontSize;
    draft.color = fmt.color;
    draft.fill = fmt.fill;
    draft.borders = { ...(fmt.borders ?? {}) };

    // Inherit "active" style/color from the first existing side, if any.
    const sides: SideKey[] = ['top', 'right', 'bottom', 'left', 'diagonalDown', 'diagonalUp'];
    let inheritedStyle: BorderStyleKey | null = null;
    let inheritedColor: string | undefined;
    for (const k of sides) {
      const s = draft.borders[k];
      const ss = sideStyle(s);
      if (ss && !inheritedStyle) inheritedStyle = ss;
      const cc = sideColor(s);
      if (cc && !inheritedColor) inheritedColor = cc;
    }
    draft.borderStyle = inheritedStyle ?? 'thin';
    draft.borderColor = inheritedColor;

    draft.hyperlink = fmt.hyperlink ?? '';
    draft.comment = fmt.comment ?? '';
    draft.validationList = fmt.validation?.kind === 'list' ? fmt.validation.source.join('\n') : '';

    syncControlsFromDraft();
    renderPreview();
    setActiveTab('number');
  };

  const syncControlsFromDraft = (): void => {
    // Number
    for (const [id, btn] of catButtons) {
      btn.setAttribute('aria-selected', id === draft.numberCategory ? 'true' : 'false');
    }
    decimalsInput.value = String(draft.decimals);
    symbolSelect.value = draft.currencySymbol;
    syncNumberControlsVisibility();

    // Alignment
    const hKey: 'default' | CellAlign = draft.align ?? 'default';
    for (const [id, r] of hAlignRadios) r.checked = id === hKey;
    const vKey: 'default' | CellVAlign = draft.vAlign ?? 'default';
    for (const [id, r] of vAlignRadios) r.checked = id === vKey;
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

    // More
    hlInput.value = draft.hyperlink;
    commentArea.value = draft.comment;
    validationArea.value = draft.validationList;
  };

  const syncNumberControlsVisibility = (): void => {
    const cat = draft.numberCategory;
    decimalsRow.hidden = cat === 'general';
    symbolRow.hidden = cat !== 'currency';
  };

  // ── Border helpers ─────────────────────────────────────────────────────
  /** Build a CellBorderSide value from the active style/color. */
  const activeSide = (): CellBorderSide => ({
    style: draft.borderStyle,
    ...(draft.borderColor ? { color: draft.borderColor } : {}),
  });

  /** Set or clear a specific side of the draft borders. We keep `false`
   *  markers (rather than deleting the key) so the eventual merge in
   *  setRangeFormat actually overrides any inherited side. */
  const setSide = (key: SideKey, on: boolean): void => {
    const next: CellBorders = { ...draft.borders };
    if (on) next[key] = activeSide();
    else next[key] = false;
    draft.borders = next;
  };

  /** Re-apply the active style/color to every currently-set side. Lets users
   *  flip the style and have existing sides update without re-toggling. */
  const restyleExistingSides = (): void => {
    const next: CellBorders = {};
    const sides: SideKey[] = ['top', 'right', 'bottom', 'left', 'diagonalDown', 'diagonalUp'];
    for (const k of sides) {
      if (draft.borders[k]) next[k] = activeSide();
    }
    draft.borders = next;
  };

  // ── Preview rendering ──────────────────────────────────────────────────
  const renderPreview = (): void => {
    preview.style.fontWeight = draft.bold ? 'bold' : 'normal';
    preview.style.fontStyle = draft.italic ? 'italic' : 'normal';
    const decos: string[] = [];
    if (draft.underline) decos.push('underline');
    if (draft.strike) decos.push('line-through');
    preview.style.textDecoration = decos.length > 0 ? decos.join(' ') : 'none';
    preview.style.textAlign = draft.align ?? 'left';
    preview.style.fontFamily = draft.fontFamily || '';
    preview.style.fontSize = draft.fontSize !== undefined ? `${draft.fontSize}px` : '';
    preview.style.color = draft.color ?? '';
    preview.style.backgroundColor = draft.fill ?? '';
    preview.style.whiteSpace = draft.wrap ? 'pre-wrap' : 'nowrap';
    const cssBorder = (s: CellBorderSide | undefined): string => {
      if (!s) return '';
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
    preview.style.borderTop = cssBorder(draft.borders.top);
    preview.style.borderRight = cssBorder(draft.borders.right);
    preview.style.borderBottom = cssBorder(draft.borders.bottom);
    preview.style.borderLeft = cssBorder(draft.borders.left);

    const numFmt = computeNumFmt();
    const numericText = formatNumber(12345, numFmt);
    preview.textContent = `${t.preview} ${numericText}`;
  };

  // ── Compute helpers ────────────────────────────────────────────────────
  const computeNumFmt = (): NumFmt => {
    const cat = draft.numberCategory;
    if (cat === 'general') return { kind: 'general' };
    if (cat === 'fixed') return { kind: 'fixed', decimals: draft.decimals };
    if (cat === 'currency') {
      return { kind: 'currency', decimals: draft.decimals, symbol: draft.currencySymbol };
    }
    return { kind: 'percent', decimals: draft.decimals };
  };

  // ── Tab switch ─────────────────────────────────────────────────────────
  const setActiveTab = (id: TabId): void => {
    activeTab = id;
    for (const [tabId, btn] of tabButtons) {
      btn.setAttribute('aria-selected', tabId === id ? 'true' : 'false');
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

    // Always send all six side keys so setRangeFormat's merge fully replaces
    // any inherited border instead of leaving stale sides intact.
    const explicitBorders: CellBorders = {
      top: draft.borders.top ?? false,
      right: draft.borders.right ?? false,
      bottom: draft.borders.bottom ?? false,
      left: draft.borders.left ?? false,
      diagonalDown: draft.borders.diagonalDown ?? false,
      diagonalUp: draft.borders.diagonalUp ?? false,
    };

    const patch: Partial<CellFormat> = {
      numFmt: computeNumFmt(),
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
      validation:
        validationLines.length > 0 ? { kind: 'list', source: validationLines } : undefined,
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
    if (id) setActiveTab(id);
  };

  const onCatClick = (e: MouseEvent): void => {
    const target = e.target as HTMLElement;
    const btn = target.closest('button[data-fc-cat]') as HTMLButtonElement | null;
    if (!btn) return;
    const id = btn.dataset.fcCat as NumberCategory | undefined;
    if (!id) return;
    draft.numberCategory = id;
    syncControlsFromDraft();
    renderPreview();
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

  const onHAlignChange = (e: Event): void => {
    const r = e.target as HTMLInputElement;
    if (!r.checked) return;
    draft.align = r.value === 'default' ? undefined : (r.value as CellAlign);
    renderPreview();
  };
  const onVAlignChange = (e: Event): void => {
    const r = e.target as HTMLInputElement;
    if (!r.checked) return;
    draft.vAlign = r.value === 'default' ? undefined : (r.value as CellVAlign);
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

  // Border events
  const onBorderStyleChange = (): void => {
    draft.borderStyle = borderStyleSelect.value as BorderStyleKey;
    restyleExistingSides();
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

  const onFillInput = (): void => {
    draft.fill = fillInput.value;
    renderPreview();
  };
  const onFillReset = (): void => {
    draft.fill = undefined;
    renderPreview();
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

  const onOk = (): void => applyAndClose();
  const onCancel = (): void => api.close();

  const onOverlayClick = (e: MouseEvent): void => {
    if (e.target === overlay) api.close();
  };

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
  tabsStrip.addEventListener('click', onTabClick);
  catList.addEventListener('click', onCatClick);
  decimalsInput.addEventListener('input', onDecimalsInput);
  symbolSelect.addEventListener('change', onSymbolChange);
  for (const r of hAlignRadios.values()) r.addEventListener('change', onHAlignChange);
  for (const r of vAlignRadios.values()) r.addEventListener('change', onVAlignChange);
  wrapCk.input.addEventListener('change', onWrapChange);
  indentInput.addEventListener('input', onIndentInput);
  rotationInput.addEventListener('input', onRotationInput);
  boldCk.input.addEventListener('change', onBoldChange);
  italicCk.input.addEventListener('change', onItalicChange);
  underlineCk.input.addEventListener('change', onUnderlineChange);
  strikeCk.input.addEventListener('change', onStrikeChange);
  familyInput.addEventListener('input', onFamilyInput);
  sizeInput.addEventListener('input', onSizeInput);
  colorInput.addEventListener('input', onColorInput);
  colorReset.addEventListener('click', onColorReset);
  borderStyleSelect.addEventListener('change', onBorderStyleChange);
  borderColorInput.addEventListener('input', onBorderColorInput);
  borderColorReset.addEventListener('click', onBorderColorReset);
  presetNone.addEventListener('click', onPresetNone);
  presetOutline.addEventListener('click', onPresetOutline);
  presetAll.addEventListener('click', onPresetAll);
  topCk.input.addEventListener('change', onTopChange);
  bottomCk.input.addEventListener('change', onBottomChange);
  leftCk.input.addEventListener('change', onLeftChange);
  rightCk.input.addEventListener('change', onRightChange);
  diagDownCk.input.addEventListener('change', onDiagDownChange);
  diagUpCk.input.addEventListener('change', onDiagUpChange);
  fillInput.addEventListener('input', onFillInput);
  fillReset.addEventListener('click', onFillReset);
  hlInput.addEventListener('input', onHlInput);
  hlClear.addEventListener('click', onHlClear);
  commentArea.addEventListener('input', onCommentInput);
  commentClear.addEventListener('click', onCommentClear);
  validationArea.addEventListener('input', onValidationInput);
  validationClear.addEventListener('click', onValidationClear);
  okBtn.addEventListener('click', onOk);
  cancelBtn.addEventListener('click', onCancel);
  overlay.addEventListener('click', onOverlayClick);
  overlay.addEventListener('keydown', onOverlayKey);

  const api: FormatDialogHandle = {
    open(): void {
      hydrateFromActive();
      overlay.hidden = false;
      requestAnimationFrame(() => {
        const first = tabButtons.get(activeTab);
        if (first) first.focus();
      });
    },
    close(): void {
      overlay.hidden = true;
      host.focus();
    },
    detach(): void {
      tabsStrip.removeEventListener('click', onTabClick);
      catList.removeEventListener('click', onCatClick);
      decimalsInput.removeEventListener('input', onDecimalsInput);
      symbolSelect.removeEventListener('change', onSymbolChange);
      for (const r of hAlignRadios.values()) r.removeEventListener('change', onHAlignChange);
      for (const r of vAlignRadios.values()) r.removeEventListener('change', onVAlignChange);
      wrapCk.input.removeEventListener('change', onWrapChange);
      indentInput.removeEventListener('input', onIndentInput);
      rotationInput.removeEventListener('input', onRotationInput);
      boldCk.input.removeEventListener('change', onBoldChange);
      italicCk.input.removeEventListener('change', onItalicChange);
      underlineCk.input.removeEventListener('change', onUnderlineChange);
      strikeCk.input.removeEventListener('change', onStrikeChange);
      familyInput.removeEventListener('input', onFamilyInput);
      sizeInput.removeEventListener('input', onSizeInput);
      colorInput.removeEventListener('input', onColorInput);
      colorReset.removeEventListener('click', onColorReset);
      borderStyleSelect.removeEventListener('change', onBorderStyleChange);
      borderColorInput.removeEventListener('input', onBorderColorInput);
      borderColorReset.removeEventListener('click', onBorderColorReset);
      presetNone.removeEventListener('click', onPresetNone);
      presetOutline.removeEventListener('click', onPresetOutline);
      presetAll.removeEventListener('click', onPresetAll);
      topCk.input.removeEventListener('change', onTopChange);
      bottomCk.input.removeEventListener('change', onBottomChange);
      leftCk.input.removeEventListener('change', onLeftChange);
      rightCk.input.removeEventListener('change', onRightChange);
      diagDownCk.input.removeEventListener('change', onDiagDownChange);
      diagUpCk.input.removeEventListener('change', onDiagUpChange);
      fillInput.removeEventListener('input', onFillInput);
      fillReset.removeEventListener('click', onFillReset);
      hlInput.removeEventListener('input', onHlInput);
      hlClear.removeEventListener('click', onHlClear);
      commentArea.removeEventListener('input', onCommentInput);
      commentClear.removeEventListener('click', onCommentClear);
      validationArea.removeEventListener('input', onValidationInput);
      validationClear.removeEventListener('click', onValidationClear);
      okBtn.removeEventListener('click', onOk);
      cancelBtn.removeEventListener('click', onCancel);
      overlay.removeEventListener('click', onOverlayClick);
      overlay.removeEventListener('keydown', onOverlayKey);
      overlay.remove();
    },
  };

  return api;
}

function isHexColor(s: string): boolean {
  return /^#[0-9a-fA-F]{6}$/.test(s);
}
