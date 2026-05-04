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
  type CellValidation,
  type NumFmt,
  type SpreadsheetStore,
  type ValidationErrorStyle,
  type ValidationOp,
  mutators,
} from '../store/store.js';

/** Discriminator for the dialog's "kind" dropdown. `none` means clear the
 *  validation; the rest mirror `CellValidation['kind']`. */
type ValidationKind = 'none' | CellValidation['kind'];

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
type NumberCategory =
  | 'general'
  | 'fixed'
  | 'currency'
  | 'percent'
  | 'scientific'
  | 'accounting'
  | 'date'
  | 'time'
  | 'datetime'
  | 'text'
  | 'custom';
type BorderStyleKey = 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double';
type SideKey = 'top' | 'right' | 'bottom' | 'left' | 'diagonalDown' | 'diagonalUp';

interface DraftState {
  numFmt: NumFmt | undefined;
  numberCategory: NumberCategory;
  decimals: number;
  currencySymbol: string;
  /** Pattern for date/time/datetime/custom categories. */
  pattern: string;
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
  /** When kind === 'list', selects between inline string array and a range
   *  reference (Excel-style `Sheet1!$A$1:$A$10`). */
  validationListSourceKind: 'literal' | 'range';
  validationListRange: string;
  validationKind: ValidationKind;
  validationOp: ValidationOp;
  validationA: number;
  validationB: number;
  validationFormula: string;
  validationAllowBlank: boolean;
  validationErrorStyle: ValidationErrorStyle;
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
    { id: 'accounting', label: t.catAccounting },
    { id: 'percent', label: t.catPercent },
    { id: 'scientific', label: t.catScientific },
    { id: 'date', label: t.catDate },
    { id: 'time', label: t.catTime },
    { id: 'datetime', label: t.catDateTime },
    { id: 'text', label: t.catText },
    { id: 'custom', label: t.catCustom },
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

  // Pattern row — visible for date/time/datetime/custom categories.
  const patternRow = document.createElement('label');
  patternRow.className = 'fc-fmtdlg__row';
  const patternLabel = document.createElement('span');
  patternLabel.textContent = t.pattern;
  const patternInput = document.createElement('input');
  patternInput.type = 'text';
  patternInput.dataset.fcInput = 'pattern';
  patternInput.spellcheck = false;
  patternInput.autocomplete = 'off';
  patternInput.placeholder = t.patternPlaceholder;
  patternRow.append(patternLabel, patternInput);
  numberControls.appendChild(patternRow);

  // Live preview readout — shows the demo number after applying the active
  //  category. Excel's number tab shows this in the "Sample" header.
  const samplePreview = document.createElement('div');
  samplePreview.className = 'fc-fmtdlg__sample';
  numberControls.appendChild(samplePreview);

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
  familyInput.dataset.fcInput = 'family';
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
  validationLegend.textContent = t.validationLegend;
  morePanel.appendChild(validationLegend);

  // Kind selector — drives the visibility of the bound/formula/list rows.
  const validationKindRow = document.createElement('label');
  validationKindRow.className = 'fc-fmtdlg__row';
  const validationKindLabel = document.createElement('span');
  validationKindLabel.textContent = t.validationKind;
  const validationKindSelect = document.createElement('select');
  const kindDefs: { id: ValidationKind; label: string }[] = [
    { id: 'none', label: t.validationKindNone },
    { id: 'list', label: t.validationKindList },
    { id: 'whole', label: t.validationKindWhole },
    { id: 'decimal', label: t.validationKindDecimal },
    { id: 'date', label: t.validationKindDate },
    { id: 'time', label: t.validationKindTime },
    { id: 'textLength', label: t.validationKindTextLength },
    { id: 'custom', label: t.validationKindCustom },
  ];
  for (const k of kindDefs) {
    const opt = document.createElement('option');
    opt.value = k.id;
    opt.textContent = k.label;
    validationKindSelect.appendChild(opt);
  }
  validationKindRow.append(validationKindLabel, validationKindSelect);
  morePanel.appendChild(validationKindRow);

  // Op selector — visible for whole/decimal/date/time/textLength.
  const validationOpRow = document.createElement('label');
  validationOpRow.className = 'fc-fmtdlg__row';
  const validationOpLabel = document.createElement('span');
  validationOpLabel.textContent = t.validationOp;
  const validationOpSelect = document.createElement('select');
  const opDefs: { id: ValidationOp; label: string }[] = [
    { id: 'between', label: t.validationOpBetween },
    { id: 'notBetween', label: t.validationOpNotBetween },
    { id: '=', label: t.validationOpEq },
    { id: '<>', label: t.validationOpNeq },
    { id: '<', label: t.validationOpLt },
    { id: '<=', label: t.validationOpLte },
    { id: '>', label: t.validationOpGt },
    { id: '>=', label: t.validationOpGte },
  ];
  for (const o of opDefs) {
    const opt = document.createElement('option');
    opt.value = o.id;
    opt.textContent = o.label;
    validationOpSelect.appendChild(opt);
  }
  validationOpRow.append(validationOpLabel, validationOpSelect);
  morePanel.appendChild(validationOpRow);

  const validationARow = document.createElement('label');
  validationARow.className = 'fc-fmtdlg__row';
  const validationALabel = document.createElement('span');
  validationALabel.textContent = t.validationValueA;
  const validationAInput = document.createElement('input');
  validationAInput.type = 'number';
  validationAInput.step = 'any';
  validationARow.append(validationALabel, validationAInput);
  morePanel.appendChild(validationARow);

  const validationBRow = document.createElement('label');
  validationBRow.className = 'fc-fmtdlg__row';
  const validationBLabel = document.createElement('span');
  validationBLabel.textContent = t.validationValueB;
  const validationBInput = document.createElement('input');
  validationBInput.type = 'number';
  validationBInput.step = 'any';
  validationBRow.append(validationBLabel, validationBInput);
  morePanel.appendChild(validationBRow);

  // Custom-kind formula.
  const validationFormulaRow = document.createElement('label');
  validationFormulaRow.className = 'fc-fmtdlg__row';
  const validationFormulaLabel = document.createElement('span');
  validationFormulaLabel.textContent = t.validationFormula;
  const validationFormulaInput = document.createElement('input');
  validationFormulaInput.type = 'text';
  validationFormulaInput.spellcheck = false;
  validationFormulaInput.autocomplete = 'off';
  validationFormulaInput.placeholder = t.validationFormulaPlaceholder;
  validationFormulaRow.append(validationFormulaLabel, validationFormulaInput);
  morePanel.appendChild(validationFormulaRow);

  // List source — visible only when kind === 'list'. The radio toggles between
  // an inline value list (textarea) and a range reference (single-line input).
  const validationListSourceKindRow = document.createElement('div');
  validationListSourceKindRow.className = 'fc-fmtdlg__row fc-fmtdlg__row--inline';
  const validationListSourceKindLabel = document.createElement('span');
  validationListSourceKindLabel.textContent = t.validationListSourceKind;
  validationListSourceKindRow.appendChild(validationListSourceKindLabel);
  const makeListSourceRadio = (
    value: 'literal' | 'range',
    label: string,
  ): { wrap: HTMLLabelElement; input: HTMLInputElement } => {
    const wrap = document.createElement('label');
    wrap.className = 'fc-fmtdlg__check';
    const input = document.createElement('input');
    input.type = 'radio';
    input.name = 'fc-validation-list-source';
    input.value = value;
    const span = document.createElement('span');
    span.textContent = label;
    wrap.append(input, span);
    return { wrap, input };
  };
  const validationListLiteralRadio = makeListSourceRadio('literal', t.validationListSourceLiteral);
  const validationListRangeRadio = makeListSourceRadio('range', t.validationListSourceRange);
  validationListSourceKindRow.append(
    validationListLiteralRadio.wrap,
    validationListRangeRadio.wrap,
  );
  morePanel.appendChild(validationListSourceKindRow);

  const validationRow = document.createElement('div');
  validationRow.className = 'fc-fmtdlg__row fc-fmtdlg__row--block';
  morePanel.appendChild(validationRow);
  const validationArea = document.createElement('textarea');
  validationArea.className = 'fc-fmtdlg__textarea';
  validationArea.rows = 4;
  validationArea.placeholder = t.validationListPlaceholder;
  const validationClear = makeBtn(t.clearField);
  validationRow.append(validationArea, validationClear);

  // Range-ref input. Hidden unless source kind === 'range'.
  const validationListRangeRow = document.createElement('label');
  validationListRangeRow.className = 'fc-fmtdlg__row';
  const validationListRangeLabel = document.createElement('span');
  validationListRangeLabel.textContent = t.validationListSourceRange;
  const validationListRangeInput = document.createElement('input');
  validationListRangeInput.type = 'text';
  validationListRangeInput.spellcheck = false;
  validationListRangeInput.autocomplete = 'off';
  validationListRangeInput.placeholder = t.validationListRangePlaceholder;
  validationListRangeRow.append(validationListRangeLabel, validationListRangeInput);
  morePanel.appendChild(validationListRangeRow);

  // Allow-blank checkbox + error-style selector.
  const validationAllowBlankRow = document.createElement('label');
  validationAllowBlankRow.className = 'fc-fmtdlg__check';
  const validationAllowBlankInput = document.createElement('input');
  validationAllowBlankInput.type = 'checkbox';
  const validationAllowBlankSpan = document.createElement('span');
  validationAllowBlankSpan.textContent = t.validationAllowBlank;
  validationAllowBlankRow.append(validationAllowBlankInput, validationAllowBlankSpan);
  morePanel.appendChild(validationAllowBlankRow);

  const validationErrorStyleRow = document.createElement('label');
  validationErrorStyleRow.className = 'fc-fmtdlg__row';
  const validationErrorStyleLabel = document.createElement('span');
  validationErrorStyleLabel.textContent = t.validationErrorStyle;
  const validationErrorStyleSelect = document.createElement('select');
  for (const e of [
    { id: 'stop' as const, label: t.validationErrorStop },
    { id: 'warning' as const, label: t.validationErrorWarning },
    { id: 'information' as const, label: t.validationErrorInfo },
  ]) {
    const opt = document.createElement('option');
    opt.value = e.id;
    opt.textContent = e.label;
    validationErrorStyleSelect.appendChild(opt);
  }
  validationErrorStyleRow.append(validationErrorStyleLabel, validationErrorStyleSelect);
  morePanel.appendChild(validationErrorStyleRow);

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
      pattern: '',
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
      validationListSourceKind: 'literal',
      validationListRange: '',
      validationKind: 'none',
      validationOp: 'between',
      validationA: 0,
      validationB: 0,
      validationFormula: '',
      validationAllowBlank: true,
      validationErrorStyle: 'stop',
    };
  }

  // ── Hydration ──────────────────────────────────────────────────────────
  const sideStyle = (s: CellBorderSide | undefined): BorderStyleKey | null => {
    if (!s) return null;
    if (typeof s === 'object') {
      // The dialog only exposes the basic 6 style keys; OOXML's extended
      //  styles (hair / mediumDashed / dashDot variants) collapse to the
      //  closest visual cousin so the dropdown still lights up correctly.
      switch (s.style) {
        case 'thin':
        case 'medium':
        case 'thick':
        case 'dashed':
        case 'dotted':
        case 'double':
          return s.style;
        case 'hair':
          return 'thin';
        case 'mediumDashed':
        case 'dashDot':
        case 'mediumDashDot':
        case 'dashDotDot':
        case 'mediumDashDotDot':
        case 'slantDashDot':
          return 'dashed';
        default:
          return 'thin';
      }
    }
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

    if (fmt.numFmt) {
      draft.numFmt = fmt.numFmt;
      switch (fmt.numFmt.kind) {
        case 'fixed':
          draft.numberCategory = 'fixed';
          draft.decimals = fmt.numFmt.decimals;
          break;
        case 'currency':
          draft.numberCategory = 'currency';
          draft.decimals = fmt.numFmt.decimals;
          draft.currencySymbol = fmt.numFmt.symbol ?? '$';
          break;
        case 'percent':
          draft.numberCategory = 'percent';
          draft.decimals = fmt.numFmt.decimals;
          break;
        case 'scientific':
          draft.numberCategory = 'scientific';
          draft.decimals = fmt.numFmt.decimals;
          break;
        case 'accounting':
          draft.numberCategory = 'accounting';
          draft.decimals = fmt.numFmt.decimals;
          draft.currencySymbol = fmt.numFmt.symbol ?? '$';
          break;
        case 'date':
          draft.numberCategory = 'date';
          draft.pattern = fmt.numFmt.pattern;
          break;
        case 'time':
          draft.numberCategory = 'time';
          draft.pattern = fmt.numFmt.pattern;
          break;
        case 'datetime':
          draft.numberCategory = 'datetime';
          draft.pattern = fmt.numFmt.pattern;
          break;
        case 'text':
          draft.numberCategory = 'text';
          break;
        case 'custom':
          draft.numberCategory = 'custom';
          draft.pattern = fmt.numFmt.pattern;
          break;
        default:
          draft.numberCategory = 'general';
      }
    } else {
      draft.numFmt = { kind: 'general' };
      draft.numberCategory = 'general';
      draft.decimals = 2;
      draft.currencySymbol = '$';
      draft.pattern = '';
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
    // Hydrate the validation block — every kind gets its own field set.
    const v = fmt.validation;
    if (v) {
      draft.validationKind = v.kind;
      draft.validationAllowBlank = v.allowBlank !== false;
      draft.validationErrorStyle = v.errorStyle ?? 'stop';
      if (v.kind === 'list') {
        if (Array.isArray(v.source)) {
          draft.validationListSourceKind = 'literal';
          draft.validationList = v.source.join('\n');
          draft.validationListRange = '';
        } else {
          draft.validationListSourceKind = 'range';
          draft.validationList = '';
          draft.validationListRange = v.source.ref;
        }
      } else {
        draft.validationList = '';
        draft.validationListRange = '';
        draft.validationListSourceKind = 'literal';
      }
      if (v.kind === 'custom') {
        draft.validationFormula = v.formula;
      } else {
        draft.validationFormula = '';
      }
      if (
        v.kind === 'whole' ||
        v.kind === 'decimal' ||
        v.kind === 'date' ||
        v.kind === 'time' ||
        v.kind === 'textLength'
      ) {
        draft.validationOp = v.op;
        draft.validationA = v.a;
        draft.validationB = v.b ?? v.a;
      } else {
        draft.validationOp = 'between';
        draft.validationA = 0;
        draft.validationB = 0;
      }
    } else {
      draft.validationKind = 'none';
      draft.validationList = '';
      draft.validationListRange = '';
      draft.validationListSourceKind = 'literal';
      draft.validationFormula = '';
      draft.validationOp = 'between';
      draft.validationA = 0;
      draft.validationB = 0;
      draft.validationAllowBlank = true;
      draft.validationErrorStyle = 'stop';
    }

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
    patternInput.value = draft.pattern;
    if (!draft.pattern) {
      patternInput.placeholder = defaultPatternFor(draft.numberCategory) || t.patternPlaceholder;
    } else {
      patternInput.placeholder = t.patternPlaceholder;
    }
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
    const decimalsCats = new Set<NumberCategory>([
      'fixed',
      'currency',
      'percent',
      'scientific',
      'accounting',
    ]);
    const symbolCats = new Set<NumberCategory>(['currency', 'accounting']);
    const patternCats = new Set<NumberCategory>(['date', 'time', 'datetime', 'custom']);
    decimalsRow.hidden = !decimalsCats.has(cat);
    symbolRow.hidden = !symbolCats.has(cat);
    patternRow.hidden = !patternCats.has(cat);
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
    // Pick a sample value that exercises the active category. Date/time
    //  categories use a serial near the present (45123 ≈ 2023-07-16).
    const isDateLike =
      numFmt.kind === 'date' || numFmt.kind === 'time' || numFmt.kind === 'datetime';
    const sampleValue = isDateLike ? 45123.625 : 12345;
    const numericText = formatNumber(sampleValue, numFmt);
    preview.textContent = `${t.preview} ${numericText}`;
    samplePreview.textContent = numericText;
  };

  // ── Compute helpers ────────────────────────────────────────────────────
  const defaultPatternFor = (cat: NumberCategory): string => {
    switch (cat) {
      case 'date':
        return 'yyyy-mm-dd';
      case 'time':
        return 'HH:MM:SS';
      case 'datetime':
        return 'yyyy-mm-dd HH:MM';
      case 'custom':
        return '0.00';
      default:
        return '';
    }
  };
  const computeNumFmt = (): NumFmt => {
    const cat = draft.numberCategory;
    switch (cat) {
      case 'general':
        return { kind: 'general' };
      case 'fixed':
        return { kind: 'fixed', decimals: draft.decimals };
      case 'currency':
        return { kind: 'currency', decimals: draft.decimals, symbol: draft.currencySymbol };
      case 'percent':
        return { kind: 'percent', decimals: draft.decimals };
      case 'scientific':
        return { kind: 'scientific', decimals: draft.decimals };
      case 'accounting':
        return { kind: 'accounting', decimals: draft.decimals, symbol: draft.currencySymbol };
      case 'text':
        return { kind: 'text' };
      case 'date':
        return { kind: 'date', pattern: draft.pattern || defaultPatternFor('date') };
      case 'time':
        return { kind: 'time', pattern: draft.pattern || defaultPatternFor('time') };
      case 'datetime':
        return { kind: 'datetime', pattern: draft.pattern || defaultPatternFor('datetime') };
      case 'custom':
        return { kind: 'custom', pattern: draft.pattern || defaultPatternFor('custom') };
    }
  };

  /** Build a CellValidation from the draft's validation block, or undefined
   *  when the kind is `none` (clear). The list-source textarea is collapsed
   *  to a non-empty array; empty lists yield `none`. */
  const computeValidation = (lines: string[]): CellValidation | undefined => {
    const k = draft.validationKind;
    if (k === 'none') return undefined;
    const meta = {
      ...(draft.validationAllowBlank ? {} : { allowBlank: false }),
      ...(draft.validationErrorStyle !== 'stop' ? { errorStyle: draft.validationErrorStyle } : {}),
    };
    switch (k) {
      case 'list':
        if (draft.validationListSourceKind === 'range') {
          const ref = draft.validationListRange.trim().replace(/^=/, '');
          if (!ref) return undefined;
          return { kind: 'list', source: { ref }, ...meta };
        }
        if (lines.length === 0) return undefined;
        return { kind: 'list', source: lines, ...meta };
      case 'custom': {
        const formula = draft.validationFormula.trim();
        if (!formula) return undefined;
        return { kind: 'custom', formula, ...meta };
      }
      case 'whole':
      case 'decimal':
      case 'date':
      case 'time':
      case 'textLength': {
        const op = draft.validationOp;
        const a = draft.validationA;
        if (op === 'between' || op === 'notBetween') {
          return { kind: k, op, a, b: draft.validationB, ...meta };
        }
        return { kind: k, op, a, ...meta };
      }
    }
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

    const validation = computeValidation(validationLines);

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
      validation,
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

  const onPatternInput = (): void => {
    draft.pattern = patternInput.value;
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
  patternInput.addEventListener('input', onPatternInput);
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
  validationListRangeInput.addEventListener('input', onValidationListRangeInput);
  validationListLiteralRadio.input.addEventListener('change', onValidationListSourceKindChange);
  validationListRangeRadio.input.addEventListener('change', onValidationListSourceKindChange);
  validationKindSelect.addEventListener('change', onValidationKindChange);
  validationOpSelect.addEventListener('change', onValidationOpChange);
  validationAInput.addEventListener('input', onValidationAInput);
  validationBInput.addEventListener('input', onValidationBInput);
  validationFormulaInput.addEventListener('input', onValidationFormulaInput);
  validationAllowBlankInput.addEventListener('change', onValidationAllowBlankChange);
  validationErrorStyleSelect.addEventListener('change', onValidationErrorStyleChange);
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
      patternInput.removeEventListener('input', onPatternInput);
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
      validationListRangeInput.removeEventListener('input', onValidationListRangeInput);
      validationListLiteralRadio.input.removeEventListener(
        'change',
        onValidationListSourceKindChange,
      );
      validationListRangeRadio.input.removeEventListener(
        'change',
        onValidationListSourceKindChange,
      );
      validationKindSelect.removeEventListener('change', onValidationKindChange);
      validationOpSelect.removeEventListener('change', onValidationOpChange);
      validationAInput.removeEventListener('input', onValidationAInput);
      validationBInput.removeEventListener('input', onValidationBInput);
      validationFormulaInput.removeEventListener('input', onValidationFormulaInput);
      validationAllowBlankInput.removeEventListener('change', onValidationAllowBlankChange);
      validationErrorStyleSelect.removeEventListener('change', onValidationErrorStyleChange);
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
