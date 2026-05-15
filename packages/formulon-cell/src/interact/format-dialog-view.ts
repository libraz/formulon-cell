import type { Strings } from '../i18n/strings.js';
import type { CellAlign, CellVAlign, ValidationOp } from '../store/store.js';
import { createDialogShell } from './dialog-shell.js';
import {
  makeButton,
  makeCheckbox,
  makeListSourceRadio,
  makeSection,
  makeSwatches,
  makeVisualSideButton,
} from './format-dialog-dom.js';
import {
  type BorderStyleKey,
  COMMON_FONTS,
  CURRENCY_SYMBOLS,
  type NumberCategory,
  type SideKey,
  type TabId,
  type ValidationKind,
} from './format-dialog-model.js';

interface CreateFormatDialogViewInput {
  host: HTMLElement;
  strings: Strings;
  t: Strings['formatDialog'];
}

export function createFormatDialogView(input: CreateFormatDialogViewInput) {
  const { host, strings, t } = input;
  const shell = createDialogShell({
    host,
    className: 'fc-fmtdlg',
    ariaLabel: t.title,
  });
  const { overlay, panel } = shell;

  // Header
  const header = document.createElement('div');
  header.className = 'fc-fmtdlg__header';
  header.textContent = t.title;
  panel.appendChild(header);

  // Preview
  const preview = document.createElement('div');
  preview.className = 'fc-fmtdlg__preview';
  const previewLabel = document.createElement('div');
  previewLabel.className = 'fc-fmtdlg__preview-label';
  previewLabel.textContent = t.preview;
  const previewCell = document.createElement('div');
  previewCell.className = 'fc-fmtdlg__preview-cell';
  preview.append(previewLabel, previewCell);

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
    { id: 'protection', label: strings.protection.tabProtection },
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

  // ── Number tab ─────────────────────────────────────────────────────────
  const numberPanel = tabPanels.get('number') as HTMLDivElement;
  const numberLayout = document.createElement('div');
  numberLayout.className = 'fc-fmtdlg__number-layout';
  numberPanel.appendChild(numberLayout);

  const catList = document.createElement('div');
  catList.className = 'fc-fmtdlg__cat';
  catList.setAttribute('role', 'listbox');
  // axe `aria-input-field-name`: a listbox needs an accessible name. Reuse
  // the Number tab label since this list IS the Number tab's category picker.
  catList.setAttribute('aria-label', t.tabNumber);
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

  const numberSummary = document.createElement('div');
  numberSummary.className = 'fc-fmtdlg__number-summary';
  const numberSummaryTitle = document.createElement('div');
  numberSummaryTitle.className = 'fc-fmtdlg__number-title';
  const numberSummaryDesc = document.createElement('p');
  numberSummaryDesc.className = 'fc-fmtdlg__number-desc';
  numberSummary.append(numberSummaryTitle, numberSummaryDesc);
  numberControls.appendChild(numberSummary);
  numberControls.appendChild(preview);

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

  const patternPresetRow = document.createElement('label');
  patternPresetRow.className = 'fc-fmtdlg__row';
  const patternPresetLabel = document.createElement('span');
  patternPresetLabel.textContent = t.patternType;
  const patternPresetSelect = document.createElement('select');
  patternPresetRow.append(patternPresetLabel, patternPresetSelect);
  numberControls.appendChild(patternPresetRow);

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

  // ── Alignment tab ──────────────────────────────────────────────────────
  const alignPanel = tabPanels.get('align') as HTMLDivElement;

  // Horizontal
  const hAlignLegend = document.createElement('div');
  hAlignLegend.textContent = t.horizontalAlign;
  alignPanel.appendChild(hAlignLegend);
  const hAlignFieldset = document.createElement('div');
  hAlignFieldset.className = 'fc-fmtdlg__choice-grid';
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
  vAlignFieldset.className = 'fc-fmtdlg__choice-grid';
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
  wrapRow.className = 'fc-fmtdlg__choice-grid';
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
  styleRow.className = 'fc-fmtdlg__choice-grid';
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
  const colorReset = makeButton(t.resetToDefault);
  colorRow.append(colorLabel, colorInput, colorReset);
  fontPanel.appendChild(colorRow);
  const fontSwatches = makeSwatches('font');
  fontPanel.appendChild(fontSwatches);

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

  const borderStyleGallery = document.createElement('div');
  borderStyleGallery.className = 'fc-fmtdlg__line-gallery';
  const borderStyleButtons = new Map<BorderStyleKey, HTMLButtonElement>();
  for (const s of styleOptions) {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = `fc-fmtdlg__line-style fc-fmtdlg__line-style--${s.id}`;
    btn.dataset.borderStyle = s.id;
    btn.setAttribute('aria-label', s.label);
    btn.setAttribute('aria-pressed', 'false');
    const sample = document.createElement('span');
    sample.className = 'fc-fmtdlg__line-sample';
    const label = document.createElement('span');
    label.textContent = s.label;
    btn.append(sample, label);
    borderStyleButtons.set(s.id, btn);
    borderStyleGallery.appendChild(btn);
  }
  borderPanel.appendChild(borderStyleGallery);

  const borderColorRow = document.createElement('div');
  borderColorRow.className = 'fc-fmtdlg__row';
  const borderColorLabel = document.createElement('span');
  borderColorLabel.textContent = t.borderColor;
  const borderColorInput = document.createElement('input');
  borderColorInput.type = 'color';
  borderColorInput.dataset.fcColor = 'border';
  const borderColorReset = makeButton(t.resetToDefault);
  borderColorRow.append(borderColorLabel, borderColorInput, borderColorReset);
  borderPanel.appendChild(borderColorRow);
  const borderSwatches = makeSwatches('border');
  borderPanel.appendChild(borderSwatches);

  // Presets
  const presetRow = document.createElement('div');
  presetRow.className = 'fc-fmtdlg__row';
  borderPanel.appendChild(presetRow);
  const presetNone = makeButton(t.borderPresetNone);
  const presetOutline = makeButton(t.borderPresetOutline);
  const presetAll = makeButton(t.borderPresetAll);
  presetRow.append(presetNone, presetOutline, presetAll);

  // Per-side checkboxes
  const sideRow = document.createElement('div');
  sideRow.className = 'fc-fmtdlg__row fc-fmtdlg__legacy-border-controls';
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
  diagonalRow.className = 'fc-fmtdlg__row fc-fmtdlg__legacy-border-controls';
  borderPanel.appendChild(diagonalRow);
  const diagDownCk = makeCheckbox(t.borderDiagonalDown);
  diagDownCk.input.dataset.fcCheck = 'borderDiagonalDown';
  const diagUpCk = makeCheckbox(t.borderDiagonalUp);
  diagUpCk.input.dataset.fcCheck = 'borderDiagonalUp';
  diagonalRow.append(diagDownCk.wrap, diagUpCk.wrap);

  const borderVisual = document.createElement('div');
  borderVisual.className = 'fc-fmtdlg__border-visual';
  const borderVisualTitle = document.createElement('div');
  borderVisualTitle.className = 'fc-fmtdlg__border-title';
  borderVisualTitle.textContent = t.preview;
  const borderVisualStage = document.createElement('div');
  borderVisualStage.className = 'fc-fmtdlg__border-stage';
  const borderVisualPreview = document.createElement('div');
  borderVisualPreview.className = 'fc-fmtdlg__border-preview';
  borderVisualPreview.textContent = t.previewText;
  const visualSideButtons = new Map<SideKey, HTMLButtonElement[]>();
  borderVisualStage.append(
    borderVisualPreview,
    makeVisualSideButton(visualSideButtons, 'top', t.borderTop),
    makeVisualSideButton(visualSideButtons, 'right', t.borderRight),
    makeVisualSideButton(visualSideButtons, 'bottom', t.borderBottom),
    makeVisualSideButton(visualSideButtons, 'left', t.borderLeft),
    makeVisualSideButton(
      visualSideButtons,
      'diagonalDown',
      t.borderDiagonalDown,
      ' fc-fmtdlg__border-hit--left-diag',
    ),
    makeVisualSideButton(
      visualSideButtons,
      'diagonalDown',
      t.borderDiagonalDown,
      ' fc-fmtdlg__border-hit--right-diag',
    ),
    makeVisualSideButton(
      visualSideButtons,
      'diagonalUp',
      t.borderDiagonalUp,
      ' fc-fmtdlg__border-hit--left-diag',
    ),
    makeVisualSideButton(
      visualSideButtons,
      'diagonalUp',
      t.borderDiagonalUp,
      ' fc-fmtdlg__border-hit--right-diag',
    ),
  );
  borderVisual.append(borderVisualTitle, borderVisualStage);
  borderPanel.appendChild(borderVisual);

  // ── Fill tab ───────────────────────────────────────────────────────────
  const fillPanel = tabPanels.get('fill') as HTMLDivElement;
  const fillSection = document.createElement('div');
  fillSection.className = 'fc-fmtdlg__section';
  const fillSectionTitle = document.createElement('div');
  fillSectionTitle.className = 'fc-fmtdlg__section-title';
  fillSectionTitle.textContent = t.fill;
  fillSection.appendChild(fillSectionTitle);
  fillPanel.appendChild(fillSection);
  const fillRow = document.createElement('div');
  fillRow.className = 'fc-fmtdlg__row';
  const fillLabel = document.createElement('span');
  fillLabel.textContent = t.fill;
  const fillInput = document.createElement('input');
  fillInput.type = 'color';
  fillInput.dataset.fcColor = 'fill';
  const fillReset = makeButton(t.fillNone);
  fillRow.append(fillLabel, fillInput, fillReset);
  fillSection.appendChild(fillRow);
  const fillSwatches = makeSwatches('fill');
  fillSection.appendChild(fillSwatches);

  // ── Protection tab ─────────────────────────────────────────────────────
  const protectionPanel = tabPanels.get('protection') as HTMLDivElement;
  const protectionSection = document.createElement('div');
  protectionSection.className = 'fc-fmtdlg__section';
  const protectionSectionTitle = document.createElement('div');
  protectionSectionTitle.className = 'fc-fmtdlg__section-title';
  protectionSectionTitle.textContent = strings.protection.tabProtection;
  protectionSection.appendChild(protectionSectionTitle);
  protectionPanel.appendChild(protectionSection);
  const lockedRow = document.createElement('div');
  lockedRow.className = 'fc-fmtdlg__row';
  const lockedCk = makeCheckbox(strings.protection.locked);
  lockedCk.input.dataset.fcCheck = 'locked';
  lockedRow.append(lockedCk.wrap);
  protectionSection.appendChild(lockedRow);
  const lockedHint = document.createElement('div');
  lockedHint.className = 'fc-fmtdlg__hint';
  lockedHint.textContent = strings.protection.lockedHint;
  protectionSection.appendChild(lockedHint);

  // ── More tab (hyperlink / comment / validation) ────────────────────────
  const morePanel = tabPanels.get('more') as HTMLDivElement;
  const hyperlinkSection = makeSection(t.hyperlink);
  const commentSection = makeSection(t.comment);
  const validationSection = makeSection(t.validationLegend);
  morePanel.append(hyperlinkSection, commentSection, validationSection);

  const hlRow = document.createElement('div');
  hlRow.className = 'fc-fmtdlg__row';
  hyperlinkSection.appendChild(hlRow);
  const hlLabel = document.createElement('span');
  hlLabel.textContent = t.hyperlink;
  const hlInput = document.createElement('input');
  hlInput.type = 'text';
  hlInput.spellcheck = false;
  hlInput.autocomplete = 'off';
  hlInput.placeholder = t.hyperlinkPlaceholder;
  const hlClear = makeButton(t.clearField);
  hlRow.append(hlLabel, hlInput, hlClear);

  const commentRow = document.createElement('div');
  commentRow.className = 'fc-fmtdlg__row fc-fmtdlg__row--block';
  commentSection.appendChild(commentRow);
  const commentArea = document.createElement('textarea');
  commentArea.className = 'fc-fmtdlg__textarea';
  commentArea.rows = 3;
  commentArea.placeholder = t.commentPlaceholder;
  const commentClear = makeButton(t.clearField);
  commentRow.append(commentArea, commentClear);

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
  validationSection.appendChild(validationKindRow);

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
  validationSection.appendChild(validationOpRow);

  const validationARow = document.createElement('label');
  validationARow.className = 'fc-fmtdlg__row';
  const validationALabel = document.createElement('span');
  validationALabel.textContent = t.validationValueA;
  const validationAInput = document.createElement('input');
  validationAInput.type = 'number';
  validationAInput.step = 'any';
  validationARow.append(validationALabel, validationAInput);
  validationSection.appendChild(validationARow);

  const validationBRow = document.createElement('label');
  validationBRow.className = 'fc-fmtdlg__row';
  const validationBLabel = document.createElement('span');
  validationBLabel.textContent = t.validationValueB;
  const validationBInput = document.createElement('input');
  validationBInput.type = 'number';
  validationBInput.step = 'any';
  validationBRow.append(validationBLabel, validationBInput);
  validationSection.appendChild(validationBRow);

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
  validationSection.appendChild(validationFormulaRow);

  // List source — visible only when kind === 'list'. The radio toggles between
  // an inline value list (textarea) and a range reference (single-line input).
  const validationListSourceKindRow = document.createElement('div');
  validationListSourceKindRow.className = 'fc-fmtdlg__row fc-fmtdlg__row--inline';
  const validationListSourceKindLabel = document.createElement('span');
  validationListSourceKindLabel.textContent = t.validationListSourceKind;
  validationListSourceKindRow.appendChild(validationListSourceKindLabel);
  const validationListLiteralRadio = makeListSourceRadio('literal', t.validationListSourceLiteral);
  const validationListRangeRadio = makeListSourceRadio('range', t.validationListSourceRange);
  validationListSourceKindRow.append(
    validationListLiteralRadio.wrap,
    validationListRangeRadio.wrap,
  );
  validationSection.appendChild(validationListSourceKindRow);

  const validationRow = document.createElement('div');
  validationRow.className = 'fc-fmtdlg__row fc-fmtdlg__row--block';
  const validationArea = document.createElement('textarea');
  validationArea.className = 'fc-fmtdlg__textarea';
  validationArea.rows = 4;
  validationArea.placeholder = t.validationListPlaceholder;
  const validationClear = makeButton(t.clearField);
  validationRow.append(validationArea, validationClear);
  validationSection.appendChild(validationRow);

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
  validationSection.appendChild(validationListRangeRow);

  // Allow-blank checkbox + error-style selector.
  const validationAllowBlankRow = document.createElement('label');
  validationAllowBlankRow.className = 'fc-fmtdlg__check';
  const validationAllowBlankInput = document.createElement('input');
  validationAllowBlankInput.type = 'checkbox';
  const validationAllowBlankSpan = document.createElement('span');
  validationAllowBlankSpan.textContent = t.validationAllowBlank;
  validationAllowBlankRow.append(validationAllowBlankInput, validationAllowBlankSpan);
  validationSection.appendChild(validationAllowBlankRow);

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
  validationSection.appendChild(validationErrorStyleRow);

  // ── Footer buttons ─────────────────────────────────────────────────────
  const cancelBtn = makeButton(t.cancel);
  const okBtn = makeButton(t.ok, true);
  footer.append(cancelBtn, okBtn);

  return {
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
    numberSummaryTitle,
    numberSummaryDesc,
    hAlignRadios,
    vAlignRadios,
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
  };
}
