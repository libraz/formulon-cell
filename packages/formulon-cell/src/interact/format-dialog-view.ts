import type { Strings } from '../i18n/strings.js';
import type {
  CellAlign,
  CellVAlign,
  FillPattern,
  NegativeStyle,
  TextDirection,
  ValidationOp,
} from '../store/store.js';
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
  const headerTitle = document.createElement('span');
  headerTitle.textContent = t.title;
  const closeBtn = document.createElement('button');
  closeBtn.type = 'button';
  closeBtn.className = 'fc-fmtdlg__close';
  closeBtn.setAttribute('aria-label', t.cancel);
  closeBtn.textContent = '×';
  header.append(headerTitle, closeBtn);
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
  tabsStrip.setAttribute('aria-label', t.title);
  panel.appendChild(tabsStrip);

  // Body
  const body = document.createElement('div');
  body.className = 'fc-fmtdlg__body';
  panel.appendChild(body);

  // Hint bar (full-width context description shown above the footer; only
  // populated for the Number tab — other tabs leave it empty/hidden).
  const hintBar = document.createElement('div');
  hintBar.className = 'fc-fmtdlg__hintbar';
  panel.appendChild(hintBar);

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
    btn.setAttribute('aria-controls', `fc-fmtdlg-panel-${def.id}`);
    btn.tabIndex = -1;
    btn.dataset.fcTab = def.id;
    tabsStrip.appendChild(btn);
    tabButtons.set(def.id, btn);

    const panelEl = document.createElement('div');
    panelEl.className = 'fc-fmtdlg__panel-tab';
    panelEl.id = `fc-fmtdlg-panel-${def.id}`;
    panelEl.setAttribute('role', 'tabpanel');
    panelEl.setAttribute('aria-labelledby', `fc-fmtdlg-tab-${def.id}`);
    panelEl.dataset.fcTab = def.id;
    panelEl.hidden = true;
    body.appendChild(panelEl);
    tabPanels.set(def.id, panelEl);
    btn.id = `fc-fmtdlg-tab-${def.id}`;
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
    { id: 'special', label: t.catOther },
    { id: 'custom', label: t.catCustom },
  ];
  const catButtons = new Map<NumberCategory, HTMLButtonElement>();
  for (const c of catDefs) {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'fc-fmtdlg__cat-item';
    b.textContent = c.label;
    b.setAttribute('role', 'option');
    b.tabIndex = -1;
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
  decimalsInput.setAttribute('aria-label', t.decimals);
  decimalsInput.min = '0';
  decimalsInput.max = '10';
  decimalsInput.step = '1';
  decimalsRow.append(decimalsLabel, decimalsInput);
  numberControls.appendChild(decimalsRow);

  const thousandsCk = makeCheckbox(t.thousandsSeparator);
  thousandsCk.input.setAttribute('aria-label', t.thousandsSeparator);
  thousandsCk.input.dataset.fcCheck = 'thousands';
  thousandsCk.wrap.classList.add('fc-fmtdlg__number-check');
  numberControls.appendChild(thousandsCk.wrap);

  const symbolRow = document.createElement('label');
  symbolRow.className = 'fc-fmtdlg__row';
  const symbolLabel = document.createElement('span');
  symbolLabel.textContent = t.symbol;
  const symbolSelect = document.createElement('select');
  symbolSelect.setAttribute('aria-label', t.symbol);
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
  patternPresetSelect.setAttribute('aria-label', t.patternType);
  patternPresetSelect.dataset.fcSelect = 'patternPreset';
  patternPresetRow.append(patternPresetLabel, patternPresetSelect);
  numberControls.appendChild(patternPresetRow);

  // Pattern listbox — Office365-style clickable list of formatted example
  // strings (shown for date/time/datetime/special). Replaces the select
  // dropdown above for those categories so users see the actual rendering
  // instead of cryptic pattern codes.
  const patternListWrap = document.createElement('div');
  patternListWrap.className = 'fc-fmtdlg__pattern-list-wrap';
  const patternListLabel = document.createElement('div');
  patternListLabel.className = 'fc-fmtdlg__pattern-list-label';
  patternListLabel.textContent = t.patternType;
  const patternList = document.createElement('div');
  patternList.className = 'fc-fmtdlg__pattern-list';
  patternList.setAttribute('role', 'listbox');
  patternList.setAttribute('aria-label', t.patternType);
  patternListWrap.append(patternListLabel, patternList);
  numberControls.appendChild(patternListWrap);

  // Pattern row — visible for date/time/datetime/custom categories.
  const patternRow = document.createElement('label');
  patternRow.className = 'fc-fmtdlg__row';
  const patternLabel = document.createElement('span');
  patternLabel.textContent = t.pattern;
  const patternInput = document.createElement('input');
  patternInput.type = 'text';
  patternInput.setAttribute('aria-label', t.pattern);
  patternInput.dataset.fcInput = 'pattern';
  patternInput.spellcheck = false;
  patternInput.autocomplete = 'off';
  patternInput.placeholder = t.patternPlaceholder;
  patternRow.append(patternLabel, patternInput);
  numberControls.appendChild(patternRow);

  const localeRow = document.createElement('label');
  localeRow.className = 'fc-fmtdlg__row';
  const localeLabel = document.createElement('span');
  localeLabel.textContent = t.languageLocation;
  const localeSelect = document.createElement('select');
  localeSelect.setAttribute('aria-label', t.languageLocation);
  for (const [value, label] of [
    ['ja', '日本語'],
    ['en', 'English'],
  ] as const) {
    const opt = document.createElement('option');
    opt.value = value;
    opt.textContent = label;
    localeSelect.appendChild(opt);
  }
  localeRow.append(localeLabel, localeSelect);
  numberControls.appendChild(localeRow);

  // Calendar type — visible only for the date category. Choices mirror
  // Office365's Gregorian/Japanese options; pattern output is unchanged
  // since downstream rendering does not yet implement era formatting.
  const calendarRow = document.createElement('label');
  calendarRow.className = 'fc-fmtdlg__row';
  const calendarLabel = document.createElement('span');
  calendarLabel.textContent = t.calendarType;
  const calendarSelect = document.createElement('select');
  calendarSelect.setAttribute('aria-label', t.calendarType);
  calendarSelect.dataset.fcSelect = 'calendarType';
  for (const [value, label] of [
    ['gregorian', t.calendarTypeGregorian],
    ['japanese', t.calendarTypeJapanese],
  ] as const) {
    const opt = document.createElement('option');
    opt.value = value;
    opt.textContent = label;
    calendarSelect.appendChild(opt);
  }
  calendarRow.append(calendarLabel, calendarSelect);
  numberControls.appendChild(calendarRow);

  const negativeList = document.createElement('div');
  negativeList.className = 'fc-fmtdlg__negative';
  const negativeLabel = document.createElement('div');
  negativeLabel.className = 'fc-fmtdlg__negative-label';
  negativeLabel.textContent = t.negativeNumbers;
  const negativeOptions = document.createElement('div');
  negativeOptions.className = 'fc-fmtdlg__negative-list';
  negativeOptions.setAttribute('role', 'listbox');
  negativeOptions.setAttribute('aria-label', t.negativeNumbers);
  const negativeSamples: Array<{ value: NegativeStyle; text: string; red?: boolean }> = [
    { value: 'parens', text: '(1234)' },
    { value: 'red-parens', text: '(1234)', red: true },
    { value: 'red', text: '-1234', red: true },
    { value: 'minus', text: '-1234' },
  ];
  for (const [index, sample] of negativeSamples.entries()) {
    const item = document.createElement('button');
    item.type = 'button';
    item.className = 'fc-fmtdlg__negative-item';
    item.setAttribute('role', 'option');
    item.setAttribute('aria-selected', index === 3 ? 'true' : 'false');
    item.dataset.fcNegativeStyle = sample.value;
    if (sample.red) item.classList.add('fc-fmtdlg__negative-item--red');
    item.textContent = sample.text;
    negativeOptions.appendChild(item);
  }
  negativeList.append(negativeLabel, negativeOptions);
  numberControls.appendChild(negativeList);

  // ── Alignment tab ──────────────────────────────────────────────────────
  const alignPanel = tabPanels.get('align') as HTMLDivElement;

  // Horizontal
  const hAlignLegend = document.createElement('div');
  hAlignLegend.textContent = t.horizontalAlign;
  alignPanel.appendChild(hAlignLegend);
  const hAlignFieldset = document.createElement('div');
  hAlignFieldset.className = 'fc-fmtdlg__choice-grid';
  hAlignFieldset.setAttribute('role', 'radiogroup');
  hAlignFieldset.setAttribute('aria-label', t.horizontalAlign);
  alignPanel.appendChild(hAlignFieldset);

  const hAlignName = `fc-fmtdlg-halign-${Math.random().toString(36).slice(2, 8)}`;
  const hAlignDefs: { id: 'default' | CellAlign; label: string }[] = [
    { id: 'default', label: t.alignDefault },
    { id: 'left', label: t.alignLeft },
    { id: 'center', label: t.alignCenter },
    { id: 'right', label: t.alignRight },
    { id: 'fill', label: t.alignFill },
    { id: 'justify', label: t.alignJustify },
    { id: 'centerContinuous', label: t.alignCenterAcrossSelection },
    { id: 'distributed', label: t.alignDistributed },
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

  const hAlignSelectRow = document.createElement('label');
  hAlignSelectRow.className = 'fc-fmtdlg__row fc-fmtdlg__align-select-row';
  const hAlignSelectLabel = document.createElement('span');
  hAlignSelectLabel.textContent = t.horizontalAlign;
  const hAlignSelect = document.createElement('select');
  hAlignSelect.setAttribute('aria-label', t.horizontalAlign);
  hAlignSelect.dataset.fcSelect = 'align';
  for (const a of hAlignDefs) {
    const opt = document.createElement('option');
    opt.value = a.id;
    opt.textContent = a.label;
    hAlignSelect.appendChild(opt);
  }
  hAlignSelectRow.append(hAlignSelectLabel, hAlignSelect);
  alignPanel.appendChild(hAlignSelectRow);

  // Vertical
  const vAlignLegend = document.createElement('div');
  vAlignLegend.textContent = t.verticalAlign;
  alignPanel.appendChild(vAlignLegend);
  const vAlignFieldset = document.createElement('div');
  vAlignFieldset.className = 'fc-fmtdlg__choice-grid';
  vAlignFieldset.setAttribute('role', 'radiogroup');
  vAlignFieldset.setAttribute('aria-label', t.verticalAlign);
  alignPanel.appendChild(vAlignFieldset);

  const vAlignName = `fc-fmtdlg-valign-${Math.random().toString(36).slice(2, 8)}`;
  const vAlignDefs: { id: 'default' | CellVAlign; label: string }[] = [
    { id: 'default', label: t.alignDefault },
    { id: 'top', label: t.vAlignTop },
    { id: 'middle', label: t.vAlignMiddle },
    { id: 'bottom', label: t.vAlignBottom },
    { id: 'justify', label: t.vAlignJustify },
    { id: 'distributed', label: t.vAlignDistributed },
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

  const vAlignSelectRow = document.createElement('label');
  vAlignSelectRow.className = 'fc-fmtdlg__row fc-fmtdlg__align-select-row';
  const vAlignSelectLabel = document.createElement('span');
  vAlignSelectLabel.textContent = t.verticalAlign;
  const vAlignSelect = document.createElement('select');
  vAlignSelect.setAttribute('aria-label', t.verticalAlign);
  vAlignSelect.dataset.fcSelect = 'vAlign';
  for (const a of vAlignDefs) {
    const opt = document.createElement('option');
    opt.value = a.id;
    opt.textContent = a.label;
    vAlignSelect.appendChild(opt);
  }
  vAlignSelectRow.append(vAlignSelectLabel, vAlignSelect);
  alignPanel.appendChild(vAlignSelectRow);

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
  indentInput.setAttribute('aria-label', t.indent);
  indentInput.min = '0';
  indentInput.max = '15';
  indentInput.step = '1';
  indentRow.append(indentLabel, indentInput);
  alignPanel.appendChild(indentRow);

  const textDirectionRow = document.createElement('label');
  textDirectionRow.className = 'fc-fmtdlg__row';
  const textDirectionLabel = document.createElement('span');
  textDirectionLabel.textContent = t.textDirection;
  const textDirectionSelect = document.createElement('select');
  textDirectionSelect.setAttribute('aria-label', t.textDirection);
  textDirectionSelect.dataset.fcSelect = 'textDirection';
  const directionDefs: Array<{ id: TextDirection; label: string }> = [
    { id: 'context', label: t.directionContext },
    { id: 'ltr', label: t.directionLeftToRight },
    { id: 'rtl', label: t.directionRightToLeft },
  ];
  for (const direction of directionDefs) {
    const opt = document.createElement('option');
    opt.value = direction.id;
    opt.textContent = direction.label;
    textDirectionSelect.appendChild(opt);
  }
  textDirectionRow.append(textDirectionLabel, textDirectionSelect);
  alignPanel.appendChild(textDirectionRow);

  const rotationRow = document.createElement('label');
  rotationRow.className = 'fc-fmtdlg__row fc-fmtdlg__rotation-row';
  const rotationLabel = document.createElement('span');
  rotationLabel.textContent = t.rotation;
  const rotationInput = document.createElement('input');
  rotationInput.type = 'number';
  rotationInput.setAttribute('aria-label', t.rotation);
  rotationInput.min = '-90';
  rotationInput.max = '90';
  rotationInput.step = '1';
  rotationRow.append(rotationLabel, rotationInput);
  alignPanel.appendChild(rotationRow);

  const alignPreview = document.createElement('div');
  alignPreview.className = 'fc-fmtdlg__align-preview';
  const alignPreviewTitle = document.createElement('div');
  alignPreviewTitle.className = 'fc-fmtdlg__align-preview-title';
  alignPreviewTitle.textContent = t.direction;
  const alignPreviewBox = document.createElement('div');
  alignPreviewBox.className = 'fc-fmtdlg__align-preview-box';
  const alignPreviewVertical = document.createElement('div');
  alignPreviewVertical.className = 'fc-fmtdlg__align-preview-vertical';
  alignPreviewVertical.textContent = t.previewText;
  const alignPreviewDial = document.createElement('div');
  alignPreviewDial.className = 'fc-fmtdlg__align-preview-dial';
  alignPreviewDial.setAttribute('role', 'group');
  alignPreviewDial.setAttribute('aria-label', t.rotation);
  const alignPreviewDialArc = document.createElement('span');
  alignPreviewDialArc.className = 'fc-fmtdlg__align-preview-arc';
  alignPreviewDialArc.setAttribute('aria-hidden', 'true');
  alignPreviewDial.appendChild(alignPreviewDialArc);
  const ANGLE_STEPS = [90, 75, 60, 45, 30, 15, 0, -15, -30, -45, -60, -75, -90];
  const alignPreviewDialDots: HTMLButtonElement[] = [];
  for (const angle of ANGLE_STEPS) {
    const dot = document.createElement('button');
    dot.type = 'button';
    dot.className = 'fc-fmtdlg__align-preview-dot';
    dot.dataset.fcAngle = String(angle);
    dot.setAttribute('aria-label', `${angle}°`);
    dot.title = `${angle}°`;
    const rad = (angle * Math.PI) / 180;
    const cx = 12;
    const cy = 66;
    const radius = 56;
    const px = cx + radius * Math.cos(rad);
    const py = cy - radius * Math.sin(rad);
    dot.style.left = `${px}px`;
    dot.style.top = `${py}px`;
    alignPreviewDial.appendChild(dot);
    alignPreviewDialDots.push(dot);
  }
  const alignPreviewDialPointer = document.createElement('span');
  alignPreviewDialPointer.className = 'fc-fmtdlg__align-preview-pointer';
  alignPreviewDialPointer.setAttribute('aria-hidden', 'true');
  alignPreviewDial.appendChild(alignPreviewDialPointer);
  const alignPreviewDialText = document.createElement('span');
  alignPreviewDialText.className = 'fc-fmtdlg__align-preview-text';
  alignPreviewDialText.textContent = t.previewText;
  alignPreviewDial.appendChild(alignPreviewDialText);
  alignPreviewBox.append(alignPreviewVertical, alignPreviewDial);
  const alignPreviewDegree = document.createElement('label');
  alignPreviewDegree.className = 'fc-fmtdlg__align-degree';
  const alignPreviewDegreeLabel = document.createElement('span');
  alignPreviewDegreeLabel.textContent = t.rotation.replace(/\s*\(.*\)$/, '');
  alignPreviewDegree.append(alignPreviewDegreeLabel, rotationInput);
  alignPreview.append(alignPreviewTitle, alignPreviewBox, alignPreviewDegree);
  alignPanel.appendChild(alignPreview);

  const textControl = document.createElement('div');
  textControl.className = 'fc-fmtdlg__text-control';
  const textControlTitle = document.createElement('div');
  textControlTitle.className = 'fc-fmtdlg__text-control-title';
  textControlTitle.textContent = t.textControl;
  const shrinkCk = makeCheckbox(t.shrinkToFit);
  shrinkCk.input.dataset.fcCheck = 'shrinkToFit';
  const mergeCk = makeCheckbox(t.mergeCells);
  mergeCk.input.dataset.fcCheck = 'mergeCells';
  textControl.append(textControlTitle, wrapCk.wrap, shrinkCk.wrap, mergeCk.wrap);
  alignPanel.appendChild(textControl);

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

  const normalFontCk = makeCheckbox(t.normalFont);
  normalFontCk.input.dataset.fcCheck = 'normalFont';
  fontPanel.appendChild(normalFontCk.wrap);

  // Font family
  const familyRow = document.createElement('label');
  familyRow.className = 'fc-fmtdlg__row';
  const familyLabel = document.createElement('span');
  familyLabel.textContent = t.fontFamily;
  const familyInput = document.createElement('input');
  familyInput.type = 'text';
  familyInput.setAttribute('aria-label', t.fontFamily);
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

  const familyList = document.createElement('div');
  familyList.className = 'fc-fmtdlg__font-list fc-fmtdlg__font-list--family';
  familyList.setAttribute('role', 'listbox');
  familyList.setAttribute('aria-label', t.fontFamily);
  for (const [index, family] of COMMON_FONTS.slice(0, 8).entries()) {
    const item = document.createElement('button');
    item.type = 'button';
    item.className = 'fc-fmtdlg__font-list-item';
    item.textContent = family;
    item.setAttribute('role', 'option');
    item.setAttribute('aria-selected', index === 0 ? 'true' : 'false');
    item.addEventListener('click', () => {
      familyInput.value = family;
      familyInput.dispatchEvent(new Event('input', { bubbles: true }));
    });
    familyList.appendChild(item);
  }
  fontPanel.appendChild(familyList);

  const fontStyleList = document.createElement('div');
  fontStyleList.className = 'fc-fmtdlg__font-list fc-fmtdlg__font-list--style';
  fontStyleList.setAttribute('role', 'listbox');
  fontStyleList.setAttribute('aria-label', t.fontStyle);
  const fontStyleOptions = [
    { id: 'regular', label: t.fontRegular },
    { id: 'italic', label: t.fontItalic },
    { id: 'bold', label: t.fontBold },
    { id: 'boldItalic', label: `${t.fontBold} ${t.fontItalic}` },
  ] as const;
  for (const [index, option] of fontStyleOptions.entries()) {
    const item = document.createElement('button');
    item.type = 'button';
    item.className = 'fc-fmtdlg__font-list-item';
    item.textContent = option.label;
    item.dataset.fcFontStyle = option.id;
    item.setAttribute('role', 'option');
    item.setAttribute('aria-selected', index === 0 ? 'true' : 'false');
    fontStyleList.appendChild(item);
  }
  fontPanel.appendChild(fontStyleList);

  // Font size
  const sizeRow = document.createElement('label');
  sizeRow.className = 'fc-fmtdlg__row';
  const sizeLabel = document.createElement('span');
  sizeLabel.textContent = t.fontSize;
  const sizeInput = document.createElement('input');
  sizeInput.type = 'number';
  sizeInput.setAttribute('aria-label', t.fontSize);
  sizeInput.min = '8';
  sizeInput.max = '72';
  sizeInput.step = '1';
  sizeRow.append(sizeLabel, sizeInput);
  fontPanel.appendChild(sizeRow);

  const sizeList = document.createElement('div');
  sizeList.className = 'fc-fmtdlg__font-list fc-fmtdlg__font-list--size';
  sizeList.setAttribute('role', 'listbox');
  sizeList.setAttribute('aria-label', t.fontSize);
  for (const size of [8, 9, 10, 11, 12, 14, 16, 18]) {
    const item = document.createElement('button');
    item.type = 'button';
    item.className = 'fc-fmtdlg__font-list-item';
    item.textContent = String(size);
    item.setAttribute('role', 'option');
    item.setAttribute('aria-selected', size === 12 ? 'true' : 'false');
    item.addEventListener('click', () => {
      sizeInput.value = String(size);
      sizeInput.dispatchEvent(new Event('input', { bubbles: true }));
    });
    sizeList.appendChild(item);
  }
  fontPanel.appendChild(sizeList);

  // Font color
  const colorRow = document.createElement('div');
  colorRow.className = 'fc-fmtdlg__row';
  const colorLabel = document.createElement('span');
  colorLabel.textContent = t.color;
  const colorInput = document.createElement('input');
  colorInput.type = 'color';
  colorInput.setAttribute('aria-label', t.color);
  colorInput.dataset.fcColor = 'font';
  const colorReset = makeButton(t.resetToDefault);
  colorRow.append(colorLabel, colorInput, colorReset);
  fontPanel.appendChild(colorRow);
  const fontSwatches = makeSwatches('font', t.themeColors, t.standardColors);
  fontPanel.appendChild(fontSwatches.el);

  const fontPreview = document.createElement('div');
  fontPreview.className = 'fc-fmtdlg__font-preview';
  const fontPreviewLabel = document.createElement('div');
  fontPreviewLabel.className = 'fc-fmtdlg__font-preview-label';
  fontPreviewLabel.textContent = t.preview;
  const fontPreviewBox = document.createElement('div');
  fontPreviewBox.className = 'fc-fmtdlg__font-preview-box';
  fontPreviewBox.textContent = t.previewText;
  fontPreview.append(fontPreviewLabel, fontPreviewBox);
  fontPanel.appendChild(fontPreview);

  // ── Border tab ─────────────────────────────────────────────────────────
  const borderPanel = tabPanels.get('border') as HTMLDivElement;

  // Active style + color row
  const borderStyleRow = document.createElement('label');
  borderStyleRow.className = 'fc-fmtdlg__row';
  const borderStyleLabel = document.createElement('span');
  borderStyleLabel.textContent = t.borderStyle;
  const borderStyleSelect = document.createElement('select');
  borderStyleSelect.setAttribute('aria-label', t.borderStyle);
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
  borderColorInput.setAttribute('aria-label', t.borderColor);
  borderColorInput.dataset.fcColor = 'border';
  const borderColorReset = makeButton(t.resetToDefault);
  borderColorRow.append(borderColorLabel, borderColorInput, borderColorReset);
  borderPanel.appendChild(borderColorRow);
  const borderSwatches = makeSwatches('border', t.themeColors, t.standardColors);
  borderPanel.appendChild(borderSwatches.el);

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
  fillInput.setAttribute('aria-label', t.fill);
  fillInput.dataset.fcColor = 'fill';
  const fillReset = makeButton(t.fillNone);
  fillRow.append(fillLabel, fillInput, fillReset);
  fillSection.appendChild(fillRow);
  const fillSwatches = makeSwatches('fill', t.themeColors, t.standardColors);
  fillSection.appendChild(fillSwatches.el);
  const fillPatternRow = document.createElement('label');
  fillPatternRow.className = 'fc-fmtdlg__row';
  const fillPatternLabel = document.createElement('span');
  fillPatternLabel.textContent = t.fillPatternStyle;
  const fillPatternSelect = document.createElement('select');
  fillPatternSelect.setAttribute('aria-label', t.fillPatternStyle);
  fillPatternSelect.dataset.fcSelect = 'fillPattern';
  const fillPatternOptions: Array<{ value: '' | FillPattern; label: string }> = [
    { value: '', label: t.fillPatternSolid },
    { value: 'gray125', label: t.fillPatternGray125 },
    { value: 'gray25', label: t.fillPatternGray25 },
    { value: 'gray50', label: t.fillPatternGray50 },
    { value: 'horizontal', label: t.fillPatternHorizontal },
    { value: 'vertical', label: t.fillPatternVertical },
    { value: 'diagonalDown', label: t.fillPatternDiagonalDown },
    { value: 'diagonalUp', label: t.fillPatternDiagonalUp },
  ];
  for (const option of fillPatternOptions) {
    const opt = document.createElement('option');
    opt.value = option.value;
    opt.textContent = option.label;
    fillPatternSelect.appendChild(opt);
  }
  fillPatternRow.append(fillPatternLabel, fillPatternSelect);
  fillSection.appendChild(fillPatternRow);
  const fillPatternColorRow = document.createElement('div');
  fillPatternColorRow.className = 'fc-fmtdlg__row';
  const fillPatternColorLabel = document.createElement('span');
  fillPatternColorLabel.textContent = t.fillPatternColor;
  const fillPatternColorInput = document.createElement('input');
  fillPatternColorInput.type = 'color';
  fillPatternColorInput.setAttribute('aria-label', t.fillPatternColor);
  fillPatternColorInput.dataset.fcColor = 'fillPattern';
  fillPatternColorRow.append(fillPatternColorLabel, fillPatternColorInput);
  fillSection.appendChild(fillPatternColorRow);

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
  const hiddenFormulaCk = makeCheckbox(strings.protection.hiddenFormula);
  hiddenFormulaCk.input.dataset.fcCheck = 'formulaHidden';
  lockedRow.append(lockedCk.wrap, hiddenFormulaCk.wrap);
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
  hlInput.setAttribute('aria-label', t.hyperlink);
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
  commentArea.setAttribute('aria-label', t.comment);
  commentArea.placeholder = t.commentPlaceholder;
  const commentClear = makeButton(t.clearField);
  commentRow.append(commentArea, commentClear);

  // Kind selector — drives the visibility of the bound/formula/list rows.
  const validationKindRow = document.createElement('label');
  validationKindRow.className = 'fc-fmtdlg__row';
  const validationKindLabel = document.createElement('span');
  validationKindLabel.textContent = t.validationKind;
  const validationKindSelect = document.createElement('select');
  validationKindSelect.setAttribute('aria-label', t.validationKind);
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
  validationOpSelect.setAttribute('aria-label', t.validationOp);
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
  validationAInput.setAttribute('aria-label', t.validationValueA);
  validationAInput.step = 'any';
  validationARow.append(validationALabel, validationAInput);
  validationSection.appendChild(validationARow);

  const validationBRow = document.createElement('label');
  validationBRow.className = 'fc-fmtdlg__row';
  const validationBLabel = document.createElement('span');
  validationBLabel.textContent = t.validationValueB;
  const validationBInput = document.createElement('input');
  validationBInput.type = 'number';
  validationBInput.setAttribute('aria-label', t.validationValueB);
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
  validationFormulaInput.setAttribute('aria-label', t.validationFormula);
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
  validationArea.setAttribute('aria-label', t.validationListSourceLiteral);
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
  validationListRangeInput.setAttribute('aria-label', t.validationListSourceRange);
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
  validationErrorStyleSelect.setAttribute('aria-label', t.validationErrorStyle);
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

  const validationShowInputMessageRow = document.createElement('label');
  validationShowInputMessageRow.className = 'fc-fmtdlg__check';
  const validationShowInputMessageInput = document.createElement('input');
  validationShowInputMessageInput.type = 'checkbox';
  const validationShowInputMessageSpan = document.createElement('span');
  validationShowInputMessageSpan.textContent = t.validationShowInputMessage;
  validationShowInputMessageRow.append(
    validationShowInputMessageInput,
    validationShowInputMessageSpan,
  );
  validationSection.appendChild(validationShowInputMessageRow);

  const validationPromptTitleRow = document.createElement('label');
  validationPromptTitleRow.className = 'fc-fmtdlg__row';
  const validationPromptTitleLabel = document.createElement('span');
  validationPromptTitleLabel.textContent = t.validationPromptTitle;
  const validationPromptTitleInput = document.createElement('input');
  validationPromptTitleInput.type = 'text';
  validationPromptTitleInput.setAttribute('aria-label', t.validationPromptTitle);
  validationPromptTitleRow.append(validationPromptTitleLabel, validationPromptTitleInput);
  validationSection.appendChild(validationPromptTitleRow);

  const validationPromptMessageRow = document.createElement('div');
  validationPromptMessageRow.className = 'fc-fmtdlg__row fc-fmtdlg__row--block';
  const validationPromptMessageArea = document.createElement('textarea');
  validationPromptMessageArea.className = 'fc-fmtdlg__textarea';
  validationPromptMessageArea.rows = 3;
  validationPromptMessageArea.setAttribute('aria-label', t.validationPromptMessage);
  validationPromptMessageRow.appendChild(validationPromptMessageArea);
  validationSection.appendChild(validationPromptMessageRow);

  const validationShowErrorMessageRow = document.createElement('label');
  validationShowErrorMessageRow.className = 'fc-fmtdlg__check';
  const validationShowErrorMessageInput = document.createElement('input');
  validationShowErrorMessageInput.type = 'checkbox';
  const validationShowErrorMessageSpan = document.createElement('span');
  validationShowErrorMessageSpan.textContent = t.validationShowErrorMessage;
  validationShowErrorMessageRow.append(
    validationShowErrorMessageInput,
    validationShowErrorMessageSpan,
  );
  validationSection.appendChild(validationShowErrorMessageRow);

  const validationErrorTitleRow = document.createElement('label');
  validationErrorTitleRow.className = 'fc-fmtdlg__row';
  const validationErrorTitleLabel = document.createElement('span');
  validationErrorTitleLabel.textContent = t.validationErrorTitle;
  const validationErrorTitleInput = document.createElement('input');
  validationErrorTitleInput.type = 'text';
  validationErrorTitleInput.setAttribute('aria-label', t.validationErrorTitle);
  validationErrorTitleRow.append(validationErrorTitleLabel, validationErrorTitleInput);
  validationSection.appendChild(validationErrorTitleRow);

  const validationErrorMessageRow = document.createElement('div');
  validationErrorMessageRow.className = 'fc-fmtdlg__row fc-fmtdlg__row--block';
  const validationErrorMessageArea = document.createElement('textarea');
  validationErrorMessageArea.className = 'fc-fmtdlg__textarea';
  validationErrorMessageArea.rows = 3;
  validationErrorMessageArea.setAttribute('aria-label', t.validationErrorMessage);
  validationErrorMessageRow.appendChild(validationErrorMessageArea);
  validationSection.appendChild(validationErrorMessageRow);

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
    calendarSelect,
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
    lockedCk,
    hiddenFormulaCk,
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
  };
}
