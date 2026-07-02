import type { Strings } from '../i18n/strings.js';
import type { NegativeStyle } from '../store/store.js';
import { createDialogSelect } from '../toolbar/dialogs/form-controls.js';
import {
  appendDialogIconButton,
  appendDialogOptionButton,
  appendDialogTabPair,
  createDialogShell,
} from './dialog-shell.js';
import { makeButton, makeCheckbox } from './format-dialog-dom.js';
import { CURRENCY_SYMBOLS, type NumberCategory, type TabId } from './format-dialog-model.js';
import { createAlignTab } from './format-dialog-tabs/align.js';
import { createBorderTab } from './format-dialog-tabs/border.js';
import { createFillTab } from './format-dialog-tabs/fill.js';
import { createFontTab } from './format-dialog-tabs/font.js';
import { createMoreTab } from './format-dialog-tabs/more.js';

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
  headerTitle.className = 'fc-fmtdlg__title';
  headerTitle.textContent = t.title;
  header.appendChild(headerTitle);
  const closeBtn = appendDialogIconButton(header, {
    label: '×',
    ariaLabel: t.cancel,
    baseClass: 'fc-fmtdlg__close',
  });
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
    const { button, panel: panelEl } = appendDialogTabPair(tabsStrip, body, {
      id: def.id,
      label: def.label,
      tabId: `fc-fmtdlg-tab-${def.id}`,
      panelId: `fc-fmtdlg-panel-${def.id}`,
      tabDatasetKey: 'fcTab',
      panelDatasetKey: 'fcTab',
    });
    tabButtons.set(def.id, button);
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
    { id: 'special', label: t.catOther },
    { id: 'custom', label: t.catCustom },
  ];
  const catButtons = new Map<NumberCategory, HTMLButtonElement>();
  for (const c of catDefs) {
    const b = appendDialogOptionButton(catList, {
      label: c.label,
      baseClass: 'fc-fmtdlg__cat-item',
      datasetKey: 'fcCat',
      value: c.id,
    });
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
  const symbolSelect = createDialogSelect(
    CURRENCY_SYMBOLS.map((symbol) => ({ value: symbol, label: symbol })),
    CURRENCY_SYMBOLS[0] ?? '',
    { className: '', ariaLabel: t.symbol },
  );
  symbolRow.append(symbolLabel, symbolSelect);
  numberControls.appendChild(symbolRow);

  const patternPresetRow = document.createElement('label');
  patternPresetRow.className = 'fc-fmtdlg__row';
  const patternPresetLabel = document.createElement('span');
  patternPresetLabel.textContent = t.patternType;
  const patternPresetSelect = createDialogSelect([], '', {
    className: '',
    ariaLabel: t.patternType,
  });
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
  const localeSelect = createDialogSelect(
    [
      { value: 'ja', label: '日本語' },
      { value: 'en', label: 'English' },
    ],
    'ja',
    { className: '', ariaLabel: t.languageLocation },
  );
  localeRow.append(localeLabel, localeSelect);
  numberControls.appendChild(localeRow);

  // Calendar type — visible only for the date category. Choices mirror
  // Office365's Gregorian/Japanese options; pattern output is unchanged
  // since downstream rendering does not yet implement era formatting.
  const calendarRow = document.createElement('label');
  calendarRow.className = 'fc-fmtdlg__row';
  const calendarLabel = document.createElement('span');
  calendarLabel.textContent = t.calendarType;
  const calendarSelect = createDialogSelect(
    [
      { value: 'gregorian', label: t.calendarTypeGregorian },
      { value: 'japanese', label: t.calendarTypeJapanese },
    ],
    'gregorian',
    { className: '', ariaLabel: t.calendarType },
  );
  calendarSelect.dataset.fcSelect = 'calendarType';
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
    appendDialogOptionButton(negativeOptions, {
      label: sample.text,
      baseClass: 'fc-fmtdlg__negative-item',
      datasetKey: 'fcNegativeStyle',
      value: sample.value,
      selected: index === 3,
      extraClass: sample.red ? 'fc-fmtdlg__negative-item--red' : undefined,
    });
  }
  negativeList.append(negativeLabel, negativeOptions);
  numberControls.appendChild(negativeList);

  // ── Alignment tab ──────────────────────────────────────────────────────
  const alignTab = createAlignTab(tabPanels.get('align') as HTMLDivElement, t);
  const {
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
  } = alignTab;

  // ── Font tab ───────────────────────────────────────────────────────────
  const fontTab = createFontTab(tabPanels.get('font') as HTMLDivElement, t);
  const {
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
  } = fontTab;

  // ── Border tab ─────────────────────────────────────────────────────────
  const borderTab = createBorderTab(tabPanels.get('border') as HTMLDivElement, t);
  const {
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
  } = borderTab;

  // ── Fill tab ───────────────────────────────────────────────────────────
  const fillTab = createFillTab(tabPanels.get('fill') as HTMLDivElement, t);
  const { fillInput, fillReset, fillSwatches, fillPatternSelect, fillPatternColorInput } = fillTab;

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
  const moreTab = createMoreTab(tabPanels.get('more') as HTMLDivElement, t);
  const {
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
  } = moreTab;

  // ── Footer buttons ─────────────────────────────────────────────────────
  const cancelBtn = makeButton(t.cancel);
  const okBtn = makeButton(t.ok, true);
  footer.append(cancelBtn, okBtn);

  return {
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
  };
}
