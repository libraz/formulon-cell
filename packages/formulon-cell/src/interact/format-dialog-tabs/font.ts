// Font tab DOM for the Format Cells dialog. Bold/italic/underline/strike
// checkboxes plus family/size/color pickers and swatch grids.

import type { Strings } from '../../i18n/strings.js';
import {
  appendDialogDatalistOptions,
  createDialogSelect,
} from '../../toolbar/dialogs/form-controls.js';
import { shouldShowFontOption } from '../../toolbar/ribbon/font-availability.js';
import { FONT_FAMILIES, FONT_SIZES } from '../../toolbar/ribbon-model.js';
import { appendDialogOptionButton, createDialogToggleButton } from '../dialog-shell.js';
import { makeButton, makeCheckbox, makeSwatches } from '../format-dialog-dom.js';

export interface FontTabRefs {
  boldCk: ReturnType<typeof makeCheckbox>;
  italicCk: ReturnType<typeof makeCheckbox>;
  underlineSelect: HTMLSelectElement;
  strikeCk: ReturnType<typeof makeCheckbox>;
  superscriptCk: ReturnType<typeof makeCheckbox>;
  subscriptCk: ReturnType<typeof makeCheckbox>;
  normalFontCk: ReturnType<typeof makeCheckbox>;
  fontStyleList: HTMLDivElement;
  familyInput: HTMLInputElement;
  sizeInput: HTMLInputElement;
  colorInput: HTMLInputElement;
  colorReset: HTMLButtonElement;
  fontSwatches: ReturnType<typeof makeSwatches>;
  fontSwatchesToggle: HTMLButtonElement;
  fontSwatchesFlyout: HTMLDivElement;
  fontPreviewBox: HTMLDivElement;
  syncFontFamilyOptions: (current: string, selectFirstWhenEmpty?: boolean) => void;
  syncFontSizeOptions: (current: number | undefined) => void;
}

export function createFontTab(
  panel: HTMLDivElement,
  t: Strings['formatDialog'],
  locale: 'ja' | 'en',
): FontTabRefs {
  const styleRow = document.createElement('div');
  styleRow.className = 'fc-fmtdlg__choice-grid fc-fmtdlg__font-effects';
  panel.appendChild(styleRow);

  const boldCk = makeCheckbox(t.fontBold);
  boldCk.input.dataset.fcCheck = 'bold';
  const italicCk = makeCheckbox(t.fontItalic);
  italicCk.input.dataset.fcCheck = 'italic';
  const underlineLabel = document.createElement('label');
  underlineLabel.className = 'fc-fmtdlg__font-underline';
  underlineLabel.textContent = t.fontUnderline;
  const underlineSelect = createDialogSelect(
    [
      { value: '', label: t.fontUnderlineNone },
      { value: 'single', label: t.fontUnderlineSingle },
      { value: 'double', label: t.fontUnderlineDouble },
      { value: 'singleAccounting', label: t.fontUnderlineSingleAccounting },
      { value: 'doubleAccounting', label: t.fontUnderlineDoubleAccounting },
    ],
    '',
    { ariaLabel: t.fontUnderline },
  );
  underlineSelect.dataset.fcInput = 'underline';
  underlineLabel.appendChild(underlineSelect);
  const strikeCk = makeCheckbox(t.fontStrike);
  strikeCk.input.dataset.fcCheck = 'strike';
  const superscriptCk = makeCheckbox(t.fontSuperscript);
  superscriptCk.input.dataset.fcCheck = 'superscript';
  const subscriptCk = makeCheckbox(t.fontSubscript);
  subscriptCk.input.dataset.fcCheck = 'subscript';
  styleRow.append(
    boldCk.wrap,
    italicCk.wrap,
    underlineLabel,
    strikeCk.wrap,
    superscriptCk.wrap,
    subscriptCk.wrap,
  );

  const normalFontCk = makeCheckbox(t.normalFont);
  normalFontCk.input.dataset.fcCheck = 'normalFont';
  normalFontCk.wrap.classList.add('fc-fmtdlg__normal-font');
  panel.appendChild(normalFontCk.wrap);

  // Font family
  const familyRow = document.createElement('label');
  familyRow.className = 'fc-fmtdlg__row fc-fmtdlg__font-family-row';
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
  familyRow.append(familyLabel, familyInput, familyDatalist);
  panel.appendChild(familyRow);

  const familyList = document.createElement('div');
  familyList.className = 'fc-fmtdlg__font-list fc-fmtdlg__font-list--family';
  familyList.setAttribute('role', 'listbox');
  familyList.setAttribute('aria-label', t.fontFamily);
  panel.appendChild(familyList);

  const syncFontFamilyOptions = (current: string, selectFirstWhenEmpty = true): void => {
    const families = FONT_FAMILIES.filter((family) =>
      shouldShowFontOption(family, current, locale),
    );
    appendDialogDatalistOptions(familyDatalist, families);
    familyList.replaceChildren();
    const visibleFamilies = families.slice(0, 8);
    const knownCurrent = current as (typeof FONT_FAMILIES)[number];
    if (
      current &&
      families.some((family) => family === current) &&
      !visibleFamilies.some((family) => family === current)
    ) {
      visibleFamilies.splice(-1, 1, knownCurrent);
    }
    for (const [index, family] of visibleFamilies.entries()) {
      const item = appendDialogOptionButton(familyList, {
        label: family,
        baseClass: 'fc-fmtdlg__font-list-item',
        datasetKey: 'fcFontFamily',
        value: family,
        selected: family === current || (selectFirstWhenEmpty && !current && index === 0),
      });
      item.addEventListener('click', () => {
        familyInput.value = family;
        familyInput.dispatchEvent(new Event('input', { bubbles: true }));
      });
    }
  };
  syncFontFamilyOptions('');

  // The style column has no free-text input of its own, so it carries a bare
  // caption to keep the three column headings on one line.
  const fontStyleLabel = document.createElement('div');
  fontStyleLabel.className = 'fc-fmtdlg__font-style-label';
  fontStyleLabel.textContent = t.fontStyle;
  panel.appendChild(fontStyleLabel);

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
    appendDialogOptionButton(fontStyleList, {
      label: option.label,
      baseClass: 'fc-fmtdlg__font-list-item',
      datasetKey: 'fcFontStyle',
      value: option.id,
      selected: index === 0,
    });
  }
  panel.appendChild(fontStyleList);

  // Font size
  const sizeRow = document.createElement('label');
  sizeRow.className = 'fc-fmtdlg__row fc-fmtdlg__font-size-row';
  const sizeLabel = document.createElement('span');
  sizeLabel.textContent = t.fontSize;
  const sizeInput = document.createElement('input');
  sizeInput.type = 'number';
  sizeInput.setAttribute('aria-label', t.fontSize);
  sizeInput.min = '1';
  sizeInput.max = '409';
  sizeInput.step = '1';
  sizeRow.append(sizeLabel, sizeInput);
  panel.appendChild(sizeRow);

  const sizeList = document.createElement('div');
  sizeList.className = 'fc-fmtdlg__font-list fc-fmtdlg__font-list--size';
  sizeList.setAttribute('role', 'listbox');
  sizeList.setAttribute('aria-label', t.fontSize);
  for (const size of FONT_SIZES) {
    const item = appendDialogOptionButton(sizeList, {
      label: String(size),
      baseClass: 'fc-fmtdlg__font-list-item',
      datasetKey: 'fcFontSize',
      value: String(size),
      selected: false,
    });
    item.addEventListener('click', () => {
      sizeInput.value = String(size);
      sizeInput.dispatchEvent(new Event('input', { bubbles: true }));
    });
  }
  panel.appendChild(sizeList);
  const syncFontSizeOptions = (current: number | undefined): void => {
    for (const item of sizeList.querySelectorAll<HTMLButtonElement>('[data-fc-font-size]')) {
      item.setAttribute(
        'aria-selected',
        current !== undefined && Number(item.dataset.fcFontSize) === current ? 'true' : 'false',
      );
    }
  };

  // Font color. The swatch palette is taller than the space this tab has left,
  // so it hangs off the color control as a flyout instead of sitting inline the
  // way the fill tab's palette does.
  const colorRow = document.createElement('div');
  colorRow.className = 'fc-fmtdlg__row fc-fmtdlg__font-color-row';
  const colorLabel = document.createElement('span');
  colorLabel.textContent = t.color;
  const colorInput = document.createElement('input');
  colorInput.type = 'color';
  colorInput.setAttribute('aria-label', t.color);
  colorInput.dataset.fcColor = 'font';
  const fontSwatchesToggle = createDialogToggleButton({
    label: t.themeColors,
    baseClass: 'fc-fmtdlg__color-toggle',
    title: t.themeColors,
    popup: true,
  });
  const fontSwatchesChevron = document.createElement('span');
  fontSwatchesChevron.className = 'fc-fmtdlg__color-toggle-chevron';
  fontSwatchesChevron.setAttribute('aria-hidden', 'true');
  fontSwatchesToggle.appendChild(fontSwatchesChevron);
  const colorControl = document.createElement('span');
  colorControl.className = 'fc-fmtdlg__color-control';
  colorControl.append(colorInput, fontSwatchesToggle);
  const colorReset = makeButton(t.resetToDefault);
  colorRow.append(colorLabel, colorControl, colorReset);
  panel.appendChild(colorRow);
  const fontSwatches = makeSwatches('font', t.themeColors, t.standardColors);
  const fontSwatchesFlyout = document.createElement('div');
  fontSwatchesFlyout.className = 'fc-fmtdlg__color-flyout';
  fontSwatchesFlyout.hidden = true;
  fontSwatchesFlyout.appendChild(fontSwatches.el);
  panel.appendChild(fontSwatchesFlyout);

  const fontPreview = document.createElement('div');
  fontPreview.className = 'fc-fmtdlg__font-preview';
  const fontPreviewLabel = document.createElement('div');
  fontPreviewLabel.className = 'fc-fmtdlg__font-preview-label';
  fontPreviewLabel.textContent = t.preview;
  const fontPreviewBox = document.createElement('div');
  fontPreviewBox.className = 'fc-fmtdlg__font-preview-box';
  fontPreviewBox.textContent = t.previewText;
  fontPreview.append(fontPreviewLabel, fontPreviewBox);
  panel.appendChild(fontPreview);

  return {
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
    syncFontSizeOptions,
  };
}
