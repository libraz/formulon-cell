// Font tab DOM for the Format Cells dialog. Bold/italic/underline/strike
// checkboxes plus family/size/color pickers and swatch grids.

import type { Strings } from '../../i18n/strings.js';
import { makeButton, makeCheckbox, makeSwatches } from '../format-dialog-dom.js';
import { COMMON_FONTS } from '../format-dialog-model.js';

export interface FontTabRefs {
  boldCk: ReturnType<typeof makeCheckbox>;
  italicCk: ReturnType<typeof makeCheckbox>;
  underlineCk: ReturnType<typeof makeCheckbox>;
  strikeCk: ReturnType<typeof makeCheckbox>;
  normalFontCk: ReturnType<typeof makeCheckbox>;
  fontStyleList: HTMLDivElement;
  familyInput: HTMLInputElement;
  sizeInput: HTMLInputElement;
  colorInput: HTMLInputElement;
  colorReset: HTMLButtonElement;
  fontSwatches: ReturnType<typeof makeSwatches>;
}

export function createFontTab(panel: HTMLDivElement, t: Strings['formatDialog']): FontTabRefs {
  const styleRow = document.createElement('div');
  styleRow.className = 'fc-fmtdlg__choice-grid';
  panel.appendChild(styleRow);

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
  panel.appendChild(normalFontCk.wrap);

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
  panel.appendChild(familyRow);

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
  panel.appendChild(familyList);

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
  panel.appendChild(fontStyleList);

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
  panel.appendChild(sizeRow);

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
  panel.appendChild(sizeList);

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
  panel.appendChild(colorRow);
  const fontSwatches = makeSwatches('font', t.themeColors, t.standardColors);
  panel.appendChild(fontSwatches.el);

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
    underlineCk,
    strikeCk,
    normalFontCk,
    fontStyleList,
    familyInput,
    sizeInput,
    colorInput,
    colorReset,
    fontSwatches,
  };
}
