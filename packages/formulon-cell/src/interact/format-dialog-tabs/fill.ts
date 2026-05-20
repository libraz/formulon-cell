// Fill tab DOM for the Format Cells dialog. Background color + pattern style
// pair, plus the swatch grid shared with the font/border tabs.

import type { Strings } from '../../i18n/strings.js';
import type { FillPattern } from '../../store/store.js';
import { createDialogSelect } from '../../toolbar/dialogs/form-controls.js';
import { makeButton, makeSwatches } from '../format-dialog-dom.js';

export interface FillTabRefs {
  fillInput: HTMLInputElement;
  fillReset: HTMLButtonElement;
  fillSwatches: ReturnType<typeof makeSwatches>;
  fillPatternSelect: HTMLSelectElement;
  fillPatternColorInput: HTMLInputElement;
}

export function createFillTab(panel: HTMLDivElement, t: Strings['formatDialog']): FillTabRefs {
  const fillSection = document.createElement('div');
  fillSection.className = 'fc-fmtdlg__section';
  const fillSectionTitle = document.createElement('div');
  fillSectionTitle.className = 'fc-fmtdlg__section-title';
  fillSectionTitle.textContent = t.fill;
  fillSection.appendChild(fillSectionTitle);
  panel.appendChild(fillSection);
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
  const fillPatternSelect = createDialogSelect(fillPatternOptions, '', {
    ariaLabel: t.fillPatternStyle,
    className: '',
  });
  fillPatternSelect.dataset.fcSelect = 'fillPattern';
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

  return {
    fillInput,
    fillReset,
    fillSwatches,
    fillPatternSelect,
    fillPatternColorInput,
  };
}
