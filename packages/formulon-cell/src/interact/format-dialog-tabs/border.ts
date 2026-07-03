// Border tab DOM for the Format Cells dialog. Style+color pickers, preset
// shortcuts, per-side checkboxes, and the visual border-stage previewer.

import type { Strings } from '../../i18n/strings.js';
import { createDialogSelect } from '../../toolbar/dialogs/form-controls.js';
import { createDialogToggleButton } from '../dialog-shell.js';
import {
  makeButton,
  makeCheckbox,
  makeSwatches,
  makeVisualSideButton,
} from '../format-dialog-dom.js';
import type { BorderStyleKey, SideKey } from '../format-dialog-model.js';

export interface BorderTabRefs {
  borderStyleSelect: HTMLSelectElement;
  borderStyleButtons: Map<BorderStyleKey, HTMLButtonElement>;
  borderStyleGallery: HTMLDivElement;
  borderColorInput: HTMLInputElement;
  borderColorReset: HTMLButtonElement;
  borderSwatches: ReturnType<typeof makeSwatches>;
  presetNone: HTMLButtonElement;
  presetOutline: HTMLButtonElement;
  presetAll: HTMLButtonElement;
  topCk: ReturnType<typeof makeCheckbox>;
  bottomCk: ReturnType<typeof makeCheckbox>;
  leftCk: ReturnType<typeof makeCheckbox>;
  rightCk: ReturnType<typeof makeCheckbox>;
  diagDownCk: ReturnType<typeof makeCheckbox>;
  diagUpCk: ReturnType<typeof makeCheckbox>;
  borderVisualStage: HTMLDivElement;
  borderVisualPreview: HTMLDivElement;
  visualSideButtons: Map<SideKey, HTMLButtonElement[]>;
}

export function createBorderTab(panel: HTMLDivElement, t: Strings['formatDialog']): BorderTabRefs {
  // Active style + color row
  const borderStyleRow = document.createElement('label');
  borderStyleRow.className = 'fc-fmtdlg__row';
  const borderStyleLabel = document.createElement('span');
  borderStyleLabel.textContent = t.borderStyle;
  const styleOptions: { id: BorderStyleKey; label: string }[] = [
    { id: 'thin', label: t.borderStyleThin },
    { id: 'medium', label: t.borderStyleMedium },
    { id: 'thick', label: t.borderStyleThick },
    { id: 'dashed', label: t.borderStyleDashed },
    { id: 'dotted', label: t.borderStyleDotted },
    { id: 'double', label: t.borderStyleDouble },
  ];
  const borderStyleSelect = createDialogSelect(
    styleOptions.map((s) => ({ value: s.id, label: s.label })),
    'thin',
    { ariaLabel: t.borderStyle, className: '' },
  );
  borderStyleRow.append(borderStyleLabel, borderStyleSelect);
  panel.appendChild(borderStyleRow);

  const borderStyleGallery = document.createElement('div');
  borderStyleGallery.className = 'fc-fmtdlg__line-gallery';
  const borderStyleButtons = new Map<BorderStyleKey, HTMLButtonElement>();
  for (const s of styleOptions) {
    const btn = createDialogToggleButton({
      label: s.label,
      baseClass: 'fc-fmtdlg__line-style',
      extraClass: `fc-fmtdlg__line-style--${s.id}`,
      datasetKey: 'borderStyle',
      value: s.id,
    });
    const sample = document.createElement('span');
    sample.className = 'fc-fmtdlg__line-sample';
    const label = document.createElement('span');
    label.textContent = s.label;
    btn.append(sample, label);
    borderStyleButtons.set(s.id, btn);
    borderStyleGallery.appendChild(btn);
  }
  panel.appendChild(borderStyleGallery);

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
  panel.appendChild(borderColorRow);
  const borderSwatches = makeSwatches('border', t.themeColors, t.standardColors);
  panel.appendChild(borderSwatches.el);

  // Presets
  const presetRow = document.createElement('div');
  presetRow.className = 'fc-fmtdlg__row fc-fmtdlg__border-presets';
  panel.appendChild(presetRow);
  const presetNone = makeButton(t.borderPresetNone);
  const presetOutline = makeButton(t.borderPresetOutline);
  const presetAll = makeButton(t.borderPresetAll);
  presetNone.classList.add('fc-fmtdlg__border-preset', 'fc-fmtdlg__border-preset--none');
  presetOutline.classList.add('fc-fmtdlg__border-preset', 'fc-fmtdlg__border-preset--outline');
  presetAll.classList.add('fc-fmtdlg__border-preset', 'fc-fmtdlg__border-preset--inside');
  presetRow.append(presetNone, presetOutline, presetAll);

  // Per-side checkboxes
  const sideRow = document.createElement('div');
  sideRow.className = 'fc-fmtdlg__row fc-fmtdlg__legacy-border-controls';
  panel.appendChild(sideRow);
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
  panel.appendChild(diagonalRow);
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
  panel.appendChild(borderVisual);

  return {
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
  };
}
