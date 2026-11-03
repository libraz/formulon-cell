// Border tab DOM for the Format Cells dialog. Style+color pickers, preset
// shortcuts, per-side checkboxes, and the visual border-stage previewer.

import type { Strings } from '../../i18n/strings.js';
import { createDialogSelect } from '../../toolbar/dialogs/form-controls.js';
import { projectDisabledState } from '../../toolbar/menu-a11y.js';
import { createLineSamplePreview } from '../../toolbar/ribbon/border-icons.js';
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
  /** Mac presentation: the No line cell, which makes edge clicks clear. */
  borderNoneButton: HTMLButtonElement | null;
  /** Mac presentation: the colour popup's current-value text. */
  borderColorValue: HTMLElement | null;
  borderColorInput: HTMLInputElement;
  borderColorReset: HTMLButtonElement;
  borderSwatches: ReturnType<typeof makeSwatches>;
  borderSwatchesToggle: HTMLButtonElement;
  borderSwatchesFlyout: HTMLDivElement;
  presetNone: HTMLButtonElement;
  presetOutline: HTMLButtonElement;
  presetInside: HTMLButtonElement;
  topCk: ReturnType<typeof makeCheckbox>;
  bottomCk: ReturnType<typeof makeCheckbox>;
  leftCk: ReturnType<typeof makeCheckbox>;
  rightCk: ReturnType<typeof makeCheckbox>;
  diagDownCk: ReturnType<typeof makeCheckbox>;
  diagUpCk: ReturnType<typeof makeCheckbox>;
  borderVisualStage: HTMLDivElement;
  borderVisualPreview: HTMLDivElement;
  visualSideButtons: Map<SideKey, HTMLButtonElement[]>;
  /** Mac presentation: the inside-horizontal and inside-vertical buttons. */
  insideButtons: HTMLButtonElement[];
}

/** The desktop dialog's line styles, down the left column then the right. */
const DIALOG_LINE_STYLES: readonly BorderStyleKey[] = [
  'hair',
  'dotted',
  'dashDotDot',
  'dashDot',
  'dashed',
  'thin',
  'mediumDashDotDot',
  'slantDashDot',
  'mediumDashDot',
  'mediumDashed',
  'medium',
  'thick',
  'double',
];

export function createBorderTab(
  panel: HTMLDivElement,
  t: Strings['formatDialog'],
  mac = false,
): BorderTabRefs {
  // Active style + color row
  const borderStyleRow = document.createElement('label');
  borderStyleRow.className = 'fc-fmtdlg__row fc-fmtdlg__border-style-row';
  const borderStyleLabel = document.createElement('span');
  borderStyleLabel.textContent = t.borderStyle;
  const styleLabels: Record<BorderStyleKey, string> = {
    thin: t.borderStyleThin,
    medium: t.borderStyleMedium,
    thick: t.borderStyleThick,
    dashed: t.borderStyleDashed,
    dotted: t.borderStyleDotted,
    double: t.borderStyleDouble,
    hair: t.borderStyleHair,
    mediumDashed: t.borderStyleMediumDashed,
    dashDot: t.borderStyleDashDot,
    mediumDashDot: t.borderStyleMediumDashDot,
    dashDotDot: t.borderStyleDashDotDot,
    mediumDashDotDot: t.borderStyleMediumDashDotDot,
    slantDashDot: t.borderStyleSlantDashDot,
  };
  const styleOptions = DIALOG_LINE_STYLES.map((id) => ({ id, label: styleLabels[id] }));
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
  let borderNoneButton: HTMLButtonElement | null = null;
  if (mac) {
    borderNoneButton = createDialogToggleButton({
      label: t.borderStyleNone,
      baseClass: 'fc-fmtdlg__line-style',
      extraClass: 'fc-fmtdlg__line-style--none',
      datasetKey: 'borderStyle',
      value: 'none',
    });
    const label = document.createElement('span');
    label.textContent = t.borderStyleNone;
    borderNoneButton.appendChild(label);
    borderStyleGallery.appendChild(borderNoneButton);
  }
  for (const s of styleOptions) {
    const btn = createDialogToggleButton({
      label: s.label,
      baseClass: 'fc-fmtdlg__line-style',
      extraClass: `fc-fmtdlg__line-style--${s.id}`,
      datasetKey: 'borderStyle',
      value: s.id,
    });
    const sample = createLineSamplePreview(s.id);
    // The Mac app draws the hairline as a fine dotted rule.
    if (mac && s.id === 'hair') {
      const rule = sample.querySelector('line');
      rule?.setAttribute('stroke-dasharray', '0.7 0.9');
      rule?.setAttribute('stroke-width', '0.8');
    }
    sample.classList.add('fc-fmtdlg__line-sample-svg');
    const label = document.createElement('span');
    label.textContent = s.label;
    btn.append(sample, label);
    borderStyleButtons.set(s.id, btn);
    borderStyleGallery.appendChild(btn);
  }
  panel.appendChild(borderStyleGallery);

  // Line color. Like the font tab, the swatch palette hangs off the color
  // control as a flyout — the two-column border layout has no room for it
  // inline, and the desktop dialog shows a color dropdown here too.
  const borderColorRow = document.createElement('div');
  borderColorRow.className = 'fc-fmtdlg__row fc-fmtdlg__border-color-row';
  const borderColorLabel = document.createElement('span');
  borderColorLabel.textContent = t.borderColor;
  const borderColorInput = document.createElement('input');
  borderColorInput.type = 'color';
  borderColorInput.setAttribute('aria-label', t.borderColor);
  borderColorInput.dataset.fcColor = 'border';
  const borderSwatchesToggle = createDialogToggleButton({
    label: t.themeColors,
    baseClass: 'fc-fmtdlg__color-toggle',
    title: t.themeColors,
    popup: true,
  });
  const borderSwatchesChevron = document.createElement('span');
  borderSwatchesChevron.className = 'fc-fmtdlg__color-toggle-chevron';
  borderSwatchesChevron.setAttribute('aria-hidden', 'true');
  borderSwatchesToggle.appendChild(borderSwatchesChevron);
  const borderColorControl = document.createElement('span');
  borderColorControl.className = 'fc-fmtdlg__color-control';
  borderColorControl.append(borderColorInput, borderSwatchesToggle);
  const borderColorReset = makeButton(mac ? t.automatic : t.resetToDefault);
  let borderColorValue: HTMLElement | null = null;
  if (mac) {
    // The colour reads as a popup: its text is the current colour, Automatic
    // first, and the reset button becomes the flyout's Automatic entry.
    borderColorValue = document.createElement('span');
    borderColorValue.className = 'fc-fmtdlg__color-toggle-label';
    borderColorValue.textContent = t.automatic;
    borderSwatchesToggle.prepend(borderColorValue);
    borderSwatchesToggle.classList.add('fc-fmtdlg__popup');
    borderColorInput.hidden = true;
    borderColorRow.append(borderColorLabel, borderColorControl);
  } else {
    borderColorRow.append(borderColorLabel, borderColorControl, borderColorReset);
  }
  panel.appendChild(borderColorRow);
  const borderSwatches = makeSwatches('border', t.themeColors, t.standardColors);
  const borderSwatchesFlyout = document.createElement('div');
  borderSwatchesFlyout.className = 'fc-fmtdlg__color-flyout';
  borderSwatchesFlyout.hidden = true;
  if (mac) borderSwatchesFlyout.appendChild(borderColorReset);
  borderSwatchesFlyout.appendChild(borderSwatches.el);
  panel.appendChild(borderSwatchesFlyout);

  // Presets
  if (mac) {
    const presetsHeading = document.createElement('div');
    presetsHeading.className = 'fc-fmtdlg__border-heading';
    presetsHeading.textContent = t.borderPresets;
    panel.appendChild(presetsHeading);
  }
  const presetRow = document.createElement('div');
  presetRow.className = 'fc-fmtdlg__row fc-fmtdlg__border-presets';
  panel.appendChild(presetRow);
  const presetNone = makeButton(t.borderPresetNone);
  const presetOutline = makeButton(t.borderPresetOutline);
  const presetInside = makeButton(t.borderPresetInside);
  presetNone.classList.add('fc-fmtdlg__border-preset', 'fc-fmtdlg__border-preset--none');
  presetOutline.classList.add('fc-fmtdlg__border-preset', 'fc-fmtdlg__border-preset--outline');
  presetInside.classList.add('fc-fmtdlg__border-preset', 'fc-fmtdlg__border-preset--inside');
  presetRow.append(presetNone, presetOutline, presetInside);

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
  borderVisualTitle.textContent = mac ? t.borderOutline : t.preview;
  if (mac) borderVisualTitle.classList.add('fc-fmtdlg__border-heading');
  const borderVisualStage = document.createElement('div');
  borderVisualStage.className = 'fc-fmtdlg__border-stage';
  const borderVisualPreview = document.createElement('div');
  borderVisualPreview.className = 'fc-fmtdlg__border-preview';
  borderVisualPreview.textContent = mac ? t.borderPreviewText : t.previewText;
  if (mac) {
    for (const corner of ['tl', 'tr', 'bl', 'br']) {
      const tick = document.createElement('span');
      tick.className = `fc-fmtdlg__border-tick fc-fmtdlg__border-tick--${corner}`;
      borderVisualPreview.appendChild(tick);
    }
  }
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
  const insideButtons: HTMLButtonElement[] = [];
  if (mac) {
    for (const [axis, label] of [
      ['h', t.borderInsideHorizontal],
      ['v', t.borderInsideVertical],
    ] as const) {
      const btn = createDialogToggleButton({
        label,
        baseClass: `fc-fmtdlg__border-hit fc-fmtdlg__border-hit--inside-${axis}`,
        datasetKey: 'borderInside',
        value: axis,
      });
      projectDisabledState(btn, true, null);
      insideButtons.push(btn);
      borderVisualStage.appendChild(btn);
    }
  }
  borderVisual.append(borderVisualTitle, borderVisualStage);
  panel.appendChild(borderVisual);
  if (mac) {
    const hint = document.createElement('div');
    hint.className = 'fc-fmtdlg__border-hint';
    hint.textContent = t.borderHint;
    panel.appendChild(hint);
  }

  return {
    borderStyleSelect,
    borderStyleButtons,
    borderStyleGallery,
    borderNoneButton,
    borderColorValue,
    borderColorInput,
    borderColorReset,
    borderSwatches,
    borderSwatchesToggle,
    borderSwatchesFlyout,
    presetNone,
    presetOutline,
    presetInside,
    topCk,
    bottomCk,
    leftCk,
    rightCk,
    diagDownCk,
    diagUpCk,
    borderVisualStage,
    borderVisualPreview,
    visualSideButtons,
    insideButtons,
  };
}
