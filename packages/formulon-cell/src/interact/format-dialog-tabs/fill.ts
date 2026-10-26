// Fill tab DOM for the Format Cells dialog. Background color + pattern style
// pair, plus the swatch grid shared with the font/border tabs.

import type { Strings } from '../../i18n/strings.js';
import type { FillPattern } from '../../store/store.js';
import { createDialogSelect } from '../../toolbar/dialogs/form-controls.js';
import { makeButton, makeSwatches } from '../format-dialog-dom.js';
import { GRADIENT_SHADING_STYLES, type GradientShadingStyle } from '../format-dialog-model.js';

export interface FillTabRefs {
  fillInput: HTMLInputElement;
  fillReset: HTMLButtonElement;
  fillSwatches: ReturnType<typeof makeSwatches>;
  fillPatternSelect: HTMLSelectElement;
  fillPatternGallery: HTMLDivElement;
  fillPatternColorInput: HTMLInputElement;
  fillSample: HTMLDivElement;
  fillEffectsButton: HTMLButtonElement;
  fillEffects: HTMLDivElement;
  gradientColor1: HTMLInputElement;
  gradientColor2: HTMLInputElement;
  gradientStyleInputs: ReadonlyMap<GradientShadingStyle, HTMLInputElement>;
  gradientVariants: readonly HTMLButtonElement[];
}

const shadingLabel = (t: Strings['formatDialog'], style: GradientShadingStyle): string => {
  switch (style) {
    case 'horizontal':
      return t.gradientHorizontal;
    case 'vertical':
      return t.gradientVertical;
    case 'diagonalUp':
      return t.gradientDiagonalUp;
    case 'diagonalDown':
      return t.gradientDiagonalDown;
    case 'fromCorner':
      return t.gradientFromCorner;
    case 'fromCenter':
      return t.gradientFromCenter;
  }
};

/** Fill Effects: two colours, a shading style and its variants. */
const createFillEffects = (
  t: Strings['formatDialog'],
): Pick<
  FillTabRefs,
  'fillEffects' | 'gradientColor1' | 'gradientColor2' | 'gradientStyleInputs' | 'gradientVariants'
> => {
  const fillEffects = document.createElement('div');
  fillEffects.className = 'fc-fmtdlg__fill-effects';
  fillEffects.hidden = true;
  const colorInput = (index: 1 | 2, label: string): HTMLInputElement => {
    const row = document.createElement('label');
    row.className = 'fc-fmtdlg__row';
    const text = document.createElement('span');
    text.textContent = label;
    const input = document.createElement('input');
    input.type = 'color';
    input.setAttribute('aria-label', label);
    input.dataset.fcGradientColor = String(index);
    row.append(text, input);
    fillEffects.appendChild(row);
    return input;
  };
  const gradientColor1 = colorInput(1, t.gradientColor1);
  const gradientColor2 = colorInput(2, t.gradientColor2);
  const styles = document.createElement('fieldset');
  styles.className = 'fc-fmtdlg__gradient-styles';
  const legend = document.createElement('legend');
  legend.textContent = t.gradientShadingStyle;
  styles.appendChild(legend);
  const gradientStyleInputs = new Map<GradientShadingStyle, HTMLInputElement>();
  for (const style of GRADIENT_SHADING_STYLES) {
    const label = document.createElement('label');
    label.className = 'fc-fmtdlg__check';
    const input = document.createElement('input');
    input.type = 'radio';
    input.name = 'fc-gradient-style';
    input.dataset.fcGradientStyle = style;
    const text = document.createElement('span');
    text.textContent = shadingLabel(t, style);
    label.append(input, text);
    styles.appendChild(label);
    gradientStyleInputs.set(style, input);
  }
  fillEffects.appendChild(styles);
  const variants = document.createElement('div');
  variants.className = 'fc-fmtdlg__gradient-variants';
  variants.setAttribute('role', 'group');
  variants.setAttribute('aria-label', t.gradientVariants);
  const gradientVariants = [0, 1, 2, 3].map((index) => {
    const button = makeButton('');
    button.className = 'fc-fmtdlg__gradient-variant';
    button.dataset.fcGradientVariant = String(index);
    button.setAttribute('aria-label', `${t.gradientVariants} ${index + 1}`);
    button.setAttribute('aria-pressed', 'false');
    variants.appendChild(button);
    return button;
  });
  fillEffects.appendChild(variants);
  return { fillEffects, gradientColor1, gradientColor2, gradientStyleInputs, gradientVariants };
};

export function createFillTab(panel: HTMLDivElement, t: Strings['formatDialog']): FillTabRefs {
  const fillSection = document.createElement('div');
  fillSection.className = 'fc-fmtdlg__section';
  const fillSectionTitle = document.createElement('div');
  fillSectionTitle.className = 'fc-fmtdlg__section-title';
  fillSectionTitle.textContent = t.fill;
  fillSection.appendChild(fillSectionTitle);
  panel.appendChild(fillSection);
  const fillRow = document.createElement('div');
  fillRow.className = 'fc-fmtdlg__row fc-fmtdlg__fill-bg-row';
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
  const fillPatternColorRow = document.createElement('div');
  fillPatternColorRow.className = 'fc-fmtdlg__row fc-fmtdlg__fill-pattern-color-row';
  const fillPatternColorLabel = document.createElement('span');
  fillPatternColorLabel.textContent = t.fillPatternColor;
  const fillPatternColorInput = document.createElement('input');
  fillPatternColorInput.type = 'color';
  fillPatternColorInput.setAttribute('aria-label', t.fillPatternColor);
  fillPatternColorInput.dataset.fcColor = 'fillPattern';
  fillPatternColorRow.append(fillPatternColorLabel, fillPatternColorInput);
  fillSection.appendChild(fillPatternColorRow);
  const fillPatternRow = document.createElement('label');
  fillPatternRow.className = 'fc-fmtdlg__row fc-fmtdlg__fill-pattern-row';
  const fillPatternLabel = document.createElement('span');
  fillPatternLabel.textContent = t.fillPatternStyle;
  const fillPatternOptions: Array<{ value: '' | FillPattern; label: string }> = [
    { value: '', label: t.fillPatternSolid },
    { value: 'gray125', label: t.fillPatternGray125 },
    { value: 'gray0625', label: t.fillPatternGray0625 },
    { value: 'gray25', label: t.fillPatternGray25 },
    { value: 'gray50', label: t.fillPatternGray50 },
    { value: 'gray75', label: t.fillPatternGray75 },
    { value: 'darkHorizontal', label: t.fillPatternDarkHorizontal },
    { value: 'darkVertical', label: t.fillPatternDarkVertical },
    { value: 'darkDown', label: t.fillPatternDarkDown },
    { value: 'darkUp', label: t.fillPatternDarkUp },
    { value: 'darkGrid', label: t.fillPatternDarkGrid },
    { value: 'darkTrellis', label: t.fillPatternDarkTrellis },
    { value: 'lightHorizontal', label: t.fillPatternLightHorizontal },
    { value: 'lightVertical', label: t.fillPatternLightVertical },
    { value: 'lightDown', label: t.fillPatternLightDown },
    { value: 'lightUp', label: t.fillPatternLightUp },
    { value: 'lightGrid', label: t.fillPatternLightGrid },
    { value: 'lightTrellis', label: t.fillPatternLightTrellis },
  ];
  const fillPatternSelect = createDialogSelect(fillPatternOptions, '', {
    ariaLabel: t.fillPatternStyle,
    className: '',
  });
  fillPatternSelect.dataset.fcSelect = 'fillPattern';
  fillPatternRow.append(fillPatternLabel, fillPatternSelect);
  fillSection.appendChild(fillPatternRow);
  const fillPatternGallery = document.createElement('div');
  fillPatternGallery.className = 'fc-fmtdlg__fill-pattern-gallery';
  fillPatternGallery.setAttribute('role', 'group');
  fillPatternGallery.setAttribute('aria-label', t.fillPatternStyle);
  for (const option of fillPatternOptions) {
    const button = makeButton(option.label);
    button.className = 'fc-fmtdlg__fill-pattern-option';
    button.textContent = '';
    button.dataset.fcFillPattern = option.value;
    button.setAttribute('aria-label', option.label);
    button.setAttribute('aria-pressed', option.value === '' ? 'true' : 'false');
    const sample = document.createElement('span');
    sample.className = 'fc-fmtdlg__fill-pattern-option-sample';
    sample.dataset.fcFillPatternSample = option.value;
    button.appendChild(sample);
    fillPatternGallery.appendChild(button);
  }
  fillSection.appendChild(fillPatternGallery);
  const fillEffectsButton = makeButton(t.fillEffects);
  fillEffectsButton.dataset.fcFillEffects = '';
  fillSection.appendChild(fillEffectsButton);
  const effects = createFillEffects(t);
  fillSection.appendChild(effects.fillEffects);

  const sample = document.createElement('div');
  sample.className = 'fc-fmtdlg__fill-sample';
  const sampleLabel = document.createElement('div');
  sampleLabel.className = 'fc-fmtdlg__fill-sample-label';
  sampleLabel.textContent = t.preview;
  const fillSample = document.createElement('div');
  fillSample.className = 'fc-fmtdlg__fill-sample-box';
  sample.append(sampleLabel, fillSample);
  panel.appendChild(sample);

  return {
    fillInput,
    fillReset,
    fillSwatches,
    fillPatternSelect,
    fillPatternGallery,
    fillPatternColorInput,
    fillSample,
    fillEffectsButton,
    ...effects,
  };
}
