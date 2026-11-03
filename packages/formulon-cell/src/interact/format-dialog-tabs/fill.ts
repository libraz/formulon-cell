// Fill tab DOM for the Format Cells dialog. Background color + pattern style
// pair, plus the swatch grid shared with the font/border tabs.

import type { Strings } from '../../i18n/strings.js';
import type { FillPattern } from '../../store/store.js';
import { createDialogSelect } from '../../toolbar/dialogs/form-controls.js';
import { createDialogToggleButton } from '../dialog-shell.js';
import { makeButton, makeSwatches } from '../format-dialog-dom.js';
import { GRADIENT_SHADING_STYLES, type GradientShadingStyle } from '../format-dialog-model.js';

/** A popup trigger with its current-value text and flyout. */
export interface FillPopup {
  button: HTMLButtonElement;
  value: HTMLElement;
  flyout: HTMLElement;
}

export interface FillPopups {
  fill: FillPopup;
  patternColor: FillPopup;
  pattern: FillPopup;
  /** Flyout entries: No colour, Automatic. */
  fillNone: HTMLButtonElement;
  patternAuto: HTMLButtonElement;
}

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
  /** Mac presentation: colour and pattern pickers shown as popups. */
  popups: FillPopups | null;
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

export function createFillTab(
  panel: HTMLDivElement,
  t: Strings['formatDialog'],
  mac = false,
): FillTabRefs {
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

  let popups: FillPopups | null = null;
  if (mac) {
    popups = macFillPopups(t, {
      fillSection,
      fillRow,
      fillInput,
      fillReset,
      fillSwatches,
      fillPatternColorRow,
      fillPatternColorInput,
      fillPatternRow,
      fillPatternSelect,
      fillPatternGallery,
      fillEffectsButton,
      effects: effects.fillEffects,
    });
  }

  const sample = document.createElement('div');
  sample.className = 'fc-fmtdlg__fill-sample';
  const sampleLabel = document.createElement('div');
  sampleLabel.className = 'fc-fmtdlg__fill-sample-label';
  sampleLabel.textContent = t.sample;
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
    popups,
    ...effects,
  };
}

interface MacFillParts {
  fillSection: HTMLElement;
  fillRow: HTMLElement;
  fillInput: HTMLInputElement;
  fillReset: HTMLButtonElement;
  fillSwatches: ReturnType<typeof makeSwatches>;
  fillPatternColorRow: HTMLElement;
  fillPatternColorInput: HTMLInputElement;
  fillPatternRow: HTMLElement;
  fillPatternSelect: HTMLSelectElement;
  fillPatternGallery: HTMLElement;
  fillEffectsButton: HTMLButtonElement;
  effects: HTMLElement;
}

const makePopup = (label: string, initial: string): FillPopup => {
  const button = createDialogToggleButton({
    label,
    baseClass: 'fc-fmtdlg__popup fc-fmtdlg__fill-popup',
    popup: true,
  });
  const value = document.createElement('span');
  value.className = 'fc-fmtdlg__fill-popup-label';
  value.textContent = initial;
  const chevron = document.createElement('span');
  chevron.className = 'fc-fmtdlg__color-toggle-chevron';
  chevron.setAttribute('aria-hidden', 'true');
  button.append(value, chevron);
  const flyout = document.createElement('div');
  flyout.className = 'fc-fmtdlg__color-flyout fc-fmtdlg__fill-popup-flyout';
  flyout.hidden = true;
  return { button, value, flyout };
};

/** Mac presentation of the fill tab: two boxes (background colour; pattern
 *  colour and style) whose pickers are popups starting on No colour and
 *  Automatic. The colour inputs stay in the DOM, hidden, as the editing model. */
function macFillPopups(t: Strings['formatDialog'], p: MacFillParts): FillPopups {
  const fill = makePopup(t.fill, t.fillNoColor);
  const patternColor = makePopup(t.fillPatternColor, t.automatic);
  const pattern = makePopup(t.fillPatternStyle, '');

  p.fillReset.textContent = t.fillNoColor;
  p.fillReset.dataset.fcFillNone = '';
  fill.flyout.append(p.fillReset, p.fillSwatches.el);
  p.fillInput.hidden = true;
  p.fillRow.replaceChildren(p.fillRow.firstElementChild as Node, fill.button, p.fillInput);
  p.fillRow.appendChild(fill.flyout);

  const patternAuto = makeButton(t.automatic);
  patternAuto.dataset.fcFillPatternAuto = '';
  patternColor.flyout.append(patternAuto, p.fillPatternColorInput);
  p.fillPatternColorRow.replaceChildren(
    p.fillPatternColorRow.firstElementChild as Node,
    patternColor.button,
    patternColor.flyout,
  );

  pattern.flyout.appendChild(p.fillPatternGallery);
  const styleLabel = p.fillPatternRow.firstElementChild as Node;
  p.fillPatternSelect.hidden = true;
  p.fillPatternRow.replaceChildren(styleLabel, pattern.button, p.fillPatternSelect, pattern.flyout);
  (p.fillPatternRow as HTMLElement).classList.add('fc-fmtdlg__fill-pattern-popup-row');

  p.fillEffectsButton.hidden = true;
  p.effects.hidden = true;
  p.fillSection.classList.add('fc-fmtdlg__fill-boxes');
  return { fill, patternColor, pattern, fillNone: p.fillReset, patternAuto };
}
