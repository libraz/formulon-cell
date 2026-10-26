// Fill tab behaviour for the Format Cells dialog: background color, swatch
// grid, pattern style gallery and pattern color.

import { gradientFillCss } from '../../render/gradient-fill-css.js';
import type { FillPattern } from '../../store/store.js';
import {
  type GradientShadingStyle,
  gradientForShading,
  gradientVariantCount,
  isHexColor,
  shadingForGradient,
} from '../format-dialog-model.js';
import { type FormatTabContext, type FormatTabController, markMixed } from './controller.js';
import type { FillTabRefs } from './fill.js';

export function attachFillTab(ctx: FormatTabContext, refs: FillTabRefs): FormatTabController {
  const { draft, on, touch, isMixed, syncControls, renderPreview, isDxfMode } = ctx;
  const {
    fillInput,
    fillReset,
    fillSwatches,
    fillPatternSelect,
    fillPatternGallery,
    fillPatternColorInput,
    fillEffectsButton,
    fillEffects,
    gradientColor1,
    gradientColor2,
    gradientStyleInputs,
    gradientVariants,
  } = refs;

  // Fill Effects selection; a loaded gradient seeds it on sync.
  const shading: { style: GradientShadingStyle; variant: number } = {
    style: 'horizontal',
    variant: 0,
  };
  let effectsOpen = false;

  const syncFillEffects = (): void => {
    const loaded = draft.fillGradient ? shadingForGradient(draft.fillGradient) : null;
    if (loaded) {
      shading.style = loaded.style;
      shading.variant = loaded.variant;
      gradientColor1.value = loaded.color1;
      gradientColor2.value = loaded.color2;
    } else if (draft.fillGradient) {
      const stops = draft.fillGradient.stops;
      gradientColor1.value = stops[0]?.color ?? '#ffffff';
      gradientColor2.value = stops[stops.length - 1]?.color ?? '#4472c4';
    }
    const offered = !isDxfMode();
    fillEffectsButton.hidden = !offered;
    fillEffects.hidden = !offered || (!effectsOpen && !draft.fillGradient);
    fillEffectsButton.setAttribute('aria-expanded', fillEffects.hidden ? 'false' : 'true');
    for (const [style, input] of gradientStyleInputs) {
      input.checked = (loaded !== null || !draft.fillGradient) && style === shading.style;
    }
    const count = gradientVariantCount(shading.style);
    gradientVariants.forEach((button, index) => {
      button.hidden = index >= count;
      button.style.backgroundImage = gradientFillCss(
        gradientForShading(shading.style, index, gradientColor1.value, gradientColor2.value),
      );
      button.setAttribute(
        'aria-pressed',
        draft.fillGradient && index === shading.variant ? 'true' : 'false',
      );
    });
  };

  const sync = (): void => {
    fillInput.value = draft.fill && isHexColor(draft.fill) ? draft.fill : '#ffffff';
    fillSwatches.setValue(draft.fill && isHexColor(draft.fill) ? draft.fill : null);
    fillPatternSelect.value = draft.fillPattern ?? '';
    for (const button of fillPatternGallery.querySelectorAll<HTMLButtonElement>(
      '[data-fc-fill-pattern]',
    )) {
      button.setAttribute(
        'aria-pressed',
        button.dataset.fcFillPattern === (draft.fillPattern ?? '') ? 'true' : 'false',
      );
    }
    fillPatternColorInput.value =
      draft.fillPatternColor && isHexColor(draft.fillPatternColor)
        ? draft.fillPatternColor
        : '#000000';
    syncFillEffects();
  };

  const syncMixed = (): void => {
    if (isMixed('fill')) {
      fillSwatches.setValue(null);
      markMixed(fillInput, true, 'fill');
    }
    if (isMixed('fillPattern')) {
      fillPatternSelect.value = '';
      for (const button of fillPatternGallery.querySelectorAll<HTMLButtonElement>(
        '[data-fc-fill-pattern]',
      )) {
        button.setAttribute('aria-pressed', 'false');
      }
      markMixed(fillPatternGallery, true, 'fillPattern');
      markMixed(fillPatternSelect, true, 'fillPattern');
    }
    if (isMixed('fillPatternColor')) {
      markMixed(fillPatternColorInput, true, 'fillPatternColor');
    }
  };

  // ── Events ─────────────────────────────────────────────────────────────
  const applyShading = (): void => {
    touch('fillGradient');
    draft.fillGradient = gradientForShading(
      shading.style,
      Math.min(shading.variant, gradientVariantCount(shading.style) - 1),
      gradientColor1.value,
      gradientColor2.value,
    );
    draft.fill = undefined;
    draft.fillPattern = undefined;
    draft.fillPatternColor = undefined;
    syncControls();
    renderPreview();
  };
  const onFillEffectsToggle = (): void => {
    effectsOpen = fillEffects.hidden === true;
    if (gradientColor1.value === '#000000' && !draft.fillGradient) {
      gradientColor1.value = draft.fill && isHexColor(draft.fill) ? draft.fill : '#ffffff';
      gradientColor2.value = '#4472c4';
    }
    syncFillEffects();
  };
  const onGradientStyleChange = (e: Event): void => {
    const style = (e.target as HTMLInputElement).dataset.fcGradientStyle as
      | GradientShadingStyle
      | undefined;
    if (!style) return;
    shading.style = style;
    applyShading();
  };
  const onGradientVariantClick = (e: Event): void => {
    const button = (e.target as HTMLElement).closest<HTMLButtonElement>(
      '[data-fc-gradient-variant]',
    );
    if (!button) return;
    shading.variant = Number(button.dataset.fcGradientVariant ?? 0);
    applyShading();
  };
  // A solid or pattern choice replaces the gradient.
  const dropGradient = (): void => {
    if (!draft.fillGradient) return;
    touch('fillGradient');
    draft.fillGradient = undefined;
  };
  const onFillInput = (): void => {
    dropGradient();
    touch('fill');
    draft.fill = fillInput.value;
    syncControls();
    renderPreview();
  };
  const onFillReset = (): void => {
    dropGradient();
    touch('fill');
    draft.fill = undefined;
    fillSwatches.setValue(null);
    syncControls();
    renderPreview();
  };
  const onFillPatternChange = (): void => {
    dropGradient();
    touch('fillPattern');
    draft.fillPattern = (fillPatternSelect.value || undefined) as FillPattern | undefined;
    syncControls();
    renderPreview();
  };
  const onFillPatternGalleryClick = (e: Event): void => {
    const button = (e.target as HTMLElement).closest<HTMLButtonElement>('[data-fc-fill-pattern]');
    if (!button) return;
    fillPatternSelect.value = button.dataset.fcFillPattern ?? '';
    onFillPatternChange();
  };
  const onFillPatternColorInput = (): void => {
    dropGradient();
    touch('fillPatternColor');
    draft.fillPatternColor = fillPatternColorInput.value;
    syncControls();
    renderPreview();
  };
  const onFillSwatchClick = (e: Event): void => {
    const btn = (e.target as HTMLElement).closest<HTMLButtonElement>('[data-color]');
    const color = btn?.dataset.color;
    if (!color) return;
    dropGradient();
    touch('fill');
    draft.fill = color;
    fillInput.value = color;
    syncControls();
    renderPreview();
  };

  on(fillInput, 'input', onFillInput);
  on(fillReset, 'click', onFillReset);
  on(fillPatternSelect, 'change', onFillPatternChange);
  on(fillPatternGallery, 'click', onFillPatternGalleryClick);
  on(fillPatternColorInput, 'input', onFillPatternColorInput);
  on(fillSwatches.el, 'click', onFillSwatchClick);
  on(fillEffectsButton, 'click', onFillEffectsToggle);
  on(gradientColor1, 'input', applyShading);
  on(gradientColor2, 'input', applyShading);
  for (const input of gradientStyleInputs.values()) on(input, 'change', onGradientStyleChange);
  for (const button of gradientVariants) on(button, 'click', onGradientVariantClick);

  return { sync, syncMixed };
}
