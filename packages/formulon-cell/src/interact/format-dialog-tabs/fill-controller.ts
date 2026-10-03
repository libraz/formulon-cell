// Fill tab behaviour for the Format Cells dialog: background color, swatch
// grid, pattern style gallery and pattern color.

import type { FillPattern } from '../../store/store.js';
import { isHexColor } from '../format-dialog-model.js';
import { type FormatTabContext, type FormatTabController, markMixed } from './controller.js';
import type { FillTabRefs } from './fill.js';

export function attachFillTab(ctx: FormatTabContext, refs: FillTabRefs): FormatTabController {
  const { draft, on, touch, isMixed, syncControls, renderPreview } = ctx;
  const {
    fillInput,
    fillReset,
    fillSwatches,
    fillPatternSelect,
    fillPatternGallery,
    fillPatternColorInput,
  } = refs;

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
  const onFillInput = (): void => {
    touch('fill');
    draft.fill = fillInput.value;
    syncControls();
    renderPreview();
  };
  const onFillReset = (): void => {
    touch('fill');
    draft.fill = undefined;
    fillSwatches.setValue(null);
    syncControls();
    renderPreview();
  };
  const onFillPatternChange = (): void => {
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
    touch('fillPatternColor');
    draft.fillPatternColor = fillPatternColorInput.value;
    syncControls();
    renderPreview();
  };
  const onFillSwatchClick = (e: Event): void => {
    const btn = (e.target as HTMLElement).closest<HTMLButtonElement>('[data-color]');
    const color = btn?.dataset.color;
    if (!color) return;
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

  return { sync, syncMixed };
}
