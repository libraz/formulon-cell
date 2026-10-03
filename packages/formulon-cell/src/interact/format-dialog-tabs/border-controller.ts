// Border tab behaviour for the Format Cells dialog: line style and color,
// presets, per-side toggles and the clickable border stage. A preset click is
// remembered separately from the per-side draft so OK can apply it as a
// selection-shaped action (outline vs. every inner edge).

import type { BorderStyleKey, SideKey } from '../format-dialog-model.js';
import { isHexColor } from '../format-dialog-model.js';
import { createPaletteFlyout, type PaletteFlyout } from '../format-dialog-palette-flyout.js';
import { activeDraftSide, type FormatDialogField, setDraftSide } from '../format-dialog-state.js';
import type { BorderTabRefs } from './border.js';
import { type FormatTabContext, type FormatTabController, markMixed } from './controller.js';

export type BorderPreset = 'none' | 'outline' | 'all';

export interface BorderTabController extends FormatTabController {
  readonly palette: PaletteFlyout;
  /** The preset clicked since the last per-side edit, if any. */
  pendingPreset(): BorderPreset | null;
  clearPendingPreset(): void;
}

const SIDES = ['top', 'right', 'bottom', 'left', 'diagonalDown', 'diagonalUp'] as const;

export function attachBorderTab(ctx: FormatTabContext, refs: BorderTabRefs): BorderTabController {
  const { draft, on, touch, isMixed, syncControls, renderPreview } = ctx;
  const {
    borderStyleSelect,
    borderStyleButtons,
    borderStyleGallery,
    borderColorInput,
    borderColorReset,
    borderSwatches,
    borderSwatchesToggle,
    borderSwatchesFlyout,
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
  } = refs;
  const palette = createPaletteFlyout(borderSwatchesToggle, borderSwatchesFlyout, borderSwatches);
  let pendingPreset: BorderPreset | null = null;
  const sideChecks: Record<SideKey, HTMLInputElement> = {
    top: topCk.input,
    right: rightCk.input,
    bottom: bottomCk.input,
    left: leftCk.input,
    diagonalDown: diagDownCk.input,
    diagonalUp: diagUpCk.input,
  };

  const activeSide = () => activeDraftSide(draft);
  const setSide = (key: SideKey, enabled: boolean): void => {
    draft.borders = setDraftSide(draft, key, enabled);
  };

  const sync = (): void => {
    borderStyleSelect.value = draft.borderStyle;
    for (const [id, btn] of borderStyleButtons) {
      btn.setAttribute('aria-pressed', id === draft.borderStyle ? 'true' : 'false');
    }
    borderColorInput.value =
      draft.borderColor && isHexColor(draft.borderColor) ? draft.borderColor : '#000000';
    borderSwatches.setValue(
      draft.borderColor && isHexColor(draft.borderColor) ? draft.borderColor : null,
    );
    topCk.input.checked = !!draft.borders.top;
    bottomCk.input.checked = !!draft.borders.bottom;
    leftCk.input.checked = !!draft.borders.left;
    rightCk.input.checked = !!draft.borders.right;
    diagDownCk.input.checked = !!draft.borders.diagonalDown;
    diagUpCk.input.checked = !!draft.borders.diagonalUp;
  };

  const syncMixed = (): void => {
    for (const side of SIDES) {
      const field = `border.${side}` as FormatDialogField;
      if (isMixed(field)) {
        const control = sideChecks[side];
        control.indeterminate = true;
        markMixed(control, true, field);
      }
    }
  };

  // ── Events ─────────────────────────────────────────────────────────────
  const onBorderStyleChange = (): void => {
    draft.borderStyle = borderStyleSelect.value as BorderStyleKey;
    pendingPreset = null;
    syncControls();
    renderPreview();
  };
  const onBorderStyleGalleryClick = (e: Event): void => {
    const btn = (e.target as HTMLElement).closest<HTMLButtonElement>('[data-border-style]');
    const style = btn?.dataset.borderStyle as BorderStyleKey | undefined;
    if (!style) return;
    draft.borderStyle = style;
    borderStyleSelect.value = style;
    pendingPreset = null;
    syncControls();
    renderPreview();
  };
  const onBorderColorInput = (): void => {
    draft.borderColor = borderColorInput.value;
    renderPreview();
  };
  const onBorderColorReset = (): void => {
    draft.borderColor = undefined;
    borderSwatches.setValue(null);
    renderPreview();
  };
  const onBorderSwatchesToggle = (): void => palette.setOpen(!palette.isOpen());
  const onBorderSwatchClick = (e: Event): void => {
    const btn = (e.target as HTMLElement).closest<HTMLButtonElement>('[data-color]');
    const color = btn?.dataset.color;
    if (!color) return;
    draft.borderColor = color;
    borderColorInput.value = color;
    renderPreview();
    palette.setOpen(false);
    borderSwatchesToggle.focus();
  };
  const onPresetNone = (): void => {
    touch(
      'border.top',
      'border.right',
      'border.bottom',
      'border.left',
      'border.diagonalDown',
      'border.diagonalUp',
    );
    pendingPreset = 'none';
    draft.borders = {};
    syncControls();
    renderPreview();
  };
  const onPresetOutline = (): void => {
    touch('border.top', 'border.right', 'border.bottom', 'border.left');
    pendingPreset = 'outline';
    draft.borders = {
      top: activeSide(),
      right: activeSide(),
      bottom: activeSide(),
      left: activeSide(),
    };
    syncControls();
    renderPreview();
  };
  const onPresetAll = (): void => {
    touch('border.top', 'border.right', 'border.bottom', 'border.left');
    pendingPreset = 'all';
    draft.borders = {
      top: activeSide(),
      right: activeSide(),
      bottom: activeSide(),
      left: activeSide(),
    };
    syncControls();
    renderPreview();
  };
  /** Per-side checkbox change: the checkbox is the source of truth. */
  const sideChange = (key: SideKey) => (): void => {
    touch(`border.${key}` as FormatDialogField);
    pendingPreset = null;
    setSide(key, sideChecks[key].checked);
    renderPreview();
  };
  const onVisualSideClick = (e: Event): void => {
    const btn = (e.target as HTMLElement).closest<HTMLButtonElement>('[data-border-side]');
    if (!btn) return;
    const key = btn.dataset.borderSide as SideKey;
    touch(`border.${key}` as FormatDialogField);
    pendingPreset = null;
    setSide(key, !draft.borders[key]);
    syncControls();
    renderPreview();
  };

  on(borderStyleSelect, 'change', onBorderStyleChange);
  on(borderStyleGallery, 'click', onBorderStyleGalleryClick);
  on(borderColorInput, 'input', onBorderColorInput);
  on(borderColorReset, 'click', onBorderColorReset);
  on(borderSwatchesToggle, 'click', onBorderSwatchesToggle);
  on(borderSwatches.el, 'click', onBorderSwatchClick);
  on(presetNone, 'click', onPresetNone);
  on(presetOutline, 'click', onPresetOutline);
  on(presetAll, 'click', onPresetAll);
  on(topCk.input, 'change', sideChange('top'));
  on(bottomCk.input, 'change', sideChange('bottom'));
  on(leftCk.input, 'change', sideChange('left'));
  on(rightCk.input, 'change', sideChange('right'));
  on(diagDownCk.input, 'change', sideChange('diagonalDown'));
  on(diagUpCk.input, 'change', sideChange('diagonalUp'));
  on(borderVisualStage, 'click', onVisualSideClick);

  return {
    sync,
    syncMixed,
    palette,
    pendingPreset: () => pendingPreset,
    clearPendingPreset: () => {
      pendingPreset = null;
    },
  };
}
