// Font tab behaviour for the Format Cells dialog: style checkboxes and list,
// underline, family/size pickers and the font color with its palette flyout.

import { appendDialogSelectOptions } from '../../toolbar/dialogs/form-controls.js';
import { projectDisabledState } from '../../toolbar/menu-a11y.js';
import { isHexColor } from '../format-dialog-model.js';
import { createPaletteFlyout, type PaletteFlyout } from '../format-dialog-palette-flyout.js';
import type { FormatDialogField } from '../format-dialog-state.js';
import {
  type FormatTabContext,
  type FormatTabController,
  markMixed,
  markMixedCheck,
} from './controller.js';
import type { FontTabRefs } from './font.js';

export interface FontTabController extends FormatTabController {
  readonly palette: PaletteFlyout;
}

const NORMAL_FONT_FIELDS: readonly FormatDialogField[] = [
  'bold',
  'italic',
  'underline',
  'strike',
  'fontVertAlign',
  'fontFamily',
  'fontSize',
  'color',
];

export function attachFontTab(ctx: FormatTabContext, refs: FontTabRefs): FontTabController {
  const { draft, on, touch, isMixed, syncControls, renderPreview } = ctx;
  const {
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
    syncFontFamilyOptions,
    syncFontSizeOptions,
  } = refs;
  const palette = createPaletteFlyout(fontSwatchesToggle, fontSwatchesFlyout, fontSwatches);

  const underlineMixedOption = ((): HTMLOptionElement => {
    const scratch = document.createElement('select');
    appendDialogSelectOptions(scratch, [{ value: 'mixed', label: '' }]);
    const option = scratch.options[0] as HTMLOptionElement;
    projectDisabledState(option, true, null);
    return option;
  })();

  const currentFontStyleId = (): 'regular' | 'italic' | 'bold' | 'boldItalic' => {
    if (draft.bold && draft.italic) return 'boldItalic';
    if (draft.bold) return 'bold';
    if (draft.italic) return 'italic';
    return 'regular';
  };

  const syncFontStyleList = (): void => {
    const id = isMixed('bold') || isMixed('italic') ? null : currentFontStyleId();
    for (const item of fontStyleList.querySelectorAll<HTMLButtonElement>('[data-fc-font-style]')) {
      item.setAttribute(
        'aria-selected',
        id !== null && item.dataset.fcFontStyle === id ? 'true' : 'false',
      );
    }
  };

  const sync = (): void => {
    boldCk.input.checked = draft.bold;
    italicCk.input.checked = draft.italic;
    underlineMixedOption.remove();
    underlineSelect.value = draft.underline === true ? 'single' : draft.underline || '';
    strikeCk.input.checked = draft.strike;
    superscriptCk.input.checked = draft.fontVertAlign === 'superscript';
    subscriptCk.input.checked = draft.fontVertAlign === 'subscript';
    normalFontCk.input.checked =
      !draft.bold &&
      !draft.italic &&
      !draft.underline &&
      !draft.strike &&
      !draft.fontVertAlign &&
      !draft.fontFamily &&
      draft.fontSize === undefined &&
      draft.color === undefined;
    syncFontStyleList();
    familyInput.value = draft.fontFamily;
    syncFontFamilyOptions(draft.fontFamily);
    sizeInput.value = draft.fontSize !== undefined ? String(draft.fontSize) : '';
    syncFontSizeOptions(draft.fontSize);
    colorInput.value = draft.color && isHexColor(draft.color) ? draft.color : '#000000';
    fontSwatches.setValue(draft.color && isHexColor(draft.color) ? draft.color : null);
  };

  const syncMixed = (): void => {
    markMixedCheck(boldCk.input, isMixed('bold'), 'bold');
    markMixedCheck(italicCk.input, isMixed('italic'), 'italic');
    markMixedCheck(strikeCk.input, isMixed('strike'), 'strike');
    markMixedCheck(superscriptCk.input, isMixed('fontVertAlign'), 'fontVertAlign');
    markMixedCheck(subscriptCk.input, isMixed('fontVertAlign'), 'fontVertAlign');
    const normalFontMixed = NORMAL_FONT_FIELDS.some((field) => isMixed(field));
    normalFontCk.input.indeterminate = normalFontMixed;
    markMixed(normalFontCk.input, normalFontMixed);
    if (isMixed('underline')) {
      // A placeholder keeps the real "None" option distinct from the mixed state,
      // so choosing None still reports a change.
      underlineSelect.append(underlineMixedOption);
      underlineSelect.value = underlineMixedOption.value;
      markMixed(underlineSelect, true, 'underline');
    }
    if (isMixed('fontFamily')) {
      familyInput.value = '';
      markMixed(familyInput, true, 'fontFamily');
    }
    if (isMixed('fontSize')) {
      sizeInput.value = '';
      markMixed(sizeInput, true, 'fontSize');
    }
    syncFontStyleList();
    syncFontFamilyOptions(isMixed('fontFamily') ? '' : draft.fontFamily, !isMixed('fontFamily'));
    syncFontSizeOptions(isMixed('fontSize') ? undefined : draft.fontSize);
    if (isMixed('color')) {
      fontSwatches.setValue(null);
      markMixed(colorInput, true, 'color');
    }
  };

  // ── Events ─────────────────────────────────────────────────────────────
  const onBoldChange = (): void => {
    touch('bold');
    draft.bold = boldCk.input.checked;
    normalFontCk.input.checked = false;
    syncControls();
    renderPreview();
  };
  const onItalicChange = (): void => {
    touch('italic');
    draft.italic = italicCk.input.checked;
    normalFontCk.input.checked = false;
    syncControls();
    renderPreview();
  };
  const onUnderlineChange = (): void => {
    touch('underline');
    switch (underlineSelect.value) {
      case 'single':
      case 'double':
      case 'singleAccounting':
      case 'doubleAccounting':
        draft.underline = underlineSelect.value;
        break;
      default:
        draft.underline = false;
    }
    normalFontCk.input.checked = false;
    syncControls();
    renderPreview();
  };
  const onStrikeChange = (): void => {
    touch('strike');
    draft.strike = strikeCk.input.checked;
    normalFontCk.input.checked = false;
    syncControls();
    renderPreview();
  };
  const onSuperscriptChange = (): void => {
    touch('fontVertAlign');
    if (superscriptCk.input.checked) {
      draft.fontVertAlign = 'superscript';
      subscriptCk.input.checked = false;
    } else if (draft.fontVertAlign === 'superscript') {
      draft.fontVertAlign = undefined;
    }
    normalFontCk.input.checked = false;
    syncControls();
    renderPreview();
  };
  const onSubscriptChange = (): void => {
    touch('fontVertAlign');
    if (subscriptCk.input.checked) {
      draft.fontVertAlign = 'subscript';
      superscriptCk.input.checked = false;
    } else if (draft.fontVertAlign === 'subscript') {
      draft.fontVertAlign = undefined;
    }
    normalFontCk.input.checked = false;
    syncControls();
    renderPreview();
  };
  const onNormalFontChange = (): void => {
    if (!normalFontCk.input.checked) return;
    touch(...NORMAL_FONT_FIELDS);
    draft.bold = false;
    draft.italic = false;
    draft.underline = false;
    draft.strike = false;
    draft.fontVertAlign = undefined;
    draft.fontFamily = '';
    draft.fontSize = undefined;
    draft.color = undefined;
    syncControls();
    renderPreview();
  };
  const onFontStyleListClick = (e: Event): void => {
    const item = (e.target as HTMLElement).closest<HTMLButtonElement>('[data-fc-font-style]');
    const style = item?.dataset.fcFontStyle;
    if (!style) return;
    touch('bold', 'italic');
    draft.bold = style === 'bold' || style === 'boldItalic';
    draft.italic = style === 'italic' || style === 'boldItalic';
    boldCk.input.checked = draft.bold;
    italicCk.input.checked = draft.italic;
    normalFontCk.input.checked = false;
    syncControls();
    renderPreview();
  };
  const onFamilyInput = (): void => {
    touch('fontFamily');
    draft.fontFamily = familyInput.value;
    normalFontCk.input.checked = false;
    syncControls();
    renderPreview();
  };
  const onSizeInput = (): void => {
    touch('fontSize');
    if (sizeInput.value === '') {
      draft.fontSize = undefined;
    } else {
      const n = Number.parseInt(sizeInput.value, 10);
      if (Number.isFinite(n)) draft.fontSize = Math.max(1, Math.min(409, n));
    }
    normalFontCk.input.checked = false;
    syncControls();
    renderPreview();
  };
  const onColorInput = (): void => {
    touch('color');
    draft.color = colorInput.value;
    normalFontCk.input.checked = false;
    syncControls();
    renderPreview();
  };
  const onColorReset = (): void => {
    touch('color');
    draft.color = undefined;
    fontSwatches.setValue(null);
    syncControls();
    renderPreview();
  };
  const onFontSwatchesToggle = (): void => palette.setOpen(!palette.isOpen());
  const onFontSwatchClick = (e: Event): void => {
    const btn = (e.target as HTMLElement).closest<HTMLButtonElement>('[data-color]');
    const color = btn?.dataset.color;
    if (!color) return;
    touch('color');
    draft.color = color;
    colorInput.value = color;
    normalFontCk.input.checked = false;
    syncControls();
    renderPreview();
    palette.setOpen(false);
    fontSwatchesToggle.focus();
  };

  on(boldCk.input, 'change', onBoldChange);
  on(italicCk.input, 'change', onItalicChange);
  on(underlineSelect, 'change', onUnderlineChange);
  on(strikeCk.input, 'change', onStrikeChange);
  on(superscriptCk.input, 'change', onSuperscriptChange);
  on(subscriptCk.input, 'change', onSubscriptChange);
  on(normalFontCk.input, 'change', onNormalFontChange);
  on(fontStyleList, 'click', onFontStyleListClick as EventListener);
  on(familyInput, 'input', onFamilyInput);
  on(sizeInput, 'input', onSizeInput);
  on(colorInput, 'input', onColorInput);
  on(colorReset, 'click', onColorReset);
  on(fontSwatchesToggle, 'click', onFontSwatchesToggle);
  on(fontSwatches.el, 'click', onFontSwatchClick);

  return { sync, syncMixed, palette };
}
