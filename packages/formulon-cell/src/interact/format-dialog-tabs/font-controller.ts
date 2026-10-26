// Font tab behaviour for the Format Cells dialog: style checkboxes and list,
// underline, family/size pickers and the font color with its palette flyout.

import { parseLocalizedNumericInput } from '../../commands/coerce-input.js';
import { fontCssPxToDisplayPoints, fontPointsToCssPx } from '../../engine/font-units.js';
import { dictionaryLocaleFor } from '../../i18n/strings.js';
import { appendDialogSelectOptions } from '../../toolbar/dialogs/form-controls.js';
import { projectDisabledState } from '../../toolbar/menu-a11y.js';
import {
  fontFamilyLabel,
  fontFamilyValueFromLabel,
} from '../../toolbar/ribbon/font-family-options.js';
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
  'fontOutline',
  'fontShadow',
  'fontFamily',
  'fontSize',
  'color',
];

export function attachFontTab(ctx: FormatTabContext, refs: FontTabRefs): FontTabController {
  const {
    draft,
    getLocale,
    on,
    touch,
    isMixed,
    syncControls,
    renderPreview,
    isMacPlatform,
    isDxfMode,
  } = ctx;
  const {
    boldCk,
    italicCk,
    underlineSelect,
    strikeCk,
    superscriptCk,
    subscriptCk,
    outlineCk,
    shadowCk,
    normalFontCk,
    fontStyleList,
    familyInput,
    sizeInput,
    colorInput,
    colorReset,
    fontSwatches,
    fontSwatchesToggle,
    fontSwatchesFlyout,
    fontLocale,
    syncFontFamilyOptions,
    syncFontSizeOptions,
  } = refs;
  const palette = createPaletteFlyout(fontSwatchesToggle, fontSwatchesFlyout, fontSwatches);
  const currentFontLocale = (): import('../../i18n/strings.js').Locale =>
    fontLocale ?? dictionaryLocaleFor(getLocale());

  // A number input is always parsed with the browser's locale-independent
  // grammar. Excel's font-size box follows the active workbook/UI locale. Keep
  // the existing number control for dot-decimal locales, and use a text
  // decimal control where a comma is part of the locale grammar.
  const language = getLocale().replace('_', '-').split('-', 1)[0]?.toLowerCase();
  if (language === 'de' || language === 'fr') {
    sizeInput.type = 'text';
    sizeInput.inputMode = 'decimal';
  }

  const displayPointSize = (px: number | undefined): string => {
    if (px === undefined) return '';
    const points = fontCssPxToDisplayPoints(px);
    try {
      return new Intl.NumberFormat(getLocale(), {
        useGrouping: false,
        maximumFractionDigits: 6,
      }).format(points);
    } catch {
      // Host integrations may expose a BCP-47-like profile identifier that
      // is not accepted by Intl. Keep the dialog usable with the base locale.
      return new Intl.NumberFormat('en-US', {
        useGrouping: false,
        maximumFractionDigits: 6,
      }).format(points);
    }
  };

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
    const effectsOffered = isMacPlatform() && !isDxfMode();
    outlineCk.wrap.hidden = !effectsOffered;
    shadowCk.wrap.hidden = !effectsOffered;
    outlineCk.input.checked = draft.fontOutline;
    shadowCk.input.checked = draft.fontShadow;
    normalFontCk.input.checked =
      !draft.bold &&
      !draft.italic &&
      !draft.underline &&
      !draft.strike &&
      !draft.fontVertAlign &&
      !draft.fontOutline &&
      !draft.fontShadow &&
      !draft.fontFamily &&
      draft.fontSize === undefined &&
      draft.color === undefined;
    syncFontStyleList();
    familyInput.value = fontFamilyLabel(currentFontLocale(), draft.fontFamily);
    syncFontFamilyOptions(draft.fontFamily);
    sizeInput.value = displayPointSize(draft.fontSize);
    syncFontSizeOptions(
      draft.fontSize !== undefined ? fontCssPxToDisplayPoints(draft.fontSize) : undefined,
    );
    colorInput.value = draft.color && isHexColor(draft.color) ? draft.color : '#000000';
    fontSwatches.setValue(draft.color && isHexColor(draft.color) ? draft.color : null);
  };

  const syncMixed = (): void => {
    markMixedCheck(boldCk.input, isMixed('bold'), 'bold');
    markMixedCheck(italicCk.input, isMixed('italic'), 'italic');
    markMixedCheck(strikeCk.input, isMixed('strike'), 'strike');
    markMixedCheck(superscriptCk.input, isMixed('fontVertAlign'), 'fontVertAlign');
    markMixedCheck(subscriptCk.input, isMixed('fontVertAlign'), 'fontVertAlign');
    markMixedCheck(outlineCk.input, isMixed('fontOutline'), 'fontOutline');
    markMixedCheck(shadowCk.input, isMixed('fontShadow'), 'fontShadow');
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
    syncFontSizeOptions(
      isMixed('fontSize') || draft.fontSize === undefined
        ? undefined
        : fontCssPxToDisplayPoints(draft.fontSize),
    );
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
  const onEffectChange = (field: 'fontOutline' | 'fontShadow', input: HTMLInputElement): void => {
    touch(field);
    draft[field] = input.checked;
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
    draft.fontOutline = false;
    draft.fontShadow = false;
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
  const applyFamilyValue = (value: string): void => {
    touch('fontFamily');
    draft.fontFamily = value;
    normalFontCk.input.checked = false;
    syncControls();
    renderPreview();
  };
  const onFamilyInput = (): void => {
    applyFamilyValue(fontFamilyValueFromLabel(currentFontLocale(), familyInput.value));
  };
  const onFamilyOption = (event: Event): void => {
    const value = (event as CustomEvent<{ value?: unknown }>).detail?.value;
    if (typeof value !== 'string') return;
    // Picker rows carry their persisted value explicitly.  This matters for
    // imported `游ゴシック Regular`, whose visible label also names the
    // canonical `游ゴシック` body theme option.
    applyFamilyValue(value);
  };
  const onSizeInput = (): void => {
    touch('fontSize');
    if (sizeInput.value === '') {
      draft.fontSize = undefined;
    } else {
      const points = parseLocalizedNumericInput(sizeInput.value, getLocale());
      if (points !== null) {
        draft.fontSize = fontPointsToCssPx(Math.max(1, Math.min(409, points)));
      }
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
  on(outlineCk.input, 'change', () => onEffectChange('fontOutline', outlineCk.input));
  on(shadowCk.input, 'change', () => onEffectChange('fontShadow', shadowCk.input));
  on(normalFontCk.input, 'change', onNormalFontChange);
  on(fontStyleList, 'click', onFontStyleListClick as EventListener);
  on(familyInput, 'input', onFamilyInput);
  on(familyInput, 'fc-font-family-option', onFamilyOption);
  on(sizeInput, 'input', onSizeInput);
  on(colorInput, 'input', onColorInput);
  on(colorReset, 'click', onColorReset);
  on(fontSwatchesToggle, 'click', onFontSwatchesToggle);
  on(fontSwatches.el, 'click', onFontSwatchClick);

  return { sync, syncMixed, palette };
}
