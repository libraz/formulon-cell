// Home tab Paste split-button dropdown. Pure label-driven list backed by the
// shared i18n dictionary so toolbar/context-menu copy stays localized in one
// place while the primary ribbon label keeps the Office-like wording.

import type { Strings } from '../../../i18n/strings.js';

import type { ToolbarLang } from '../../ribbon-model.js';
import { createMenu, menuIconButton, menuSeparator } from './general.js';
import { applyMenuVariant, type MenuVariant } from './mac-variants.js';

export const createPasteMenu = (
  t: Strings,
  lang: ToolbarLang,
  variant?: MenuVariant,
): HTMLDivElement => {
  const pasteText = t.contextMenu;
  const menu = createMenu('menu-paste');
  menu.append(
    menuIconButton(t.ribbon.paste, 'pasteAction', 'all', 'paste-all'),
    menuIconButton(pasteText.pasteFormulas, 'pasteAction', 'formulas', 'paste-formulas'),
    menuIconButton(
      pasteText.pasteFormulasNumFmt,
      'pasteAction',
      'formulas-and-numfmt',
      'paste-formulas-numfmt',
    ),
    menuIconButton(pasteText.pasteValues, 'pasteAction', 'values', 'paste-values'),
    menuIconButton(
      pasteText.pasteValuesNumFmt,
      'pasteAction',
      'values-and-numfmt',
      'paste-values-numfmt',
    ),
    menuIconButton(pasteText.pasteFormatsOnly, 'pasteAction', 'formats', 'paste-formats'),
    menuSeparator(),
    menuIconButton(pasteText.pasteTranspose, 'pasteAction', 'transpose', 'paste-transpose'),
    menuIconButton(pasteText.pasteSpecialDialog, 'pasteAction', 'dialog', 'paste-special'),
  );
  if (variant?.mac) {
    // With only text on the clipboard the Mac menu offers these two ahead of the dialog row.
    const dialog = menu.querySelector('[data-paste-action="dialog"]');
    const textSeparator = menuSeparator();
    textSeparator.dataset.pasteText = 'true';
    const textRows = [
      menuIconButton(
        lang === 'ja' ? 'テキストのみ保持' : 'Keep Text Only',
        'pasteAction',
        'text-only',
        'paste-values',
      ),
      textSeparator,
      menuIconButton(
        lang === 'ja' ? 'テキスト ファイル ウィザードを使用する...' : 'Use Text Import Wizard...',
        'pasteAction',
        'text-wizard',
        'paste-values',
      ),
    ];
    for (const row of textRows) menu.insertBefore(row, dialog);
  }
  return applyMenuVariant(menu, lang, variant);
};
