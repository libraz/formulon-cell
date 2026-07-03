// Home tab Paste split-button dropdown. Pure label-driven list backed by the
// shared i18n dictionary so toolbar/context-menu copy stays localized in one
// place while the primary ribbon label keeps the Office-like wording.

import type { Strings } from '../../../index.js';

import { createMenu, menuIconButton, menuSeparator } from './general.js';

export const createPasteMenu = (t: Strings): HTMLDivElement => {
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
  return menu;
};
