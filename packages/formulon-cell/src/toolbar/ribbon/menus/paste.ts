// Home tab Paste split-button dropdown. Pure label-driven menuButton list, so
// the only dep is `ribbonLang` for the inline ja/en string fork.

import type { ToolbarLang } from '@libraz/formulon-cell';

import { createMenu, menuButton, menuSeparator } from './general.js';

export const createPasteMenu = (ribbonLang: ToolbarLang): HTMLDivElement => {
  const ja = ribbonLang === 'ja';
  const menu = createMenu('menu-paste');
  menu.append(
    menuButton(ja ? '貼り付け' : 'Paste', 'pasteAction', 'all'),
    menuButton(ja ? '数式' : 'Formulas', 'pasteAction', 'formulas'),
    menuButton(
      ja ? '数式と数値の書式' : 'Formulas & Number Formatting',
      'pasteAction',
      'formulas-and-numfmt',
    ),
    menuButton(ja ? '値' : 'Values', 'pasteAction', 'values'),
    menuButton(
      ja ? '値と数値の書式' : 'Values & Number Formatting',
      'pasteAction',
      'values-and-numfmt',
    ),
    menuButton(ja ? '書式設定' : 'Formatting', 'pasteAction', 'formats'),
    menuSeparator(),
    menuButton(ja ? '行/列の入れ替え' : 'Transpose', 'pasteAction', 'transpose'),
    menuButton(ja ? '形式を選択して貼り付け...' : 'Paste Special...', 'pasteAction', 'dialog'),
  );
  return menu;
};
