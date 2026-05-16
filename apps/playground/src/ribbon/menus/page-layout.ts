// Page-Layout tab menus: PrintArea / PageBreaks / SheetBackground / PrintTitles
// / PageTheme. Plain static label menus extracted from main.ts; the parent wires
// up click handlers via the data-* attributes on the buttons.

import type { ToolbarMenuText } from '@libraz/formulon-cell';

import { createMenu, menuButton, menuSeparator } from './general.js';

export interface PageLayoutMenuFactories {
  createPrintAreaMenu: () => HTMLDivElement;
  createPageBreaksMenu: () => HTMLDivElement;
  createSheetBackgroundMenu: () => HTMLDivElement;
  createPrintTitlesMenu: () => HTMLDivElement;
  createPageThemeMenu: () => HTMLDivElement;
}

export const createPageLayoutMenuFactories = (
  ribbonMenuText: ToolbarMenuText,
): PageLayoutMenuFactories => {
  const t = ribbonMenuText;

  const createPrintAreaMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-print-area');
    menu.append(
      menuButton(t.printAreaSet, 'printAreaAction', 'set'),
      menuButton(t.printAreaClear, 'printAreaAction', 'clear'),
    );
    return menu;
  };

  const createPageBreaksMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-page-breaks');
    menu.append(
      menuButton(t.pageBreakInsertRow, 'pageBreakAction', 'insert-row'),
      menuButton(t.pageBreakInsertCol, 'pageBreakAction', 'insert-col'),
      menuSeparator(),
      menuButton(t.pageBreakRemoveRow, 'pageBreakAction', 'remove-row'),
      menuButton(t.pageBreakRemoveCol, 'pageBreakAction', 'remove-col'),
      menuSeparator(),
      menuButton(t.pageBreakResetAll, 'pageBreakAction', 'reset-all'),
    );
    return menu;
  };

  const createSheetBackgroundMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-sheet-background');
    menu.append(
      menuButton(t.sheetBackgroundSet, 'sheetBackgroundAction', 'set'),
      menuButton(t.sheetBackgroundClear, 'sheetBackgroundAction', 'clear'),
    );
    return menu;
  };

  const createPrintTitlesMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-print-titles');
    menu.append(
      menuButton(t.printTitleRowsSet, 'printTitlesAction', 'rows'),
      menuButton(t.printTitleColsSet, 'printTitlesAction', 'cols'),
      menuSeparator(),
      menuButton(t.printTitlesClear, 'printTitlesAction', 'clear'),
    );
    return menu;
  };

  const createPageThemeMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-page-theme');
    menu.append(
      menuButton(t.themePaper, 'pageThemeAction', 'light'),
      menuButton(t.themeInk, 'pageThemeAction', 'dark'),
      menuButton(t.themeContrast, 'pageThemeAction', 'contrast'),
    );
    return menu;
  };

  return {
    createPrintAreaMenu,
    createPageBreaksMenu,
    createSheetBackgroundMenu,
    createPrintTitlesMenu,
    createPageThemeMenu,
  };
};
