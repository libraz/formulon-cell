// Page-Layout tab menus: PrintArea / PageBreaks / SheetBackground / PrintTitles
// / PageTheme. Uses shared menu/icon/visual-tile primitives; the parent wires
// up click handlers via the data-* attributes on the buttons.

import type { ToolbarMenuText } from '../../menu-text.js';

import { createMenu, menuIconButton, menuSeparator, visualMenuTileGrid } from './general.js';

export interface PageLayoutMenuFactories {
  createPrintAreaMenu: () => HTMLDivElement;
  createPageBreaksMenu: () => HTMLDivElement;
  createPageThemeMenu: () => HTMLDivElement;
  createArrangeMenu: () => HTMLDivElement;
}

export const createPageLayoutMenuFactories = (
  ribbonMenuText: ToolbarMenuText,
): PageLayoutMenuFactories => {
  const t = ribbonMenuText;

  const createPrintAreaMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-print-area');
    menu.append(
      menuIconButton(t.printAreaSet, 'printAreaAction', 'set', 'print-area-set'),
      menuIconButton(t.printAreaAdd, 'printAreaAction', 'add', 'print-area-add'),
      menuIconButton(t.printAreaClear, 'printAreaAction', 'clear', 'clear'),
    );
    return menu;
  };

  const createPageBreaksMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-page-breaks');
    menu.append(
      menuIconButton(t.pageBreakInsert, 'pageBreakAction', 'insert', 'break-page'),
      menuIconButton(t.pageBreakRemove, 'pageBreakAction', 'remove', 'break-remove'),
      menuSeparator(),
      menuIconButton(t.pageBreakResetAll, 'pageBreakAction', 'reset-all', 'reset'),
    );
    return menu;
  };

  const createPageThemeMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-page-theme');
    menu.classList.add('fc-tb__menu--visual', 'fc-tb__menu--themes');
    const grid = visualMenuTileGrid('fc-tb__visual-grid--themes', [
      {
        label: t.themePaper,
        attr: 'pageThemeAction',
        value: 'paper',
        icon: 'theme-light',
      },
      {
        label: t.themeInk,
        attr: 'pageThemeAction',
        value: 'ink',
        icon: 'theme-dark',
      },
      {
        label: t.themeContrast,
        attr: 'pageThemeAction',
        value: 'contrast',
        icon: 'theme-contrast',
      },
    ]);
    // The three tiles above are the host's paper/ink/contrast chrome themes.
    // Workbook themes are separate actions: they edit the OOXML clrScheme
    // and typefaces and therefore follow the file when it is saved.
    menu.append(
      grid,
      menuSeparator(),
      menuIconButton(t.themeWorkbookOffice, 'workbookThemeAction', 'office', 'table-style-new'),
      menuIconButton(t.themeWorkbookGray, 'workbookThemeAction', 'gray', 'format-dialog'),
    );
    return menu;
  };

  const createArrangeMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-arrange-objects');
    menu.append(
      menuIconButton(t.arrangeBringForward, 'arrangeAction', 'bring-forward', 'bring-forward'),
      menuIconButton(t.arrangeSendBackward, 'arrangeAction', 'send-backward', 'send-backward'),
      menuSeparator(),
      menuIconButton(t.arrangeBringToFront, 'arrangeAction', 'bring-front', 'bring-front'),
      menuIconButton(t.arrangeSendToBack, 'arrangeAction', 'send-back', 'send-back'),
      menuSeparator(),
      menuIconButton(t.arrangeSelectionPane, 'arrangeAction', 'selection-pane', 'pane'),
    );
    return menu;
  };

  return {
    createPrintAreaMenu,
    createPageBreaksMenu,
    createPageThemeMenu,
    createArrangeMenu,
  };
};
