// Home tab + cross-tab static menus: Fill/Clear/Freeze/InsertCells/DeleteCells/
// FormatCells/TextToColumns/Sort/FindSelect. Each factory is a static label
// list extracted from main.ts so the parent can wire `data-*` click handlers
// without dragging the menu DOM along with it.
//
// Sort/FormatCells/InsertCells/DeleteCells consume more than just
// `ribbonMenuText` — they reference `ribbonText` (cross-tab toolbar dict) and
// `dictionaries[lang].sheetTabs` for the sheet-management entries — so the
// factory takes a `HomeMenuDeps` bundle instead of just the menu dict.

import type { Strings, ToolbarLang, ToolbarMenuText, ToolbarText } from '@libraz/formulon-cell';

import { createMenu, menuButton, menuSeparator } from './general.js';

export interface HomeMenuDeps {
  ribbonLang: ToolbarLang;
  ribbonMenuText: ToolbarMenuText;
  ribbonText: ToolbarText;
  sheetTabs: Strings['sheetTabs'];
}

export interface HomeMenuFactories {
  createFreezeMenu: () => HTMLDivElement;
  createFillMenu: () => HTMLDivElement;
  createClearMenu: () => HTMLDivElement;
  createInsertCellsMenu: () => HTMLDivElement;
  createDeleteCellsMenu: () => HTMLDivElement;
  createFormatCellsMenu: () => HTMLDivElement;
  createSortMenu: (id: string) => HTMLDivElement;
  createTextToColumnsMenu: () => HTMLDivElement;
  createFindSelectMenu: () => HTMLDivElement;
}

export const createHomeMenuFactories = (deps: HomeMenuDeps): HomeMenuFactories => {
  const { ribbonLang, ribbonMenuText: t, ribbonText, sheetTabs } = deps;
  const ja = ribbonLang === 'ja';

  const createFreezeMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-freeze');
    menu.append(
      menuButton(ja ? '先頭行の固定' : 'Freeze Top Row', 'freeze', 'row'),
      menuButton(ja ? '先頭列の固定' : 'Freeze First Column', 'freeze', 'col'),
      menuButton(ja ? 'ウィンドウ枠の固定' : 'Freeze Panes', 'freeze', 'selection'),
      menuButton(ja ? 'ウィンドウ枠固定の解除' : 'Unfreeze Panes', 'freeze', 'off'),
    );
    return menu;
  };

  const createFillMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-fill');
    menu.append(
      menuButton(t.fillDown, 'fill', 'down'),
      menuButton(t.fillRight, 'fill', 'right'),
      menuButton(t.fillUp, 'fill', 'up'),
      menuButton(t.fillLeft, 'fill', 'left'),
      menuSeparator(),
      menuButton(t.series, 'fill', 'series'),
      menuSeparator(),
      menuButton(t.fillDays, 'fill', 'days'),
      menuButton(t.fillWeekdays, 'fill', 'weekdays'),
      menuButton(t.fillMonths, 'fill', 'months'),
      menuButton(t.fillYears, 'fill', 'years'),
    );
    return menu;
  };

  const createClearMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-clear');
    menu.append(
      menuButton(t.clearAll, 'clear', 'all'),
      menuButton(t.clearFormats, 'clear', 'formats'),
      menuButton(t.clearContents, 'clear', 'contents'),
      menuButton(t.clearComments, 'clear', 'comments'),
      menuButton(t.clearHyperlinks, 'clear', 'hyperlinks'),
      menuButton(t.removeHyperlinks, 'clear', 'remove-hyperlinks'),
      menuButton(t.clearConditional, 'clear', 'conditional'),
    );
    return menu;
  };

  const createInsertCellsMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-insert-cells');
    menu.append(
      menuButton(ja ? 'セルを挿入...' : 'Insert Cells...', 'cellInsert', 'cells'),
      menuButton(t.insertShiftDown, 'cellInsert', 'shift-down'),
      menuButton(t.insertShiftRight, 'cellInsert', 'shift-right'),
      menuSeparator(),
      menuButton(ja ? 'シートの行を挿入' : 'Insert Sheet Rows', 'cellInsert', 'rows'),
      menuButton(ja ? 'シートの列を挿入' : 'Insert Sheet Columns', 'cellInsert', 'cols'),
      menuSeparator(),
      menuButton(sheetTabs.insertSheet, 'cellInsert', 'sheet'),
    );
    return menu;
  };

  const createDeleteCellsMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-delete-cells');
    menu.append(
      menuButton(ja ? 'セルを削除...' : 'Delete Cells...', 'cellDelete', 'cells'),
      menuButton(t.deleteShiftUp, 'cellDelete', 'shift-up'),
      menuButton(t.deleteShiftLeft, 'cellDelete', 'shift-left'),
      menuSeparator(),
      menuButton(ja ? 'シートの行を削除' : 'Delete Sheet Rows', 'cellDelete', 'rows'),
      menuButton(ja ? 'シートの列を削除' : 'Delete Sheet Columns', 'cellDelete', 'cols'),
      menuSeparator(),
      menuButton(sheetTabs.deleteSheet, 'cellDelete', 'sheet'),
    );
    return menu;
  };

  const createFormatCellsMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-format-cells');
    menu.append(
      menuButton(t.formatCells, 'cellFormat', 'dialog'),
      menuSeparator(),
      menuButton(t.rowHeight, 'cellFormat', 'row-height'),
      menuButton(t.autoFitRowHeight, 'cellFormat', 'row-autofit'),
      menuButton(t.colWidth, 'cellFormat', 'col-width'),
      menuButton(t.autoFitColWidth, 'cellFormat', 'col-autofit'),
      menuSeparator(),
      menuButton(t.hideRows, 'cellFormat', 'hide-rows'),
      menuButton(t.showRows, 'cellFormat', 'show-rows'),
      menuButton(t.hideCols, 'cellFormat', 'hide-cols'),
      menuButton(t.showCols, 'cellFormat', 'show-cols'),
      menuSeparator(),
      menuButton(sheetTabs.rename, 'cellFormat', 'rename-sheet'),
      menuButton(sheetTabs.moveLeft, 'cellFormat', 'move-sheet-left'),
      menuButton(sheetTabs.moveRight, 'cellFormat', 'move-sheet-right'),
      menuButton(sheetTabs.hideSheet, 'cellFormat', 'hide-sheet'),
      menuButton(sheetTabs.unhideSheet, 'cellFormat', 'unhide-sheet'),
      menuSeparator(),
      menuButton(`${sheetTabs.tabColor}: ${sheetTabs.noColor}`, 'cellFormat', 'tab-color-none'),
      menuButton(`${sheetTabs.tabColor}: ${sheetTabs.tabColorRed}`, 'cellFormat', 'tab-color-red'),
      menuButton(
        `${sheetTabs.tabColor}: ${sheetTabs.tabColorOrange}`,
        'cellFormat',
        'tab-color-orange',
      ),
      menuButton(
        `${sheetTabs.tabColor}: ${sheetTabs.tabColorYellow}`,
        'cellFormat',
        'tab-color-yellow',
      ),
      menuButton(
        `${sheetTabs.tabColor}: ${sheetTabs.tabColorGreen}`,
        'cellFormat',
        'tab-color-green',
      ),
      menuButton(
        `${sheetTabs.tabColor}: ${sheetTabs.tabColorBlue}`,
        'cellFormat',
        'tab-color-blue',
      ),
      menuButton(
        `${sheetTabs.tabColor}: ${sheetTabs.tabColorPurple}`,
        'cellFormat',
        'tab-color-purple',
      ),
      menuButton(
        `${sheetTabs.tabColor}: ${sheetTabs.tabColorGray}`,
        'cellFormat',
        'tab-color-gray',
      ),
      menuSeparator(),
      menuButton(t.lockCell, 'cellFormat', 'lock-cell'),
      menuButton(t.unlockCell, 'cellFormat', 'unlock-cell'),
      menuButton(t.protectSheet, 'cellFormat', 'protect-sheet'),
    );
    return menu;
  };

  const createSortMenu = (id: string): HTMLDivElement => {
    const menu = createMenu(id);
    menu.append(
      menuButton(t.sortAscendingMenu, 'sort', 'asc'),
      menuButton(t.sortDescendingMenu, 'sort', 'desc'),
      menuButton(t.sortCustom, 'sort', 'custom'),
      menuSeparator(),
      menuButton(t.filterToggle, 'sort', 'filter'),
      menuButton(t.filterBySelectedCellValue, 'sort', 'filter-by-value'),
      menuButton(t.filterClearAll, 'sort', 'filter-clear'),
      menuButton(t.filterReapply, 'sort', 'filter-reapply'),
      menuButton(t.filterAdvanced, 'sort', 'filter-advanced'),
      menuSeparator(),
      menuButton(ribbonText.removeDuplicates, 'sort', 'dedupe'),
      menuButton(ribbonText.conditionalFormatting, 'sort', 'conditional'),
      menuButton(t.nameManager, 'sort', 'named'),
    );
    return menu;
  };

  const createTextToColumnsMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-text-to-columns');
    menu.append(
      menuButton(t.textToColumnsComma, 'textToColumnsDelimiter', ','),
      menuButton(t.textToColumnsTab, 'textToColumnsDelimiter', '\\t'),
      menuButton(t.textToColumnsSemicolon, 'textToColumnsDelimiter', ';'),
      menuButton(t.textToColumnsSpace, 'textToColumnsDelimiter', ' '),
      menuSeparator(),
      menuButton(t.textToColumnsCustom, 'textToColumnsDelimiter', 'custom'),
    );
    return menu;
  };

  const createFindSelectMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-find-select');
    menu.append(
      menuButton(t.find, 'findSelect', 'find'),
      menuButton(t.replace, 'findSelect', 'replace'),
      menuButton(t.goTo, 'findSelect', 'go-to'),
      menuButton(t.goToSpecial, 'findSelect', 'go-to-special'),
      menuSeparator(),
      menuButton(t.findFormulas, 'findSelect', 'formulas'),
      menuButton(t.findConstants, 'findSelect', 'constants'),
      menuButton(t.findConditionalFormatting, 'findSelect', 'conditional-format'),
      menuButton(t.findDataValidation, 'findSelect', 'data-validation'),
      menuButton(t.comments, 'findSelect', 'comments'),
    );
    return menu;
  };

  return {
    createFreezeMenu,
    createFillMenu,
    createClearMenu,
    createInsertCellsMenu,
    createDeleteCellsMenu,
    createFormatCellsMenu,
    createSortMenu,
    createTextToColumnsMenu,
    createFindSelectMenu,
  };
};
