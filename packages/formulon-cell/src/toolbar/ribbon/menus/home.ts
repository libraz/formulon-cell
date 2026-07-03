// Home tab + cross-tab menus: Fill/Clear/Freeze/InsertCells/DeleteCells/
// FormatCells/TextToColumns/Sort/FindSelect. Each factory builds shared icon,
// swatch, or submenu rows so the parent can wire `data-*` click handlers
// without dragging the menu DOM along with it.
//
// Sort/FormatCells/InsertCells/DeleteCells consume more than just
// `ribbonMenuText` — they reference `ribbonText` (cross-tab toolbar dict) and
// `dictionaries[lang].sheetTabs` for the sheet-management entries — so the
// factory takes a `HomeMenuDeps` bundle instead of just the menu dict.

import type { Strings, ToolbarLang, ToolbarMenuText, ToolbarText } from '../../../index.js';
import { SHEET_TAB_COLOR_CHOICES, sheetTabColorChoiceLabel } from '../../../sheet-tab-colors.js';
import {
  colorSwatchButton,
  colorSwatchGrid,
  createMenu,
  createSubmenu,
  menuIconButton,
  menuIconSpacer,
  menuIdForCommand,
  menuPresetButton,
  menuSectionHeader,
  menuSeparator,
  menuSubmenuTrigger,
} from './general.js';

export interface HomeMenuDeps {
  ribbonLang: ToolbarLang;
  ribbonMenuText: ToolbarMenuText;
  ribbonText: ToolbarText;
  formatDialog: Strings['formatDialog'];
  sheetTabs: Strings['sheetTabs'];
  viewToolbar: Strings['viewToolbar'];
}

export interface HomeMenuFactories {
  createCopyMenu: () => HTMLDivElement;
  createUnderlineMenu: () => HTMLDivElement;
  createWrapMenu: () => HTMLDivElement;
  createMergeMenu: () => HTMLDivElement;
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

type HomeMenuText = ToolbarMenuText & {
  copyAsPicture: string;
  insertCells: string;
  deleteCells: string;
  underlineSingle: string;
  underlineDouble: string;
  objectSelect: string;
  selectionPane: string;
};

const formatSubmenuId = (key: 'visibility' | 'tabColor'): string => `menu-format-cells-${key}`;

const formatSubmenuTrigger = (
  label: string,
  key: 'visibility' | 'tabColor',
  icon: string,
): HTMLButtonElement =>
  menuSubmenuTrigger(menuIconButton(label, 'formatSubmenu', key, icon), undefined, {
    controlsId: formatSubmenuId(key),
  });

export const createHomeMenuFactories = (deps: HomeMenuDeps): HomeMenuFactories => {
  const { formatDialog, ribbonMenuText, ribbonText, sheetTabs, viewToolbar } = deps;
  const t = ribbonMenuText as HomeMenuText;

  const createTabColorGrid = (): HTMLDivElement => {
    const grid = colorSwatchGrid('app__color-swatch-grid--tab-color');
    for (const entry of SHEET_TAB_COLOR_CHOICES.filter((choice) => choice.color !== null)) {
      grid.append(
        colorSwatchButton({
          label: `${sheetTabs.tabColor}: ${sheetTabColorChoiceLabel(entry, sheetTabs)}`,
          attr: 'cellFormat',
          value: entry.action,
          color: entry.color,
        }),
      );
    }
    return grid;
  };

  const createTabColorPalette = (): HTMLDivElement => {
    const palette = document.createElement('div');
    palette.className = 'app__format-tab-color-palette';
    const highContrast = menuPresetButton(
      t.formatHighContrastOnly,
      'cellFormat',
      'tab-color-high-contrast',
      menuIconSpacer(),
    );
    highContrast.classList.add('app__format-tab-color-high-contrast');
    const noColor = menuPresetButton(
      sheetTabs.noColor,
      'cellFormat',
      'tab-color-none',
      menuIconSpacer(),
    );
    noColor.classList.add('app__format-tab-color-none');
    const moreColors = menuPresetButton(
      formatDialog.moreColors,
      'cellFormat',
      'tab-color-more',
      menuIconSpacer(),
    );
    moreColors.classList.add('app__format-tab-color-more');
    palette.append(
      highContrast,
      noColor,
      menuSectionHeader(formatDialog.themeColors),
      createTabColorGrid(),
      menuSectionHeader(formatDialog.standardColors),
      createTabColorGrid(),
      moreColors,
    );
    return palette;
  };

  const createFreezeMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-freeze');
    menu.append(
      menuIconButton(viewToolbar.freezePanes, 'freeze', 'selection', 'freeze-panes'),
      menuIconButton(viewToolbar.freezeTopRow, 'freeze', 'row', 'freeze-row'),
      menuIconButton(viewToolbar.freezeFirstColumn, 'freeze', 'col', 'freeze-col'),
    );
    return menu;
  };

  const createMergeMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-merge');
    menu.append(
      menuIconButton(ribbonText.mergeAndCenter, 'mergeAction', 'mergeCenter', 'merge'),
      menuIconButton(ribbonText.mergeAcross, 'mergeAction', 'mergeAcross', 'merge'),
      menuIconButton(ribbonText.mergeCells, 'mergeAction', 'mergeCells', 'merge'),
      menuIconButton(ribbonText.unmergeCells, 'mergeAction', 'unmergeCells', 'merge'),
    );
    return menu;
  };

  const createCopyMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-copy');
    menu.append(
      menuIconButton(ribbonText.copy, 'copyAction', 'copy', 'copy'),
      menuIconButton(ribbonMenuText.copyAsPicture, 'copyAction', 'picture', 'picture'),
    );
    return menu;
  };

  const createUnderlineMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-underline');
    menu.append(
      menuIconButton(t.underlineSingle, 'underlineAction', 'single', 'underline-single'),
      menuIconButton(t.underlineDouble, 'underlineAction', 'double', 'underline-double'),
    );
    return menu;
  };

  const createWrapMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-wrap');
    menu.append(
      menuIconButton(ribbonText.wrapText, 'wrapAction', 'wrapText', 'wrap'),
      menuPresetButton(formatDialog.shrinkToFit, 'wrapAction', 'shrinkToFit', menuIconSpacer()),
    );
    return menu;
  };

  const createFillMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-fill');
    menu.append(
      menuIconButton(t.fillDown, 'fill', 'down', 'fill-down'),
      menuIconButton(t.fillRight, 'fill', 'right', 'fill-right'),
      menuIconButton(t.fillUp, 'fill', 'up', 'fill-up'),
      menuIconButton(t.fillLeft, 'fill', 'left', 'fill-left'),
      menuSeparator(),
      menuIconButton(t.fillGroup, 'fill', 'group', 'fill-group'),
      menuIconButton(t.series, 'fill', 'series', 'fill-series'),
      menuIconButton(t.fillJustify, 'fill', 'justify', 'fill-justify'),
      menuIconButton(t.flashFill, 'fill', 'flash', 'flash-fill'),
    );
    return menu;
  };

  const createClearMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-clear');
    menu.append(
      menuIconButton(t.clearAll, 'clear', 'all', 'clear-all'),
      menuIconButton(t.clearFormats, 'clear', 'formats', 'clear-formats'),
      menuIconButton(t.clearContents, 'clear', 'contents', 'clear-contents'),
      menuIconButton(t.clearComments, 'clear', 'comments', 'clear-comments'),
      menuIconButton(t.clearHyperlinks, 'clear', 'hyperlinks', 'clear-hyperlinks'),
      menuIconButton(t.removeHyperlinks, 'clear', 'remove-hyperlinks', 'clear-hyperlinks'),
      menuIconButton(t.clearConditional, 'clear', 'conditional', 'clear-conditional'),
    );
    return menu;
  };

  const createInsertCellsMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-insert-cells');
    menu.append(
      menuIconButton(t.insertCells, 'cellInsert', 'cells', 'insert-cells'),
      menuSeparator(),
      menuIconButton(t.insertRows, 'cellInsert', 'rows', 'insert-row'),
      menuIconButton(t.insertCols, 'cellInsert', 'cols', 'insert-col'),
      menuSeparator(),
      menuIconButton(sheetTabs.insertSheet, 'cellInsert', 'sheet', 'insert-sheet'),
    );
    return menu;
  };

  const createDeleteCellsMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-delete-cells');
    menu.append(
      menuIconButton(t.deleteCells, 'cellDelete', 'cells', 'delete-cells'),
      menuSeparator(),
      menuIconButton(t.deleteRows, 'cellDelete', 'rows', 'delete-row'),
      menuIconButton(t.deleteCols, 'cellDelete', 'cols', 'delete-col'),
      menuSeparator(),
      menuIconButton(t.deleteRow, 'cellDelete', 'row', 'delete-row'),
      menuIconButton(t.deleteCol, 'cellDelete', 'col', 'delete-col'),
      menuSeparator(),
      menuIconButton(t.deleteSheet, 'cellDelete', 'sheet', 'delete-sheet'),
    );
    return menu;
  };

  const createFormatCellsMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-format-cells');
    const visibilitySubmenu = createSubmenu({
      id: formatSubmenuId('visibility'),
      className: 'app__submenu app__submenu--format app__submenu--format-visibility',
      label: t.formatHideUnhide,
      dataset: { formatPanel: 'visibility' },
    });
    visibilitySubmenu.append(
      menuIconButton(t.hideRows, 'cellFormat', 'hide-rows', 'format-hide-rows'),
      menuIconButton(t.hideCols, 'cellFormat', 'hide-cols', 'format-hide-cols'),
      menuIconButton(t.hideSheet, 'cellFormat', 'hide-sheet', 'format-hide-sheet'),
      menuSeparator(),
      menuIconButton(t.showRows, 'cellFormat', 'show-rows', 'format-show-rows'),
      menuIconButton(t.showCols, 'cellFormat', 'show-cols', 'format-show-cols'),
      menuIconButton(t.showSheet, 'cellFormat', 'unhide-sheet', 'format-unhide-sheet'),
    );
    const tabColorSubmenu = createSubmenu({
      id: formatSubmenuId('tabColor'),
      className: 'app__submenu app__submenu--format app__submenu--format-tab-color',
      label: t.formatSheetTabColor,
      dataset: { formatPanel: 'tabColor' },
    });
    tabColorSubmenu.append(createTabColorPalette());
    menu.append(
      menuSectionHeader(t.formatCellSize),
      menuIconButton(t.rowHeight, 'cellFormat', 'row-height', 'format-row-height'),
      menuIconButton(t.autoFitRowHeight, 'cellFormat', 'row-autofit', 'format-row-autofit'),
      menuIconButton(t.colWidth, 'cellFormat', 'col-width', 'format-col-width'),
      menuIconButton(t.autoFitColWidth, 'cellFormat', 'col-autofit', 'format-col-autofit'),
      menuSeparator(),
      menuSectionHeader(t.formatVisibility),
      formatSubmenuTrigger(t.formatHideUnhide, 'visibility', 'format-hide-rows'),
      visibilitySubmenu,
      menuSeparator(),
      menuSectionHeader(t.formatSheetOrganization),
      menuIconButton(t.renameSheet, 'cellFormat', 'rename-sheet', 'format-rename-sheet'),
      menuIconButton(t.moveOrCopySheet, 'cellFormat', 'move-sheet-copy', 'format-move-copy'),
      formatSubmenuTrigger(t.formatSheetTabColor, 'tabColor', 'format-dialog'),
      tabColorSubmenu,
      menuSeparator(),
      menuSectionHeader(t.formatProtection),
      menuIconButton(t.lockCell, 'cellFormat', 'lock-cell', 'format-lock'),
      menuIconButton(t.unlockCell, 'cellFormat', 'unlock-cell', 'format-unlock'),
      menuIconButton(t.protectSheet, 'cellFormat', 'protect-sheet', 'format-protect'),
      menuSeparator(),
      menuIconButton(t.formatCells, 'cellFormat', 'dialog', 'format-dialog'),
    );
    return menu;
  };

  const createSortMenu = (id: string): HTMLDivElement => {
    const menu = createMenu(menuIdForCommand(id));
    menu.append(
      menuIconButton(t.sortAscendingMenu, 'sort', 'asc', 'sort-asc'),
      menuIconButton(t.sortDescendingMenu, 'sort', 'desc', 'sort-desc'),
      menuIconButton(t.sortCustom, 'sort', 'custom', 'sort-custom'),
      menuSeparator(),
      menuIconButton(t.filterToggle, 'sort', 'filter', 'filter-toggle'),
      menuIconButton(t.filterBySelectedCellValue, 'sort', 'filter-by-value', 'filter-by-value'),
      menuIconButton(t.filterClearAll, 'sort', 'filter-clear', 'filter-clear'),
      menuIconButton(t.filterReapply, 'sort', 'filter-reapply', 'filter-reapply'),
      menuIconButton(t.filterAdvanced, 'sort', 'filter-advanced', 'filter-advanced'),
      menuSeparator(),
      menuIconButton(ribbonText.removeDuplicates, 'sort', 'dedupe', 'remove-duplicates'),
      menuIconButton(ribbonText.conditionalFormatting, 'sort', 'conditional', 'conditional'),
      menuIconButton(t.nameManager, 'sort', 'named', 'name-manager'),
    );
    return menu;
  };

  const createTextToColumnsMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-text-to-columns');
    menu.append(
      menuIconButton(t.textToColumnsComma, 'textToColumnsDelimiter', ',', 'text-column-comma'),
      menuIconButton(t.textToColumnsTab, 'textToColumnsDelimiter', '\\t', 'text-column-tab'),
      menuIconButton(
        t.textToColumnsSemicolon,
        'textToColumnsDelimiter',
        ';',
        'text-column-semicolon',
      ),
      menuIconButton(t.textToColumnsSpace, 'textToColumnsDelimiter', ' ', 'text-column-space'),
      menuSeparator(),
      menuIconButton(
        t.textToColumnsCustom,
        'textToColumnsDelimiter',
        'custom',
        'text-column-custom',
      ),
    );
    return menu;
  };

  const createFindSelectMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-find-select');
    menu.append(
      menuIconButton(t.find, 'findSelect', 'find', 'find'),
      menuIconButton(t.replace, 'findSelect', 'replace', 'replace'),
      menuIconButton(t.goTo, 'findSelect', 'go-to', 'go-to'),
      menuIconButton(t.goToSpecial, 'findSelect', 'go-to-special', 'go-to-special'),
      menuSeparator(),
      menuIconButton(t.findFormulas, 'findSelect', 'formulas', 'find-formulas'),
      menuIconButton(t.comments, 'findSelect', 'comments', 'find-comments'),
      menuIconButton(
        t.findConditionalFormatting,
        'findSelect',
        'conditional-format',
        'find-conditional',
      ),
      menuIconButton(t.findConstants, 'findSelect', 'constants', 'find-constants'),
      menuIconButton(t.findDataValidation, 'findSelect', 'data-validation', 'find-validation'),
      menuSeparator(),
      menuIconButton(t.objectSelect, 'findSelect', 'object-select', 'object-select'),
      menuIconButton(t.selectionPane, 'findSelect', 'selection-pane', 'selection-pane'),
    );
    return menu;
  };

  return {
    createCopyMenu,
    createUnderlineMenu,
    createWrapMenu,
    createMergeMenu,
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
