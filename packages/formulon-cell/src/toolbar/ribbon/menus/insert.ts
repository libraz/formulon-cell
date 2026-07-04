// Insert tab menus: Symbol grid, PivotTable, Chart/Picture/Shapes/Screenshot,
// Links, DataValidation, DefinedNames, Script, AddIns, PDF. Each factory is a
// shared icon, symbol, or visual-tile menu extracted from main.ts so the entry
// file no longer has to hold them inline.

import type { ToolbarMenuText } from '../../menu-text.js';
import { toolbarSymbolGroups } from '../symbols.js';
import {
  createMenu,
  menuIconButton,
  menuIdForCommand,
  menuSectionHeader,
  menuSeparator,
  symbolMenuGrid,
  visualMenuTile,
  visualMenuTileGrid,
} from './general.js';

export interface InsertMenuFactories {
  createSymbolMenu: () => HTMLDivElement;
  createPivotTableMenu: () => HTMLDivElement;
  createDefinedNamesMenu: (id: string) => HTMLDivElement;
  createLinksMenu: (id: string) => HTMLDivElement;
  createDataValidationMenu: () => HTMLDivElement;
  createChartInsertMenu: () => HTMLDivElement;
  createPictureInsertMenu: () => HTMLDivElement;
  createShapesInsertMenu: () => HTMLDivElement;
  createScreenshotInsertMenu: () => HTMLDivElement;
  createScriptMenu: () => HTMLDivElement;
  createAddInMenu: () => HTMLDivElement;
  createPdfMenu: () => HTMLDivElement;
}

export const createInsertMenuFactories = (ribbonMenuText: ToolbarMenuText): InsertMenuFactories => {
  const t = ribbonMenuText;

  const createSymbolMenu = (): HTMLDivElement => {
    const menu = createMenu(menuIdForCommand('symbolInsert'));
    menu.classList.add('fc-tb__menu--symbols');
    for (const group of toolbarSymbolGroups(t)) {
      menu.append(menuSectionHeader(group.label));
      menu.append(symbolMenuGrid(group.label, group.symbols));
    }
    menu.append(
      menuSeparator(),
      menuIconButton(t.symbolMore, 'symbolAction', 'more', 'symbol-more'),
    );
    return menu;
  };

  const createPivotTableMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-pivot-table');
    menu.append(
      menuIconButton(t.pivotTableFromRange, 'pivotTableAction', 'dialog', 'pivot-range'),
      menuIconButton(
        t.recommendedPivotTables,
        'pivotTableAction',
        'recommended',
        'pivot-recommended',
      ),
      menuIconButton(t.pivotTableRefreshData, 'pivotTableAction', 'refresh', 'pivot-refresh'),
      menuSeparator(),
      menuIconButton(t.pivotTableNewSheet, 'pivotTableAction', 'new-sheet', 'pivot-new-sheet'),
      menuIconButton(
        t.pivotTableExistingSheet,
        'pivotTableAction',
        'existing-sheet',
        'pivot-existing-sheet',
      ),
    );
    return menu;
  };

  const createDefinedNamesMenu = (id: string): HTMLDivElement => {
    const menu = createMenu(menuIdForCommand(id));
    menu.append(
      menuIconButton(t.defineName, 'definedNameAction', 'define', 'defined-name-define'),
      menuIconButton(t.nameManager, 'definedNameAction', 'manager', 'defined-name-manager'),
      menuSeparator(),
      menuIconButton(
        t.createFromSelectionTop,
        'definedNameAction',
        'create-top-row',
        'defined-name-create-top',
      ),
      menuIconButton(
        t.createFromSelectionBottom,
        'definedNameAction',
        'create-bottom-row',
        'defined-name-create-bottom',
      ),
      menuIconButton(
        t.createFromSelectionLeft,
        'definedNameAction',
        'create-left-column',
        'defined-name-create-left',
      ),
      menuIconButton(
        t.createFromSelectionRight,
        'definedNameAction',
        'create-right-column',
        'defined-name-create-right',
      ),
      menuSeparator(),
      menuIconButton(t.useInFormula, 'definedNameAction', 'use-formula', 'defined-name-use'),
    );
    return menu;
  };

  const createLinksMenu = (id: string): HTMLDivElement => {
    const menu = createMenu(menuIdForCommand(id));
    menu.append(
      menuIconButton(t.linkInsertOrEdit, 'linkAction', 'hyperlink', 'link-edit'),
      menuIconButton(t.linkOpen, 'linkAction', 'open', 'link-open'),
      menuIconButton(t.linkClear, 'linkAction', 'clear', 'link-clear'),
      menuSeparator(),
      menuIconButton(t.linkExternalLinks, 'linkAction', 'external', 'link-external'),
    );
    return menu;
  };

  const createDataValidationMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-data-validation');
    menu.append(
      menuIconButton(t.validationSettings, 'validationAction', 'settings', 'validation-settings'),
      menuIconButton(
        t.validationCircleInvalid,
        'validationAction',
        'circle-invalid',
        'validation-circle',
      ),
      menuIconButton(
        t.validationClearCircles,
        'validationAction',
        'clear-circles',
        'validation-clear-circles',
      ),
      menuSeparator(),
      menuIconButton(
        t.validationClearRules,
        'validationAction',
        'clear-rules',
        'validation-clear-rules',
      ),
    );
    return menu;
  };

  const createChartInsertMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-chart-insert');
    menu.classList.add('fc-tb__menu--visual', 'fc-tb__menu--charts');
    const grid = visualMenuTileGrid('fc-tb__visual-grid--charts', [
      {
        label: t.chartColumn,
        attr: 'chartInsert',
        value: 'column',
        icon: 'chart-column',
      },
      {
        label: t.chartBar,
        attr: 'chartInsert',
        value: 'bar',
        icon: 'chart-bar',
      },
      {
        label: t.chartLine,
        attr: 'chartInsert',
        value: 'line',
        icon: 'chart-line',
      },
      {
        label: t.chartArea,
        attr: 'chartInsert',
        value: 'area',
        icon: 'chart-area',
      },
      {
        label: t.chartPie,
        attr: 'chartInsert',
        value: 'pie',
        icon: 'chart-pie',
      },
      {
        label: t.chartScatter,
        attr: 'chartInsert',
        value: 'scatter',
        icon: 'chart-scatter',
      },
    ]);
    menu.append(
      grid,
      menuSeparator(),
      visualMenuTile({
        label: t.recommendedCharts,
        attr: 'chartInsert',
        value: 'recommended',
        icon: 'chart-recommended',
        className: 'fc-tb__visual-tile--wide',
      }),
    );
    return menu;
  };

  const createPictureInsertMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-picture-insert');
    menu.classList.add('fc-tb__menu--visual', 'fc-tb__menu--pictures');
    const grid = visualMenuTileGrid('fc-tb__visual-grid--pictures', [
      {
        label: t.pictureThisDevice,
        attr: 'pictureInsert',
        value: 'device',
        icon: 'device-picture',
      },
      {
        label: t.pictureOnline,
        attr: 'pictureInsert',
        value: 'online',
        icon: 'online-picture',
      },
      {
        label: t.pictureStock,
        attr: 'pictureInsert',
        value: 'stock',
        icon: 'stock-picture',
      },
    ]);
    menu.append(grid);
    return menu;
  };

  const createShapesInsertMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-shapes-insert');
    menu.classList.add('fc-tb__menu--visual', 'fc-tb__menu--shapes');
    const lines = visualMenuTileGrid('fc-tb__visual-grid--shapes', [
      {
        label: t.shapeLine,
        attr: 'shapeInsert',
        value: 'line',
        icon: 'shape-line',
      },
      {
        label: t.shapeArrow,
        attr: 'shapeInsert',
        value: 'arrow',
        icon: 'shape-arrow',
      },
    ]);
    const rectangles = visualMenuTileGrid('fc-tb__visual-grid--shapes', [
      {
        label: t.shapeRectangle,
        attr: 'shapeInsert',
        value: 'rectangle',
        icon: 'shape-rectangle',
      },
      {
        label: t.shapeRoundedRectangle,
        attr: 'shapeInsert',
        value: 'rounded-rectangle',
        icon: 'shape-rounded-rectangle',
      },
    ]);
    const basic = visualMenuTileGrid('fc-tb__visual-grid--shapes', [
      {
        label: t.shapeOval,
        attr: 'shapeInsert',
        value: 'oval',
        icon: 'shape-oval',
      },
      {
        label: t.shapeTriangle,
        attr: 'shapeInsert',
        value: 'triangle',
        icon: 'shape-triangle',
      },
      {
        label: t.shapeDiamond,
        attr: 'shapeInsert',
        value: 'diamond',
        icon: 'shape-diamond',
      },
    ]);
    menu.append(
      menuSectionHeader(t.shapeLines),
      lines,
      menuSectionHeader(t.shapeRectangles),
      rectangles,
      menuSectionHeader(t.shapeBasicShapes),
      basic,
    );
    return menu;
  };

  const createScreenshotInsertMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-screenshot-insert');
    menu.classList.add('fc-tb__menu--visual', 'fc-tb__menu--screenshots');
    const grid = visualMenuTileGrid('fc-tb__visual-grid--screenshots', [
      {
        label: t.screenshotCurrentView,
        attr: 'screenshotInsert',
        value: 'current-view',
        icon: 'screenshot-window',
        className: 'fc-tb__visual-tile--screenshot-preview',
      },
    ]);
    menu.append(
      menuSectionHeader(t.screenshotAvailableWindows),
      grid,
      menuSeparator(),
      visualMenuTile({
        label: t.screenshotScreenClipping,
        attr: 'screenshotInsert',
        value: 'screen-clipping',
        icon: 'screen-clipping',
        className: 'fc-tb__visual-tile--wide',
      }),
    );
    return menu;
  };

  const createScriptMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-script');
    menu.append(
      menuIconButton(t.scriptCommandUppercase, 'scriptAction', 'uppercase', 'script-uppercase'),
      menuIconButton(t.scriptCommandLowercase, 'scriptAction', 'lowercase', 'script-lowercase'),
      menuIconButton(t.scriptCommandTrim, 'scriptAction', 'trim', 'script-trim'),
      menuIconButton(t.scriptCommandClear, 'scriptAction', 'clear', 'script-clear'),
      menuSeparator(),
      menuIconButton(t.scriptRunCustom, 'scriptAction', 'custom', 'script-custom'),
    );
    return menu;
  };

  const createAddInMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-add-ins');
    menu.append(
      menuIconButton(t.addInGet, 'addInAction', 'get', 'addin-get'),
      menuIconButton(t.addInMy, 'addInAction', 'my', 'addin-my'),
      menuSeparator(),
      menuIconButton(t.addInManage, 'addInAction', 'manage', 'addin-manage'),
    );
    return menu;
  };

  const createPdfMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-pdf');
    menu.append(
      menuIconButton(t.pdfCreate, 'pdfAction', 'create', 'pdf-create'),
      menuIconButton(t.pdfShare, 'pdfAction', 'share', 'pdf-share'),
      menuSeparator(),
      menuIconButton(t.pdfPreferences, 'pdfAction', 'preferences', 'pdf-preferences'),
    );
    return menu;
  };

  return {
    createSymbolMenu,
    createPivotTableMenu,
    createDefinedNamesMenu,
    createLinksMenu,
    createDataValidationMenu,
    createChartInsertMenu,
    createPictureInsertMenu,
    createShapesInsertMenu,
    createScreenshotInsertMenu,
    createScriptMenu,
    createAddInMenu,
    createPdfMenu,
  };
};
