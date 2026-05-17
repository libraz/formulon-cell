// Insert tab menus: Symbol grid, PivotTable, Chart/Picture/Shapes/Screenshot,
// Links, DataValidation, DefinedNames, Script, AddIns, PDF. Each factory is a
// static label list extracted from main.ts so the entry file no longer has to
// hold them inline.

import type { ToolbarMenuText } from '@libraz/formulon-cell';

import { createMenu, menuButton, menuIdForCommand, menuSeparator } from './general.js';

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

const symbolMenuHeading = (label: string): HTMLDivElement => {
  const heading = document.createElement('div');
  heading.className = 'app__menu-heading';
  heading.setAttribute('role', 'presentation');
  heading.textContent = label;
  return heading;
};

export const createInsertMenuFactories = (ribbonMenuText: ToolbarMenuText): InsertMenuFactories => {
  const t = ribbonMenuText;

  const SYMBOL_GROUPS = [
    {
      label: t.symbolMath,
      symbols: ['±', '×', '÷', '≤', '≥', '≠', '≈', '∞', '√', '∑', '∫', 'π'],
    },
    {
      label: t.symbolGreek,
      symbols: ['Α', 'Β', 'Γ', 'Δ', 'Θ', 'Λ', 'Ξ', 'Π', 'Σ', 'Φ', 'Ψ', 'Ω'],
    },
    { label: t.symbolCurrency, symbols: ['$', '€', '¥', '£', '¢', '₩', '₹', '₽'] },
    { label: t.symbolLegal, symbols: ['©', '®', '™', '§', '¶', '†', '‡', '•'] },
  ] as const;

  const createSymbolMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-symbol');
    menu.classList.add('app__menu--symbols');
    for (const group of SYMBOL_GROUPS) {
      menu.append(symbolMenuHeading(group.label));
      for (const symbol of group.symbols) {
        const button = menuButton(symbol, 'symbol', symbol);
        button.title = symbol;
        menu.append(button);
      }
    }
    menu.append(menuSeparator(), menuButton(t.symbolMore, 'symbolAction', 'more'));
    return menu;
  };

  const createPivotTableMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-pivot-table');
    menu.append(
      menuButton(t.pivotTableFromRange, 'pivotTableAction', 'dialog'),
      menuButton(t.recommendedPivotTables, 'pivotTableAction', 'recommended'),
      menuSeparator(),
      menuButton(t.pivotTableNewSheet, 'pivotTableAction', 'new-sheet'),
      menuButton(t.pivotTableExistingSheet, 'pivotTableAction', 'existing-sheet'),
    );
    return menu;
  };

  const createDefinedNamesMenu = (id: string): HTMLDivElement => {
    const menu = createMenu(menuIdForCommand(id));
    menu.append(
      menuButton(t.defineName, 'definedNameAction', 'define'),
      menuButton(t.nameManager, 'definedNameAction', 'manager'),
      menuSeparator(),
      menuButton(t.createFromSelectionTop, 'definedNameAction', 'create-top-row'),
      menuButton(t.createFromSelectionBottom, 'definedNameAction', 'create-bottom-row'),
      menuButton(t.createFromSelectionLeft, 'definedNameAction', 'create-left-column'),
      menuButton(t.createFromSelectionRight, 'definedNameAction', 'create-right-column'),
      menuSeparator(),
      menuButton(t.useInFormula, 'definedNameAction', 'use-formula'),
    );
    return menu;
  };

  const createLinksMenu = (id: string): HTMLDivElement => {
    const menu = createMenu(menuIdForCommand(id));
    menu.append(
      menuButton(t.linkInsertOrEdit, 'linkAction', 'hyperlink'),
      menuButton(t.linkOpen, 'linkAction', 'open'),
      menuButton(t.linkClear, 'linkAction', 'clear'),
      menuSeparator(),
      menuButton(t.linkExternalLinks, 'linkAction', 'external'),
    );
    return menu;
  };

  const createDataValidationMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-data-validation');
    menu.append(
      menuButton(t.validationSettings, 'validationAction', 'settings'),
      menuButton(t.validationCircleInvalid, 'validationAction', 'circle-invalid'),
      menuButton(t.validationClearCircles, 'validationAction', 'clear-circles'),
      menuSeparator(),
      menuButton(t.validationClearRules, 'validationAction', 'clear-rules'),
    );
    return menu;
  };

  const createChartInsertMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-chart-insert');
    menu.append(
      menuButton(t.chartColumn, 'chartInsert', 'column'),
      menuButton(t.chartBar, 'chartInsert', 'bar'),
      menuButton(t.chartLine, 'chartInsert', 'line'),
      menuButton(t.chartArea, 'chartInsert', 'area'),
      menuButton(t.chartPie, 'chartInsert', 'pie'),
      menuButton(t.chartScatter, 'chartInsert', 'scatter'),
      menuSeparator(),
      menuButton(t.recommendedCharts, 'chartInsert', 'recommended'),
    );
    return menu;
  };

  const createPictureInsertMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-picture-insert');
    menu.append(
      menuButton(t.pictureThisDevice, 'pictureInsert', 'device'),
      menuButton(t.pictureOnline, 'pictureInsert', 'online'),
    );
    return menu;
  };

  const createShapesInsertMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-shapes-insert');
    menu.append(
      menuButton(t.shapeRectangle, 'shapeInsert', 'rectangle'),
      menuButton(t.shapeRoundedRectangle, 'shapeInsert', 'rounded-rectangle'),
      menuButton(t.shapeOval, 'shapeInsert', 'oval'),
      menuSeparator(),
      menuButton(t.shapeLine, 'shapeInsert', 'line'),
      menuButton(t.shapeArrow, 'shapeInsert', 'arrow'),
    );
    return menu;
  };

  const createScreenshotInsertMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-screenshot-insert');
    menu.append(
      menuButton(t.screenshotCurrentView, 'screenshotInsert', 'current-view'),
      menuButton(t.screenshotScreenClipping, 'screenshotInsert', 'screen-clipping'),
    );
    return menu;
  };

  const createScriptMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-script');
    menu.append(
      menuButton(t.scriptCommandUppercase, 'scriptAction', 'uppercase'),
      menuButton(t.scriptCommandLowercase, 'scriptAction', 'lowercase'),
      menuButton(t.scriptCommandTrim, 'scriptAction', 'trim'),
      menuButton(t.scriptCommandClear, 'scriptAction', 'clear'),
      menuSeparator(),
      menuButton(t.scriptRunCustom, 'scriptAction', 'custom'),
    );
    return menu;
  };

  const createAddInMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-add-ins');
    menu.append(
      menuButton(t.addInGet, 'addInAction', 'get'),
      menuButton(t.addInMy, 'addInAction', 'my'),
      menuSeparator(),
      menuButton(t.addInManage, 'addInAction', 'manage'),
    );
    return menu;
  };

  const createPdfMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-pdf');
    menu.append(
      menuButton(t.pdfCreate, 'pdfAction', 'create'),
      menuButton(t.pdfShare, 'pdfAction', 'share'),
      menuSeparator(),
      menuButton(t.pdfPreferences, 'pdfAction', 'preferences'),
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
