// Formulas tab menus: AutoSum, calc options, formula audit error checks. The
// AutoSum / formula-audit dropdowns are pure label lists; calc options uses a
// mix of menuitem/menuitemradio entries so the parent can toggle the radio
// state after rendering.

import type { ToolbarLang, ToolbarMenuText } from '@libraz/formulon-cell';

import { createMenu, menuButton, menuSeparator } from './general.js';

export type AutoSumFormulaName = 'SUM' | 'AVERAGE' | 'COUNT' | 'MAX' | 'MIN' | 'MORE';

export interface FormulasMenuFactories {
  createAutoSumMenu: (id: string) => HTMLDivElement;
  createCalcOptionsMenu: () => HTMLDivElement;
  createClearArrowsMenu: () => HTMLDivElement;
  createErrorCheckingMenu: () => HTMLDivElement;
}

const calcOptionButton = (label: string, value: string): HTMLButtonElement => {
  const button = menuButton(label, 'calcOption', value);
  if (value === 'auto' || value === 'manual' || value === 'auto-no-table') {
    button.setAttribute('role', 'menuitemradio');
    button.setAttribute('aria-checked', 'false');
  }
  return button;
};

export const createFormulasMenuFactories = (
  ribbonMenuText: ToolbarMenuText,
  ribbonLang: ToolbarLang,
): FormulasMenuFactories => {
  const t = ribbonMenuText;
  const ja = ribbonLang === 'ja';

  const createAutoSumMenu = (id: string): HTMLDivElement => {
    const menu = createMenu(id);
    menu.append(
      menuButton(t.autosumSum, 'autosumFn', 'SUM'),
      menuButton(t.autosumAverage, 'autosumFn', 'AVERAGE'),
      menuButton(t.autosumCount, 'autosumFn', 'COUNT'),
      menuButton(t.autosumMax, 'autosumFn', 'MAX'),
      menuButton(t.autosumMin, 'autosumFn', 'MIN'),
      menuSeparator(),
      menuButton(t.autosumMoreFunctions, 'autosumFn', 'MORE'),
    );
    return menu;
  };

  const createCalcOptionsMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-calc-options');
    menu.append(
      calcOptionButton(ja ? '自動' : 'Automatic', 'auto'),
      calcOptionButton(
        ja ? 'データ テーブル以外自動' : 'Automatic Except for Data Tables',
        'auto-no-table',
      ),
      calcOptionButton(ja ? '手動' : 'Manual', 'manual'),
      menuSeparator(),
      calcOptionButton(ja ? '再計算実行' : 'Calculate Now', 'calculate-now'),
      calcOptionButton(ja ? 'シート再計算' : 'Calculate Sheet', 'calculate-sheet'),
      menuSeparator(),
      calcOptionButton(ja ? '反復計算...' : 'Iterative Calculation...', 'iterative'),
    );
    return menu;
  };

  const createClearArrowsMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-clear-arrows');
    menu.append(
      menuButton(t.removeArrowsAll, 'formulaAuditAction', 'clear-all'),
      menuButton(t.removePrecedentArrows, 'formulaAuditAction', 'clear-precedents'),
      menuButton(t.removeDependentArrows, 'formulaAuditAction', 'clear-dependents'),
    );
    return menu;
  };

  const createErrorCheckingMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-error-checking');
    menu.append(
      menuButton(t.errorChecking, 'formulaAuditAction', 'error-checking'),
      menuButton(t.traceError, 'formulaAuditAction', 'trace-error'),
      menuSeparator(),
      menuButton(t.ignoreError, 'formulaAuditAction', 'ignore-error'),
    );
    return menu;
  };

  return {
    createAutoSumMenu,
    createCalcOptionsMenu,
    createClearArrowsMenu,
    createErrorCheckingMenu,
  };
};
