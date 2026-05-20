// Formulas tab menus: AutoSum, calc options, formula audit error checks. The
// factories use shared icon menu rows; calc options uses a mix of
// menuitem/menuitemradio entries so the parent can toggle the radio state after
// rendering.

import type { ToolbarLang, ToolbarMenuText } from '@libraz/formulon-cell';

import { createMenu, menuIconButton, menuIdForCommand, menuSeparator } from './general.js';

export type AutoSumFormulaName = 'SUM' | 'AVERAGE' | 'COUNT' | 'MAX' | 'MIN' | 'MORE';

export interface FormulasMenuFactories {
  createAutoSumMenu: (id: string) => HTMLDivElement;
  createCalcOptionsMenu: () => HTMLDivElement;
  createClearArrowsMenu: () => HTMLDivElement;
  createErrorCheckingMenu: () => HTMLDivElement;
}

type FormulasMenuText = ToolbarMenuText & {
  calcAutomatic: string;
  calcAutoNoTable: string;
  calcManual: string;
  calcCalculateNow: string;
  calcCalculateSheet: string;
  calcIterative: string;
};

const calcOptionButton = (label: string, value: string, icon: string): HTMLButtonElement => {
  const button = menuIconButton(label, 'calcOption', value, icon);
  if (value === 'auto' || value === 'manual' || value === 'auto-no-table') {
    button.setAttribute('role', 'menuitemradio');
    button.setAttribute('aria-checked', 'false');
  }
  return button;
};

export const createFormulasMenuFactories = (
  ribbonMenuText: ToolbarMenuText,
  _ribbonLang: ToolbarLang,
): FormulasMenuFactories => {
  const t = ribbonMenuText as FormulasMenuText;

  const createAutoSumMenu = (id: string): HTMLDivElement => {
    const menu = createMenu(menuIdForCommand(id));
    menu.append(
      menuIconButton(t.autosumSum, 'autosumFn', 'SUM', 'autosum-sum'),
      menuIconButton(t.autosumAverage, 'autosumFn', 'AVERAGE', 'autosum-average'),
      menuIconButton(t.autosumCount, 'autosumFn', 'COUNT', 'autosum-count'),
      menuIconButton(t.autosumMax, 'autosumFn', 'MAX', 'autosum-max'),
      menuIconButton(t.autosumMin, 'autosumFn', 'MIN', 'autosum-min'),
      menuSeparator(),
      menuIconButton(t.autosumMoreFunctions, 'autosumFn', 'MORE', 'autosum-more'),
    );
    return menu;
  };

  const createCalcOptionsMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-calc-options');
    menu.append(
      calcOptionButton(t.calcAutomatic, 'auto', 'calc-auto'),
      calcOptionButton(t.calcAutoNoTable, 'auto-no-table', 'calc-auto-no-table'),
      calcOptionButton(t.calcManual, 'manual', 'calc-manual'),
      menuSeparator(),
      calcOptionButton(t.calcCalculateNow, 'calculate-now', 'calc-now'),
      calcOptionButton(t.calcCalculateSheet, 'calculate-sheet', 'calc-sheet'),
      menuSeparator(),
      calcOptionButton(t.calcIterative, 'iterative', 'calc-iterative'),
    );
    return menu;
  };

  const createClearArrowsMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-clear-arrows');
    menu.append(
      menuIconButton(t.removeArrowsAll, 'formulaAuditAction', 'clear-all', 'audit-clear-all'),
      menuIconButton(
        t.removePrecedentArrows,
        'formulaAuditAction',
        'clear-precedents',
        'audit-clear-precedents',
      ),
      menuIconButton(
        t.removeDependentArrows,
        'formulaAuditAction',
        'clear-dependents',
        'audit-clear-dependents',
      ),
    );
    return menu;
  };

  const createErrorCheckingMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-error-checking');
    menu.append(
      menuIconButton(t.errorChecking, 'formulaAuditAction', 'error-checking', 'error-checking'),
      menuIconButton(t.traceError, 'formulaAuditAction', 'trace-error', 'trace-error'),
      menuSeparator(),
      menuIconButton(t.ignoreError, 'formulaAuditAction', 'ignore-error', 'ignore-error'),
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
