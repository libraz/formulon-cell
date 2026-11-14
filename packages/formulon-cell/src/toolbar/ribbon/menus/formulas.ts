// Formulas tab menus: AutoSum, calc options, formula audit error checks. The
// factories use shared icon menu rows; calc options uses a mix of
// menuitem/menuitemradio entries so the parent can toggle the radio state after
// rendering.

import {
  COMPATIBILITY_VERSIONS,
  compatibilityVersionLabel,
} from '../../../commands/compatibility-version.js';
import type { ToolbarMenuText } from '../../menu-text.js';
import type { ToolbarLang } from '../../ribbon-model.js';

import {
  createMenu,
  createSubmenu,
  type MenuIconSlug,
  menuIconButton,
  menuIconSpacer,
  menuIdForCommand,
  menuPresetButton,
  menuSeparator,
  menuSubmenuTrigger,
} from './general.js';

export type AutoSumFormulaName = 'SUM' | 'AVERAGE' | 'COUNT' | 'MAX' | 'MIN' | 'MORE';

export interface FormulasMenuFactories {
  createAutoSumMenu: (id: string) => HTMLDivElement;
  createCalcOptionsMenu: () => HTMLDivElement;
  createClearArrowsMenu: () => HTMLDivElement;
  createErrorCheckingMenu: () => HTMLDivElement;
  createWhatIfMenu: () => HTMLDivElement;
}

type FormulasMenuText = ToolbarMenuText & {
  calcAutomatic: string;
  calcAutoNoTable: string;
  calcManual: string;
  calcCalculateNow: string;
  calcCalculateSheet: string;
  calcIterative: string;
};

const calcOptionButton = (label: string, value: string, icon: MenuIconSlug): HTMLButtonElement => {
  const button = menuIconButton(label, 'calcOption', value, icon);
  if (value === 'auto' || value === 'manual' || value === 'auto-no-table') {
    button.setAttribute('role', 'menuitemradio');
    button.setAttribute('aria-checked', 'false');
  }
  return button;
};

/** A checkable entry without an icon: `menuitemcheckbox` or `menuitemradio`, unchecked. */
const checkableOption = (
  label: string,
  attr: string,
  value: string,
  role: 'menuitemcheckbox' | 'menuitemradio',
): HTMLButtonElement => {
  const button = menuPresetButton(label, attr, value, menuIconSpacer());
  button.setAttribute('role', role);
  button.setAttribute('aria-checked', 'false');
  return button;
};

/** A submenu row and its panel, in the shared format-submenu markup. */
const formatSubmenu = (
  menuId: string,
  key: string,
  label: string,
  children: readonly Node[],
): [HTMLButtonElement, HTMLDivElement] => {
  const id = `${menuId}-${key}`;
  const trigger = menuSubmenuTrigger(
    menuPresetButton(label, 'formatSubmenu', key, menuIconSpacer()),
    undefined,
    { controlsId: id },
  );
  const panel = createSubmenu({
    id,
    className: 'fc-tb__submenu fc-tb__submenu--format',
    label,
    dataset: { formatPanel: key },
  });
  panel.append(...children);
  return [trigger, panel];
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
      menuPresetButton(t.autosumAverage, 'autosumFn', 'AVERAGE', menuIconSpacer()),
      menuPresetButton(t.autosumCount, 'autosumFn', 'COUNT', menuIconSpacer()),
      menuPresetButton(t.autosumMax, 'autosumFn', 'MAX', menuIconSpacer()),
      menuPresetButton(t.autosumMin, 'autosumFn', 'MIN', menuIconSpacer()),
      menuSeparator(),
      menuPresetButton(t.autosumMoreFunctions, 'autosumFn', 'MORE', menuIconSpacer()),
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
      checkableOption(t.formatStaleValues, 'calcOption', 'format-stale', 'menuitemcheckbox'),
      menuSeparator(),
      ...formatSubmenu('menu-calc-options', 'compat-version', t.compatibilityVersion, [
        ...COMPATIBILITY_VERSIONS.map((version) =>
          checkableOption(
            compatibilityVersionLabel(t, version),
            'calcOption',
            `compat-${version}`,
            'menuitemradio',
          ),
        ),
        menuSeparator(),
        menuPresetButton(
          t.compatibilityVersionLearnMore,
          'calcOption',
          'compat-learn-more',
          menuIconSpacer(),
        ),
      ]),
      menuSeparator(),
      calcOptionButton(t.calcCalculateNow, 'calculate-now', 'calc-now'),
      calcOptionButton(t.calcCalculateSheet, 'calculate-sheet', 'calc-sheet'),
      menuSeparator(),
      calcOptionButton(t.calcIterative, 'iterative', 'calc-iterative'),
    );
    return menu;
  };

  // Data > Forecast > What-If Analysis.
  const createWhatIfMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-what-if');
    menu.append(
      menuIconButton(t.whatIfGoalSeek, 'whatIf', 'goal-seek', 'what-if-goal-seek'),
      menuIconButton(t.whatIfDataTable, 'whatIf', 'data-table', 'what-if-data-table'),
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
      // Its entries are the active sheet's cycle members, listed when the menu opens.
      ...formatSubmenu('menu-error-checking', 'circular', t.circularReferences, []),
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
    createWhatIfMenu,
  };
};
