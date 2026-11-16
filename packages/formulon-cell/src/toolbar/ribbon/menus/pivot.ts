// PivotTable Design and Analyze tab menus. Every item emits a `data-pivot-layout`
// action the dynamic-dropdown dispatcher reads; entries whose feature the host
// or engine lacks stay listed, disabled with a reason, as the desktop menus are.

import { projectDisabledState } from '../../menu-a11y.js';
import type { ToolbarMenuText } from '../../menu-text.js';
import {
  createMenu,
  menuIconSpacer,
  menuIdForCommand,
  menuPresetButton,
  menuSeparator,
} from './general.js';

export interface PivotMenuDeps {
  ribbonMenuText: ToolbarMenuText;
}

export interface PivotMenuFactories {
  /** The menu of one Design / Analyze command; an unknown command gets an empty menu. */
  createPivotLayoutMenu: (commandId: string) => HTMLDivElement;
}

/** Keys of the menu text that are plain labels. */
type LabelKey = {
  [K in keyof ToolbarMenuText]: ToolbarMenuText[K] extends string ? K : never;
}[keyof ToolbarMenuText];

type Entry = readonly [action: string, label: LabelKey, available: boolean] | null;

const MENUS: Readonly<Record<string, readonly Entry[]>> = {
  pivotSubtotals: [
    ['subtotals-none', 'pivotSubtotalsNone', false],
    ['subtotals-bottom', 'pivotSubtotalsBottom', true],
    ['subtotals-top', 'pivotSubtotalsTop', true],
    null,
    ['subtotals-filtered', 'pivotSubtotalsFiltered', false],
  ],
  pivotGrandTotals: [
    ['totals-off', 'pivotTotalsOff', true],
    ['totals-on', 'pivotTotalsOn', true],
    ['totals-rows', 'pivotTotalsRows', true],
    ['totals-cols', 'pivotTotalsCols', true],
  ],
  pivotReportLayout: [
    ['layout-compact', 'pivotLayoutCompact', true],
    ['layout-outline', 'pivotLayoutOutline', true],
    ['layout-tabular', 'pivotLayoutTabular', true],
    null,
    ['labels-repeat', 'pivotLabelsRepeat', false],
    ['labels-no-repeat', 'pivotLabelsNoRepeat', false],
  ],
  pivotBlankRows: [
    ['blank-insert', 'pivotBlankInsert', false],
    ['blank-remove', 'pivotBlankRemove', false],
  ],
  pivotClear: [
    ['clear-all', 'pivotClearAll', true],
    ['clear-filters', 'pivotClearFilters', true],
  ],
  pivotSelect: [
    ['select-labels-values', 'pivotSelectLabelsValues', false],
    ['select-values', 'pivotSelectValues', false],
    ['select-labels', 'pivotSelectLabels', false],
    ['select-table', 'pivotSelectTable', true],
    null,
    ['select-enable', 'pivotSelectEnable', false],
  ],
  pivotFieldsItemsSets: [
    ['calc-field', 'pivotCalcField', false],
    ['calc-item', 'pivotCalcItem', false],
    ['solve-order', 'pivotSolveOrder', false],
    ['list-formulas', 'pivotListFormulas', false],
  ],
};

export const createPivotMenuFactories = (deps: PivotMenuDeps): PivotMenuFactories => {
  const t = deps.ribbonMenuText;
  const createPivotLayoutMenu = (commandId: string): HTMLDivElement => {
    const menu = createMenu(menuIdForCommand(commandId));
    for (const entry of MENUS[commandId] ?? []) {
      if (!entry) {
        menu.appendChild(menuSeparator());
        continue;
      }
      const [action, label, available] = entry;
      const button = menuPresetButton(t[label], 'pivotLayout', action, menuIconSpacer());
      if (!available) {
        projectDisabledState(button, true, t.pivotUnavailable, { datasetKey: 'disabledReason' });
      }
      // Selection stays enabled on the desktop, so its entry reads checked.
      if (action === 'select-enable') {
        button.setAttribute('role', 'menuitemcheckbox');
        button.setAttribute('aria-checked', 'true');
      }
      menu.appendChild(button);
    }
    return menu;
  };
  return { createPivotLayoutMenu };
};
