import { reportDialogLabels, showReport } from '../toolbar/dialogs/report.js';
import { analyzeAccessibilityCells, reviewCellsFromState } from '../toolbar/review-tools.js';
import { toolbarLangForLocale } from '../toolbar/ribbon/mac/locale.js';
import type { SpreadsheetInstance } from './types.js';

/** The accessibility check where the host has no docked checker: a report on the selection. */
export async function showAccessibilityReport(instance: SpreadsheetInstance): Promise<void> {
  const state = instance.store.getState();
  const cells = reviewCellsFromState(state, state.data.sheetIndex, state.selection.range);
  const strings = instance.i18n.strings;
  await showReport({
    title: strings.ribbon.accessibility,
    items: analyzeAccessibilityCells(cells, toolbarLangForLocale(instance.i18n.locale)),
    ...reportDialogLabels(strings),
  });
  instance.host.focus();
}
