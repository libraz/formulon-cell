import {
  type CatalogFunctionCategory,
  supportedFunctionNames,
} from '../../../commands/function-categories.js';
import { MAC_AUTOMATION_GALLERY } from './automation-gallery.js';
import {
  EMBEDDED_UNSUPPORTED,
  type Label,
  labels,
  MENU_LABELS,
  NOT_IMPLEMENTED,
  OFFICE_REQUIRED,
} from './labels.js';

/** A leaf rendered in one of the Mac Office-style command menus. */
export interface MacRibbonMenuItem {
  id: string;
  label: string;
  labelJa?: string;
  disabled?: boolean;
  disabledReason?: string;
  icon?: string;
}

/** Root command ids whose primary face opens a Mac menu. The mount can use
 * this explicit list to apply menu-first behavior without changing the
 * generic ribbon activation manifest. */
export const MAC_RIBBON_MENU_COMMANDS = [
  'mac.insert.pivotTable',
  'mac.insert.shapes',
  'mac.insert.recommendedChart',
  'mac.insert.chartColumn',
  'mac.insert.chartBar',
  'mac.insert.chartLine',
  'mac.insert.chartArea',
  'mac.insert.chartPie',
  'mac.insert.chartScatter',
  'mac.draw.eraser',
  'mac.draw.add',
  'mac.page.margins',
  'mac.page.orientation',
  'mac.page.size',
  'mac.page.printArea',
  'mac.page.pageBreaks',
  'mac.formulas.autoSum',
  'mac.formulas.recent',
  'mac.formulas.financial',
  'mac.formulas.logical',
  'mac.formulas.text',
  'mac.formulas.dateTime',
  'mac.formulas.lookup',
  'mac.formulas.math',
  'mac.formulas.more',
  'mac.formulas.calcOptions',
  'mac.formulas.createFromSelection',
  'mac.formulas.removeArrows',
  'mac.formulas.errorCheck',
  'mac.data.dataTypes',
  'mac.data.sortFilter',
  'mac.data.validation',
  'mac.data.whatIf',
  'mac.review.notes',
  'mac.view.show',
  'mac.view.freeze',
  'mac.automate.gallery',
] as const;

export const MAC_RIBBON_MENU_COMMAND_SET: ReadonlySet<string> = new Set(MAC_RIBBON_MENU_COMMANDS);

const menu = (items: readonly MacRibbonMenuItem[]): readonly MacRibbonMenuItem[] =>
  items.map((item) => {
    const label = MENU_LABELS[item.id] ?? labels[item.id.split('.').at(-1) ?? ''];
    return label ? { ...item, label: label.en, labelJa: label.ja } : item;
  });
const leaf = (
  id: string,
  key: string,
  disabled = false,
  disabledReason?: Label,
): MacRibbonMenuItem => ({
  id,
  label: labels[key]?.en ?? key,
  labelJa: labels[key]?.ja,
  disabled,
  disabledReason: disabledReason?.en,
});

const functionMenu = (category: CatalogFunctionCategory): readonly MacRibbonMenuItem[] =>
  menu(
    supportedFunctionNames(category).map((name) => ({
      id: `mac.function.${name}`,
      label: name,
    })),
  );

/** Menu leaves are data commands so the host's normal ribbon click delegation
 *  and restricted-interaction guard can handle them exactly like top-level
 *  buttons. */
export const MAC_RIBBON_MENU_ITEMS: Readonly<Record<string, readonly MacRibbonMenuItem[]>> = {
  'mac.insert.pivotTable': menu([
    leaf('mac.insert.pivotTable', 'pivotTable'),
    leaf('mac.insert.recommendedPivotTable', 'recommendedPivotTable'),
  ]),
  'mac.insert.shapes': menu([
    ...['Line', 'Arrow', 'Rectangle', 'RoundedRectangle', 'Oval', 'Triangle', 'Diamond'].map(
      (kind) => ({
        ...leaf(`mac.insert.shape${kind}`, `shape${kind}`),
        icon: `shape${kind}`,
      }),
    ),
    { ...leaf('mac.insert.shapeCallout', 'shapeCallout', true, NOT_IMPLEMENTED), icon: 'shapes' },
  ]),
  'mac.insert.recommendedChart': menu([
    leaf('mac.insert.chartColumn', 'chartColumn'),
    leaf('mac.insert.chartLine', 'chartLine'),
    leaf('mac.insert.chartPie', 'chartPie'),
    leaf('mac.insert.chartHierarchy', 'chartHierarchy', true, NOT_IMPLEMENTED),
    leaf('mac.insert.chartStatistical', 'chartStatistical', true, NOT_IMPLEMENTED),
    leaf('mac.insert.chartScatter', 'chartScatter'),
    leaf('mac.insert.chartWaterfall', 'chartWaterfall', true, NOT_IMPLEMENTED),
    leaf('mac.insert.chartCombo', 'chartCombo', true, NOT_IMPLEMENTED),
    leaf('mac.insert.chartMap', 'chartMap', true, OFFICE_REQUIRED),
    leaf('mac.insert.pivotChart', 'pivotChart', true, NOT_IMPLEMENTED),
  ]),
  'mac.insert.chartColumn': menu([
    leaf('mac.insert.chartColumn', 'chartColumn'),
    leaf('mac.insert.chartBar', 'chartBar'),
  ]),
  'mac.insert.chartBar': menu([leaf('mac.insert.chartBar', 'chartBar')]),
  'mac.insert.chartLine': menu([
    leaf('mac.insert.chartLine', 'chartLine'),
    leaf('mac.insert.chartArea', 'chartArea'),
  ]),
  'mac.insert.chartArea': menu([leaf('mac.insert.chartArea', 'chartArea')]),
  'mac.insert.chartPie': menu([leaf('mac.insert.chartPie', 'chartPie')]),
  'mac.insert.chartScatter': menu([leaf('mac.insert.chartScatter', 'chartScatter')]),
  'mac.draw.eraser': menu([leaf('mac.draw.eraser', 'eraser'), leaf('mac.draw.toggle', 'draw')]),
  'mac.draw.add': menu([
    leaf('mac.draw.penBlack', 'penBlack'),
    leaf('mac.draw.penRed', 'penRed'),
    leaf('mac.draw.pencil', 'pencil'),
    leaf('mac.draw.highlighter', 'highlighter'),
  ]),
  'mac.page.margins': menu([
    { id: 'mac.page.margins.normal', label: 'Normal' },
    { id: 'mac.page.margins.wide', label: 'Wide' },
    { id: 'mac.page.margins.narrow', label: 'Narrow' },
    { id: 'mac.page.margins.custom', label: 'Custom Margins' },
  ]),
  'mac.page.orientation': menu([
    { id: 'mac.page.orientation.portrait', label: 'Portrait' },
    { id: 'mac.page.orientation.landscape', label: 'Landscape' },
  ]),
  'mac.page.size': menu([
    { id: 'mac.page.size.a4', label: 'A4' },
    { id: 'mac.page.size.a3', label: 'A3' },
    { id: 'mac.page.size.letter', label: 'Letter' },
    { id: 'mac.page.size.legal', label: 'Legal' },
  ]),
  'mac.page.printArea': menu([
    { id: 'mac.page.printArea.set', label: 'Set Print Area' },
    { id: 'mac.page.printArea.add', label: 'Add to Print Area' },
    { id: 'mac.page.printArea.clear', label: 'Clear Print Area' },
  ]),
  'mac.page.pageBreaks': menu([
    { id: 'mac.page.break.insert', label: 'Insert Page Break', labelJa: '改ページを挿入' },
    { id: 'mac.page.break.remove', label: 'Remove Page Break', labelJa: '改ページを削除' },
    {
      id: 'mac.page.break.reset',
      label: 'Reset All Page Breaks',
      labelJa: 'すべての改ページをリセット',
    },
  ]),
  'mac.formulas.autoSum': menu([
    { id: 'mac.autosum.SUM', label: 'Sum', labelJa: '合計' },
    { id: 'mac.autosum.AVERAGE', label: 'Average', labelJa: '平均' },
    { id: 'mac.autosum.COUNT', label: 'Count Numbers', labelJa: '数値の個数' },
    { id: 'mac.autosum.MAX', label: 'Max', labelJa: '最大値' },
    { id: 'mac.autosum.MIN', label: 'Min', labelJa: '最小値' },
    { id: 'mac.formulas.category.all', label: 'More Functions...', labelJa: 'その他の関数...' },
  ]),
  'mac.formulas.recent': menu([
    { id: 'mac.formulas.category.recent', label: 'More Functions...', labelJa: 'その他の関数...' },
  ]),
  'mac.formulas.financial': functionMenu('financial'),
  'mac.formulas.logical': functionMenu('logical'),
  'mac.formulas.text': functionMenu('text'),
  'mac.formulas.dateTime': functionMenu('datetime'),
  'mac.formulas.lookup': functionMenu('lookup'),
  'mac.formulas.math': functionMenu('math'),
  'mac.formulas.more': menu([]),
  'mac.formulas.calcOptions': menu([
    { id: 'mac.formulas.calc.auto', label: 'Automatic', labelJa: '自動' },
    {
      id: 'mac.formulas.calc.autoNoTable',
      label: 'Automatic Except for Data Tables',
      labelJa: 'データ テーブル以外は自動',
    },
    { id: 'mac.formulas.calc.manual', label: 'Manual', labelJa: '手動' },
    {
      id: 'mac.formulas.calc.iterative',
      label: 'Enable Iterative Calculation',
      labelJa: '反復計算を有効にする',
    },
  ]),
  'mac.formulas.createFromSelection': menu([
    { id: 'mac.formulas.createNames.topRow', label: 'Top row' },
    { id: 'mac.formulas.createNames.bottomRow', label: 'Bottom row' },
    { id: 'mac.formulas.createNames.leftColumn', label: 'Left column' },
    { id: 'mac.formulas.createNames.rightColumn', label: 'Right column' },
  ]),
  'mac.data.dataTypes': menu([
    leaf('mac.data.dataTypes.stocks', 'dataTypes', true, OFFICE_REQUIRED),
    leaf('mac.data.dataTypes.geography', 'dataTypes', true, OFFICE_REQUIRED),
  ]),
  'mac.data.sortFilter': menu([
    leaf('mac.data.sortAsc', 'sortAsc'),
    leaf('mac.data.sortDesc', 'sortDesc'),
    leaf('mac.data.sortCustom', 'sortCustom'),
    leaf('mac.data.filter', 'filter'),
    leaf('mac.data.clear', 'clear'),
    leaf('mac.data.reapply', 'reapply'),
    leaf('mac.data.advancedFilter', 'advancedFilter'),
  ]),
  'mac.formulas.removeArrows': menu([
    { id: 'mac.formulas.removeArrows.all', label: 'Remove Arrows' },
    { id: 'mac.formulas.removeArrows.precedents', label: 'Remove Precedent Arrows' },
    { id: 'mac.formulas.removeArrows.dependents', label: 'Remove Dependent Arrows' },
  ]),
  'mac.formulas.errorCheck': menu([
    { id: 'mac.formulas.errorCheck.run', label: 'Error Checking...' },
    { id: 'mac.formulas.errorCheck.trace', label: 'Trace Error' },
    { id: 'mac.formulas.errorCheck.ignore', label: 'Ignore Error' },
  ]),
  'mac.data.validation': menu([
    { id: 'mac.data.validation.settings', label: 'Data Validation...' },
    { id: 'mac.data.validation.circleInvalid', label: 'Circle Invalid Data' },
    { id: 'mac.data.validation.clearCircles', label: 'Clear Validation Circles' },
    { id: 'mac.data.validation.clearRules', label: 'Clear Validation' },
  ]),
  'mac.data.whatIf': menu([
    { id: 'mac.data.goalSeek', label: 'Goal Seek' },
    {
      id: 'mac.data.scenarioManager',
      label: 'Scenario Manager',
      disabled: true,
      disabledReason: NOT_IMPLEMENTED.en,
    },
    {
      id: 'mac.data.dataTable',
      label: 'Data Table',
      disabled: true,
      disabledReason: EMBEDDED_UNSUPPORTED.en,
    },
  ]),
  'mac.review.notes': menu([
    { id: 'mac.review.newNote', label: 'New Note' },
    { id: 'mac.review.nextNote', label: 'Next Note' },
    { id: 'mac.review.previousNote', label: 'Previous Note' },
    { id: 'mac.review.showNotes', label: 'Show Notes' },
  ]),
  'mac.view.show': menu([
    { id: 'mac.view.gridlines', label: 'Gridlines' },
    { id: 'mac.view.headings', label: 'Headings' },
    { id: 'mac.view.formulaBar', label: 'Formula Bar' },
    { id: 'mac.view.zeros', label: 'Zero Values' },
  ]),
  'mac.view.freeze': menu([
    { id: 'mac.view.freeze', label: 'Freeze Panes' },
    { id: 'mac.view.freeze.off', label: 'Unfreeze Panes' },
    { id: 'mac.view.freeze.firstRow', label: 'Freeze Top Row' },
    { id: 'mac.view.freeze.firstColumn', label: 'Freeze First Column' },
  ]),
  'mac.automate.gallery': menu(
    MAC_AUTOMATION_GALLERY.map((script) => ({ id: script.commandId, label: script.label })),
  ),
};

/** DOM id of the Mac Formulas "More Functions" menu. */
export const MAC_FORMULAS_MORE_MENU_ID = 'menu-mac-formulas-more';

export const macRibbonMenuIdForCommand = (commandId: string): string | null => {
  if (!MAC_RIBBON_MENU_COMMAND_SET.has(commandId)) return null;
  const ids: Readonly<Record<string, string>> = {
    'mac.insert.pivotTable': 'menu-mac-insert-pivot',
    'mac.insert.shapes': 'menu-mac-insert-shapes',
    'mac.insert.recommendedChart': 'menu-mac-insert-chart',
    'mac.insert.chartColumn': 'menu-mac-insert-chart-column',
    'mac.insert.chartBar': 'menu-mac-insert-chart-bar',
    'mac.insert.chartLine': 'menu-mac-insert-chart-line',
    'mac.insert.chartArea': 'menu-mac-insert-chart-area',
    'mac.insert.chartPie': 'menu-mac-insert-chart-pie',
    'mac.insert.chartScatter': 'menu-mac-insert-chart-scatter',
    'mac.draw.eraser': 'menu-mac-draw-eraser',
    'mac.draw.add': 'menu-mac-draw-add',
    'mac.page.margins': 'menu-mac-page-margins',
    'mac.page.orientation': 'menu-mac-page-orientation',
    'mac.page.size': 'menu-mac-page-size',
    'mac.page.printArea': 'menu-mac-page-print-area',
    'mac.page.pageBreaks': 'menu-mac-page-breaks',
    'mac.formulas.autoSum': 'menu-mac-formulas-autosum',
    'mac.formulas.recent': 'menu-mac-formulas-recent',
    'mac.formulas.financial': 'menu-mac-formulas-financial',
    'mac.formulas.logical': 'menu-mac-formulas-logical',
    'mac.formulas.text': 'menu-mac-formulas-text',
    'mac.formulas.dateTime': 'menu-mac-formulas-date-time',
    'mac.formulas.lookup': 'menu-mac-formulas-lookup',
    'mac.formulas.math': 'menu-mac-formulas-math',
    'mac.formulas.more': MAC_FORMULAS_MORE_MENU_ID,
    'mac.formulas.calcOptions': 'menu-mac-formulas-calc-options',
    'mac.formulas.createFromSelection': 'menu-mac-formulas-create-names',
    'mac.formulas.removeArrows': 'menu-mac-formulas-remove-arrows',
    'mac.formulas.errorCheck': 'menu-mac-formulas-error-check',
    'mac.data.dataTypes': 'menu-mac-data-types',
    'mac.data.sortFilter': 'menu-mac-data-sort-filter',
    'mac.data.validation': 'menu-mac-data-validation',
    'mac.data.whatIf': 'menu-mac-data-what-if',
    'mac.review.notes': 'menu-mac-review-notes',
    'mac.view.show': 'menu-mac-view-show',
    'mac.view.freeze': 'menu-mac-view-freeze',
    'mac.automate.gallery': 'menu-mac-automate-gallery',
  };
  return ids[commandId] ?? null;
};
