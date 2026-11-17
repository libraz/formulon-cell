import { macFunctionCommandId } from '../../../commands/built-in-command-policy.js';
import { COMPATIBILITY_VERSIONS } from '../../../commands/compatibility-version.js';
import {
  type CatalogFunctionCategory,
  supportedFunctionNames,
} from '../../../commands/function-categories.js';
import type { MacRibbonReason } from './labels-data.js';
import { MAC_SHAPE_GALLERY } from './shape-gallery.js';

/** A leaf rendered in one of the Mac Office-style command menus. Its text comes from its id. */
export interface MacRibbonMenuItem {
  id: string;
  disabled?: boolean;
  disabledReason?: MacRibbonReason;
  icon?: string;
  /** A section caption rather than a command. */
  heading?: boolean;
  /** The label is the gallery's own wording, not the command's ribbon label. */
  own?: boolean;
  /** The row opens a side panel of these entries (the shared `data-format-submenu` markup). The
   *  row itself is not a command. */
  submenu?: readonly MacRibbonMenuItem[];
  /** A checkable row: `menuitemcheckbox` or `menuitemradio`, its state in `aria-checked`. */
  check?: 'checkbox' | 'radio';
  /** A divider (the shared `role="separator"` markup), not a command. */
  separator?: boolean;
}

/** The command rows of a menu: a submenu row contributes its entries, not itself, and a
 *  separator contributes nothing. */
export const macMenuLeaves = (items: readonly MacRibbonMenuItem[]): MacRibbonMenuItem[] =>
  items.flatMap((item) =>
    item.submenu ? macMenuLeaves(item.submenu) : item.separator ? [] : [item],
  );

const separator: MacRibbonMenuItem = { id: 'separator', separator: true };

/** DOM id of the Mac Formulas "More Functions" menu. */
export const MAC_FORMULAS_MORE_MENU_ID = 'menu-mac-formulas-more';

/** DOM id of the menu each root command opens. Its keys are the root command ids whose primary
 * face opens a Mac menu, which the mount uses to apply menu-first behavior without changing the
 * generic ribbon activation manifest. */
const MAC_RIBBON_MENU_DOM_IDS: Readonly<Record<string, string>> = {
  'mac.insert.forms': 'menu-mac-insert-forms',
  'mac.insert.pictureFromData': 'menu-mac-insert-picture-from-data',
  'mac.insert.photo': 'menu-mac-insert-photo',
  'mac.insert.shapes': 'menu-mac-insert-shapes',
  'mac.insert.threeDModel': 'menu-mac-insert-3d-models',
  'mac.insert.smartArt': 'menu-mac-insert-smartart',
  'mac.insert.screenshot': 'menu-mac-insert-screenshot',
  'mac.insert.recommendedChart': 'menu-mac-insert-chart',
  'mac.insert.chartColumn': 'menu-mac-insert-chart-column',
  'mac.insert.chartBar': 'menu-mac-insert-chart-bar',
  'mac.insert.chartLine': 'menu-mac-insert-chart-line',
  'mac.insert.chartArea': 'menu-mac-insert-chart-area',
  'mac.insert.chartPie': 'menu-mac-insert-chart-pie',
  'mac.insert.chartScatter': 'menu-mac-insert-chart-scatter',
  'mac.insert.chartHierarchy': 'menu-mac-insert-chart-hierarchy',
  'mac.insert.chartStatistical': 'menu-mac-insert-chart-statistical',
  'mac.insert.chartWaterfall': 'menu-mac-insert-chart-waterfall',
  'mac.insert.chartCombo': 'menu-mac-insert-chart-combo',
  'mac.insert.chartMap': 'menu-mac-insert-chart-map',
  'mac.insert.sparkline': 'menu-mac-insert-sparkline',
  'mac.insert.textBox': 'menu-mac-insert-text-box',
  'mac.insert.wordArt': 'menu-mac-insert-wordart',
  'mac.insert.equation': 'menu-mac-insert-equation',
  'mac.draw.eraser': 'menu-mac-draw-eraser',
  'mac.draw.add': 'menu-mac-draw-add',
  'mac.page.theme': 'menu-mac-page-theme',
  'mac.page.color': 'menu-mac-page-color',
  'mac.page.font': 'menu-mac-page-font',
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
  'mac.formulas.reset': 'menu-mac-formulas-reset',
  'mac.formulas.defineName': 'menu-mac-formulas-define-name',
  'mac.formulas.useInFormula': 'menu-mac-formulas-use-in-formula',
  'mac.formulas.calcOptions': 'menu-mac-formulas-calc-options',
  'mac.formulas.removeArrows': 'menu-mac-formulas-remove-arrows',
  'mac.formulas.errorCheck': 'menu-mac-formulas-error-check',
  'mac.data.dataTypes': 'menu-mac-data-types',
  'mac.data.pictureFromData': 'menu-mac-data-picture',
  'mac.data.sortFilter': 'menu-mac-data-sort-filter',
  'mac.data.validation': 'menu-mac-data-validation',
  'mac.data.whatIf': 'menu-mac-data-what-if',
  'mac.review.accessibility': 'menu-mac-review-accessibility',
  'mac.review.hideInk': 'menu-mac-review-hide-ink',
  'mac.review.notes': 'menu-mac-review-notes',
  'mac.view.cellFocus': 'menu-mac-view-cell-focus',
  'mac.view.freeze': 'menu-mac-view-freeze',
  'mac.automate.newScript': 'menu-mac-automate-new-script',
  'mac.automate.showScripts': 'menu-mac-automate-view-scripts',
};

export const MAC_RIBBON_MENU_COMMANDS: readonly string[] = Object.keys(MAC_RIBBON_MENU_DOM_IDS);

export const MAC_RIBBON_MENU_COMMAND_SET: ReadonlySet<string> = new Set(MAC_RIBBON_MENU_COMMANDS);

/** Disable a row and, recursively, every entry of its submenu. */
const greyOutItem = (item: MacRibbonMenuItem, reason: MacRibbonReason): MacRibbonMenuItem => ({
  ...item,
  disabled: true,
  disabledReason: reason,
  ...(item.submenu ? { submenu: item.submenu.map((entry) => greyOutItem(entry, reason)) } : {}),
});

/** A menu whose entries the host cannot run: every row greyed out with the same reason. */
const unavailable = (
  items: readonly MacRibbonMenuItem[],
  reason: MacRibbonReason,
): readonly MacRibbonMenuItem[] => items.map((item) => greyOutItem(item, reason));

/** A gallery section caption, worded by its generated `gallery:` label. */
const section = (id: string): MacRibbonMenuItem => ({
  id: `mac.insert.heading.${id}`,
  heading: true,
  own: true,
});

/** A gallery entry; `enabled` entries run their chart command, the rest stay greyed out. */
const galleryItem = (id: string, enabled = false): MacRibbonMenuItem => ({
  id,
  own: true,
  disabled: !enabled,
  disabledReason: enabled ? undefined : 'notImplemented',
});

const chartGallery = {
  column: [
    section('column2d'),
    galleryItem('mac.insert.chartColumn', true),
    galleryItem('mac.insert.chartColumnStacked', true),
    galleryItem('mac.insert.chartColumn100', true),
    section('column3d'),
    galleryItem('mac.insert.chartColumn3d'),
    galleryItem('mac.insert.chartColumnStacked3d'),
    galleryItem('mac.insert.chartColumn1003d'),
    galleryItem('mac.insert.chartColumnTrue3d'),
    section('bar2d'),
    galleryItem('mac.insert.chartBar', true),
    galleryItem('mac.insert.chartBarStacked', true),
    galleryItem('mac.insert.chartBar100', true),
    section('bar3d'),
    galleryItem('mac.insert.chartBar3d'),
    galleryItem('mac.insert.chartBarStacked3d'),
    galleryItem('mac.insert.chartBar1003d'),
  ],
  line: [
    section('line2d'),
    galleryItem('mac.insert.chartLine', true),
    galleryItem('mac.insert.chartLineStacked', true),
    galleryItem('mac.insert.chartLine100', true),
    galleryItem('mac.insert.chartLineMarkers'),
    galleryItem('mac.insert.chartLineStackedMarkers'),
    galleryItem('mac.insert.chartLine100Markers'),
    section('line3d'),
    galleryItem('mac.insert.chartLine3d'),
    section('area2d'),
    galleryItem('mac.insert.chartArea', true),
    galleryItem('mac.insert.chartAreaStacked', true),
    galleryItem('mac.insert.chartArea100', true),
    section('area3d'),
    galleryItem('mac.insert.chartArea3d'),
    galleryItem('mac.insert.chartAreaStacked3d'),
    galleryItem('mac.insert.chartArea1003d'),
  ],
  pie: [
    section('pie2d'),
    galleryItem('mac.insert.chartPie', true),
    galleryItem('mac.insert.chartPieOfPie'),
    galleryItem('mac.insert.chartBarOfPie'),
    section('pie3d'),
    galleryItem('mac.insert.chartPie3d'),
    section('doughnut'),
    galleryItem('mac.insert.chartDoughnut', true),
  ],
  scatter: [
    section('scatter'),
    galleryItem('mac.insert.chartScatter', true),
    galleryItem('mac.insert.chartScatterSmooth'),
    galleryItem('mac.insert.chartScatterLinesMarkers'),
    galleryItem('mac.insert.chartScatterLines'),
    section('bubble'),
    galleryItem('mac.insert.chartBubble'),
    galleryItem('mac.insert.chartBubble3d'),
  ],
  hierarchy: [
    section('treemap'),
    galleryItem('mac.insert.chartTreemap'),
    section('sunburst'),
    galleryItem('mac.insert.chartSunburst'),
  ],
  statistical: [
    section('histogramGroup'),
    galleryItem('mac.insert.chartHistogram'),
    galleryItem('mac.insert.chartPareto'),
    section('boxWhisker'),
    galleryItem('mac.insert.chartBoxWhisker'),
  ],
  waterfall: [
    section('waterfall'),
    galleryItem('mac.insert.chartWaterfall.plot'),
    section('funnel'),
    galleryItem('mac.insert.chartFunnel'),
    section('stock'),
    galleryItem('mac.insert.chartStockHlc'),
    galleryItem('mac.insert.chartStockOhlc'),
    galleryItem('mac.insert.chartStockVhlc'),
    galleryItem('mac.insert.chartStockVohlc'),
    section('surface'),
    galleryItem('mac.insert.chartSurface'),
    galleryItem('mac.insert.chartSurfaceWireframe'),
    galleryItem('mac.insert.chartContour'),
    galleryItem('mac.insert.chartContourWireframe'),
    section('radar'),
    galleryItem('mac.insert.chartRadar'),
    galleryItem('mac.insert.chartRadarMarkers'),
    galleryItem('mac.insert.chartRadarFilled'),
  ],
  combo: [
    section('combo'),
    galleryItem('mac.insert.chartComboColumnLine'),
    galleryItem('mac.insert.chartComboColumnLineSecondary'),
    galleryItem('mac.insert.chartComboAreaColumn'),
  ],
  map: [section('filledMap'), galleryItem('mac.insert.chartFilledMap')],
} as const satisfies Record<string, readonly MacRibbonMenuItem[]>;

const functionMenu = (category: CatalogFunctionCategory): readonly MacRibbonMenuItem[] =>
  supportedFunctionNames(category).map((name) => ({ id: macFunctionCommandId(name) }));

/** Menu leaves are data commands so the host's normal ribbon click delegation
 *  and restricted-interaction guard can handle them exactly like top-level
 *  buttons. */
export const MAC_RIBBON_MENU_ITEMS: Readonly<Record<string, readonly MacRibbonMenuItem[]>> = {
  // Forms is a cloud feature the host does not provide, so its menu stays greyed out.
  'mac.insert.forms': unavailable(
    [
      { id: 'mac.insert.forms.new' },
      { id: 'mac.insert.forms.preview' },
      { id: 'mac.insert.forms.edit' },
      { id: 'mac.insert.forms.send' },
    ],
    'officeRequired',
  ),
  'mac.insert.pictureFromData': unavailable(
    [{ id: 'mac.insert.pictureFromData.file' }, { id: 'mac.insert.pictureFromData.clipboard' }],
    'officeRequired',
  ),
  // Photo Browser, Stock Images and Online Pictures need the cloud or the OS photo library.
  'mac.insert.photo': [
    { id: 'mac.insert.photo.browser' },
    { id: 'mac.insert.photo.file' },
    { id: 'mac.insert.photo.stock' },
    { id: 'mac.insert.photo.online' },
  ].map((item) =>
    item.id === 'mac.insert.photo.file' ? item : greyOutItem(item, 'notImplemented'),
  ),
  'mac.insert.shapes': MAC_SHAPE_GALLERY,
  'mac.insert.threeDModel': unavailable(
    [{ id: 'mac.insert.threeDModel.device' }, { id: 'mac.insert.threeDModel.stock' }],
    'notImplemented',
  ),
  'mac.insert.smartArt': unavailable(
    [
      { id: 'mac.insert.smartArt.list' },
      { id: 'mac.insert.smartArt.process' },
      { id: 'mac.insert.smartArt.cycle' },
      { id: 'mac.insert.smartArt.hierarchy' },
      { id: 'mac.insert.smartArt.relationship' },
      { id: 'mac.insert.smartArt.matrix' },
      { id: 'mac.insert.smartArt.pyramid' },
      { id: 'mac.insert.smartArt.picture' },
    ],
    'notImplemented',
  ),
  'mac.insert.screenshot': [{ id: 'mac.insert.screenshot.clipping' }],
  // With no data selected the menu carries only the explanation; otherwise the recommendation.
  'mac.insert.recommendedChart': [
    {
      id: 'mac.insert.recommendedChart.none',
      own: true,
      disabled: true,
    },
    {
      id: 'mac.insert.recommendedChart.apply',
      own: true,
    },
  ],
  'mac.insert.chartColumn': chartGallery.column,
  'mac.insert.chartBar': [{ id: 'mac.insert.chartBar' }],
  'mac.insert.chartLine': chartGallery.line,
  'mac.insert.chartArea': [{ id: 'mac.insert.chartArea' }],
  'mac.insert.chartPie': chartGallery.pie,
  'mac.insert.chartScatter': chartGallery.scatter,
  'mac.insert.chartHierarchy': chartGallery.hierarchy,
  'mac.insert.chartStatistical': chartGallery.statistical,
  'mac.insert.chartWaterfall': chartGallery.waterfall,
  'mac.insert.chartCombo': chartGallery.combo,
  'mac.insert.chartMap': chartGallery.map,
  // Each type opens the sparkline dialog with that type chosen.
  'mac.insert.sparkline': [
    { id: 'mac.insert.sparkline.line' },
    { id: 'mac.insert.sparkline.column' },
    { id: 'mac.insert.sparkline.winLoss' },
  ],
  // The file the engine writes holds no vertical text.
  'mac.insert.textBox': [
    { id: 'mac.insert.textBox.horizontal' },
    { id: 'mac.insert.textBox.vertical' },
  ].map((item) =>
    item.id === 'mac.insert.textBox.vertical' ? greyOutItem(item, 'embeddedUnsupported') : item,
  ),
  // The WordArt gallery carries styles, not named entries.
  'mac.insert.wordArt': [],
  'mac.insert.equation': unavailable([{ id: 'mac.insert.equation.insert' }], 'notImplemented'),
  // Only the whole-stroke eraser exists; point and segment erasing are not built.
  'mac.draw.eraser': [
    { id: 'mac.draw.eraser.stroke' },
    {
      id: 'mac.draw.eraser.small',
      disabled: true,
      disabledReason: 'notImplemented',
    },
    {
      id: 'mac.draw.eraser.medium',
      disabled: true,
      disabledReason: 'notImplemented',
    },
    {
      id: 'mac.draw.eraser.segment',
      disabled: true,
      disabledReason: 'notImplemented',
    },
  ],
  // Adding a pen kind to the strip is not built; the four pens have fixed buttons.
  'mac.draw.add': unavailable(
    [{ id: 'mac.draw.add.pen' }, { id: 'mac.draw.add.highlighter' }, { id: 'mac.draw.add.pencil' }],
    'notImplemented',
  ),
  // Office and Grayscale are the themes the host can apply; the commands below them
  // are the desktop's online-gallery, browse and save-to-file entries, which it has no backing for.
  'mac.page.theme': [
    { id: 'mac.page.theme.office' },
    { id: 'mac.page.theme.gray' },
    {
      id: 'mac.page.theme.enableUpdates',
      disabled: true,
      disabledReason: 'officeRequired',
    },
    {
      id: 'mac.page.theme.browse',
      disabled: true,
      disabledReason: 'notImplemented',
    },
    {
      id: 'mac.page.theme.save',
      disabled: true,
      disabledReason: 'notImplemented',
    },
  ],
  'mac.page.color': [{ id: 'mac.page.color.office' }, { id: 'mac.page.color.gray' }],
  'mac.page.font': [{ id: 'mac.page.font.office' }, { id: 'mac.page.font.gray' }],
  'mac.page.margins': [
    { id: 'mac.page.margins.normal' },
    { id: 'mac.page.margins.wide' },
    { id: 'mac.page.margins.narrow' },
    { id: 'mac.page.margins.custom' },
  ],
  'mac.page.orientation': [
    { id: 'mac.page.orientation.portrait' },
    { id: 'mac.page.orientation.landscape' },
  ],
  'mac.page.size': [
    { id: 'mac.page.size.a4' },
    { id: 'mac.page.size.a3' },
    { id: 'mac.page.size.letter' },
    { id: 'mac.page.size.legal' },
  ],
  'mac.page.printArea': [
    { id: 'mac.page.printArea.set' },
    { id: 'mac.page.printArea.add' },
    { id: 'mac.page.printArea.clear' },
  ],
  'mac.page.pageBreaks': [
    { id: 'mac.page.break.insert' },
    { id: 'mac.page.break.remove' },
    {
      id: 'mac.page.break.reset',
    },
  ],
  'mac.formulas.autoSum': [
    { id: 'mac.autosum.SUM' },
    { id: 'mac.autosum.AVERAGE' },
    { id: 'mac.autosum.COUNT' },
    { id: 'mac.autosum.MAX' },
    { id: 'mac.autosum.MIN' },
    { id: 'mac.formulas.category.all' },
  ],
  'mac.formulas.recent': [{ id: 'mac.formulas.category.recent' }],
  'mac.formulas.financial': functionMenu('financial'),
  'mac.formulas.logical': functionMenu('logical'),
  'mac.formulas.text': functionMenu('text'),
  'mac.formulas.dateTime': functionMenu('datetime'),
  'mac.formulas.lookup': functionMenu('lookup'),
  'mac.formulas.math': functionMenu('math'),
  'mac.formulas.more': [],
  // Python is a cloud feature the host does not provide, so its menu stays greyed out.
  'mac.formulas.reset': unavailable(
    [
      { id: 'mac.formulas.reset.clear' },
      { id: 'mac.formulas.reset.runtime' },
      separator,
      {
        id: 'mac.formulas.reset.version',
        submenu: [
          { id: 'mac.formulas.reset.versionStable' },
          { id: 'mac.formulas.reset.versionLatest' },
          separator,
          { id: 'mac.formulas.reset.versionLearnMore' },
        ],
      },
    ],
    'officeRequired',
  ),
  'mac.formulas.defineName': [
    { id: 'mac.formulas.defineName.define' },
    {
      id: 'mac.formulas.defineName.apply',
      disabled: true,
      disabledReason: 'notImplemented',
    },
  ],
  'mac.formulas.useInFormula': [{ id: 'mac.formulas.useInFormula.paste' }],
  'mac.formulas.calcOptions': [
    { id: 'mac.formulas.calc.auto' },
    {
      id: 'mac.formulas.calc.autoNoTable',
    },
    { id: 'mac.formulas.calc.manual' },
    separator,
    { id: 'mac.formulas.calc.formatStale', check: 'checkbox' },
    separator,
    {
      id: 'mac.formulas.calc.compatVersion',
      submenu: [
        ...COMPATIBILITY_VERSIONS.map(
          (version): MacRibbonMenuItem => ({
            id: `mac.formulas.calc.compatVersion.${version}`,
            check: 'radio',
          }),
        ),
        separator,
        {
          id: 'mac.formulas.calc.compatVersion.learnMore',
          disabled: true,
          disabledReason: 'notImplemented',
        },
      ],
    },
  ],
  // The host has no linked data types, so the gallery lists Excel's entries greyed out.
  'mac.data.dataTypes': [
    galleryItem('mac.data.dataTypes.stocks'),
    galleryItem('mac.data.dataTypes.currencies'),
    galleryItem('mac.data.dataTypes.geography'),
  ],
  'mac.data.pictureFromData': [
    galleryItem('mac.data.pictureFromData.file'),
    galleryItem('mac.data.pictureFromData.clipboard'),
  ],
  'mac.data.sortFilter': [
    { id: 'mac.data.sortAsc' },
    { id: 'mac.data.sortDesc' },
    { id: 'mac.data.sortCustom' },
    { id: 'mac.data.filter' },
    { id: 'mac.data.clear' },
    { id: 'mac.data.reapply' },
    { id: 'mac.data.advancedFilter' },
  ],
  'mac.formulas.removeArrows': [
    { id: 'mac.formulas.removeArrows.all' },
    { id: 'mac.formulas.removeArrows.precedents' },
    { id: 'mac.formulas.removeArrows.dependents' },
  ],
  'mac.formulas.errorCheck': [
    { id: 'mac.formulas.errorCheck.run' },
    { id: 'mac.formulas.errorCheck.trace' },
    // Its entries are the active sheet's cycle members, listed when the menu opens.
    { id: 'mac.formulas.errorCheck.circular', submenu: [] },
  ],
  'mac.data.validation': [
    { id: 'mac.data.validation.settings' },
    { id: 'mac.data.validation.circleInvalid' },
    { id: 'mac.data.validation.clearCircles' },
    { id: 'mac.data.validation.clearRules' },
  ],
  'mac.data.whatIf': [
    {
      id: 'mac.data.scenarioManager',
      disabled: true,
      disabledReason: 'notImplemented',
    },
    { id: 'mac.data.goalSeek' },
    { id: 'mac.data.dataTable' },
  ],
  // Alt Text and Accessibility Help have nothing behind them.
  'mac.review.accessibility': [
    { id: 'mac.review.accessibility.check' },
    ...unavailable([{ id: 'mac.review.accessibility.altText' }], 'notImplemented'),
    { id: 'mac.review.accessibility.formatTable' },
    { id: 'mac.review.accessibility.unmerge' },
    ...unavailable([{ id: 'mac.review.accessibility.help' }], 'notImplemented'),
  ],
  // Ink cannot be hidden or deleted in this host, so every entry is greyed out.
  'mac.review.hideInk': [
    ...unavailable([{ id: 'mac.review.hideInk.hide' }], 'embeddedUnsupported'),
    ...unavailable(
      [{ id: 'mac.review.hideInk.deleteSheet' }, { id: 'mac.review.hideInk.deleteWorkbook' }],
      'notImplemented',
    ),
  ],
  'mac.review.notes': [
    { id: 'mac.review.newNote' },
    { id: 'mac.review.previousNote' },
    { id: 'mac.review.nextNote' },
    { id: 'mac.review.showHideNote' },
    { id: 'mac.review.showNotes' },
    { id: 'mac.review.convertNotes' },
  ],
  // Only the measured default colour is built; Focus Cell Color stays greyed out.
  'mac.view.cellFocus': [
    { id: 'mac.view.cellFocus.focusCell', check: 'checkbox' },
    ...unavailable([{ id: 'mac.view.cellFocus.color' }], 'notImplemented'),
    { id: 'mac.view.cellFocus.autoHighlight', check: 'checkbox' },
  ],
  'mac.view.freeze': [
    { id: 'mac.view.freeze.panes' },
    { id: 'mac.view.freeze.firstRow' },
    { id: 'mac.view.freeze.firstColumn' },
  ],
  // Office Scripts is a cloud feature the host does not provide, so New Script's menu stays greyed out.
  'mac.automate.newScript': unavailable(
    [{ id: 'mac.automate.newScript.editor' }, { id: 'mac.automate.newScript.recording' }],
    'officeRequired',
  ),
  // Samples opens the script gallery; the other two lists have no script storage behind them.
  'mac.automate.showScripts': [
    ...unavailable(
      [{ id: 'mac.automate.showScripts.recent' }, { id: 'mac.automate.showScripts.workbook' }],
      'notImplemented',
    ),
    { id: 'mac.automate.showScripts.samples' },
  ],
};

export const macRibbonMenuIdForCommand = (commandId: string): string | null => {
  if (!MAC_RIBBON_MENU_COMMAND_SET.has(commandId)) return null;
  return MAC_RIBBON_MENU_DOM_IDS[commandId] ?? null;
};
