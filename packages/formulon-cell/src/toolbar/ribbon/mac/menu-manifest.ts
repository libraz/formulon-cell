import { macFunctionCommandId } from '../../../commands/built-in-command-policy.js';
import { COMPATIBILITY_VERSIONS } from '../../../commands/compatibility-version.js';
import {
  type CatalogFunctionCategory,
  supportedFunctionNames,
} from '../../../commands/function-categories.js';
import { macBundleLabel } from './bundle-text.js';
import {
  EMBEDDED_UNSUPPORTED,
  type Label,
  labelFor,
  labels,
  NOT_IMPLEMENTED,
  OFFICE_REQUIRED,
} from './labels.js';
import { MAC_SHAPE_GALLERY } from './shape-gallery.js';

/** A leaf rendered in one of the Mac Office-style command menus. */
export interface MacRibbonMenuItem {
  id: string;
  label: string;
  disabled?: boolean;
  disabledReason?: string;
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

const separator: MacRibbonMenuItem = { id: 'separator', label: '', separator: true };

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

const menu = (items: readonly MacRibbonMenuItem[]): readonly MacRibbonMenuItem[] =>
  items.map((item) => {
    const label = labelFor(item.id);
    const localized = label ? { ...item, label: label.en } : item;
    return item.submenu ? { ...localized, submenu: menu(item.submenu) } : localized;
  });
/** Disable a row and, recursively, every entry of its submenu. */
const greyOutItem = (item: MacRibbonMenuItem, reason: Label): MacRibbonMenuItem => ({
  ...item,
  disabled: true,
  disabledReason: reason.en,
  ...(item.submenu ? { submenu: item.submenu.map((entry) => greyOutItem(entry, reason)) } : {}),
});

/** A menu whose entries the host cannot run: every row greyed out with the same reason. */
const unavailable = (
  items: readonly MacRibbonMenuItem[],
  reason: Label,
): readonly MacRibbonMenuItem[] => menu(items).map((item) => greyOutItem(item, reason));
const leaf = (
  id: string,
  key: string,
  disabled = false,
  disabledReason?: Label,
): MacRibbonMenuItem => ({
  id,
  label: labels[key]?.en ?? key,
  disabled,
  disabledReason: disabledReason?.en,
});

/** English wording of a gallery's own row; every locale's is generated as `gallery:<id>`. */
const galleryLabel = (id: string): string => macBundleLabel(`gallery:${id}`).en;

/** A gallery section caption. */
const section = (id: string): MacRibbonMenuItem => ({
  id: `mac.insert.heading.${id}`,
  label: galleryLabel(`mac.insert.heading.${id}`),
  heading: true,
  own: true,
});

/** A gallery entry; `enabled` entries run their chart command, the rest stay greyed out. */
const galleryItem = (id: string, enabled = false): MacRibbonMenuItem => ({
  id,
  label: galleryLabel(id),
  own: true,
  disabled: !enabled,
  disabledReason: enabled ? undefined : NOT_IMPLEMENTED.en,
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
  menu(
    supportedFunctionNames(category).map((name) => ({
      id: macFunctionCommandId(name),
      label: name,
    })),
  );

/** Menu leaves are data commands so the host's normal ribbon click delegation
 *  and restricted-interaction guard can handle them exactly like top-level
 *  buttons. */
export const MAC_RIBBON_MENU_ITEMS: Readonly<Record<string, readonly MacRibbonMenuItem[]>> = {
  // Forms is a cloud feature the host does not provide, so its menu stays greyed out.
  'mac.insert.forms': unavailable(
    [
      { id: 'mac.insert.forms.new', label: 'New Form' },
      { id: 'mac.insert.forms.preview', label: 'Preview Form' },
      { id: 'mac.insert.forms.edit', label: 'Edit Form' },
      { id: 'mac.insert.forms.send', label: 'Send Form' },
    ],
    OFFICE_REQUIRED,
  ),
  'mac.insert.pictureFromData': unavailable(
    [
      { id: 'mac.insert.pictureFromData.file', label: 'Picture From File...' },
      { id: 'mac.insert.pictureFromData.clipboard', label: 'Picture From Clipboard' },
    ],
    OFFICE_REQUIRED,
  ),
  // Photo Browser, Stock Images and Online Pictures need the cloud or the OS photo library.
  'mac.insert.photo': menu([
    { id: 'mac.insert.photo.browser', label: 'Photo Browser…' },
    { id: 'mac.insert.photo.file', label: 'Picture from File...' },
    { id: 'mac.insert.photo.stock', label: 'Stock Images...' },
    { id: 'mac.insert.photo.online', label: 'Online Pictures...' },
  ]).map((item) =>
    item.id === 'mac.insert.photo.file'
      ? item
      : { ...item, disabled: true, disabledReason: NOT_IMPLEMENTED.en },
  ),
  'mac.insert.shapes': MAC_SHAPE_GALLERY,
  'mac.insert.threeDModel': unavailable(
    [
      { id: 'mac.insert.threeDModel.device', label: 'This Device...' },
      { id: 'mac.insert.threeDModel.stock', label: 'Stock 3D Models...' },
    ],
    NOT_IMPLEMENTED,
  ),
  'mac.insert.smartArt': unavailable(
    [
      { id: 'mac.insert.smartArt.list', label: 'List' },
      { id: 'mac.insert.smartArt.process', label: 'Process' },
      { id: 'mac.insert.smartArt.cycle', label: 'Cycle' },
      { id: 'mac.insert.smartArt.hierarchy', label: 'Hierarchy' },
      { id: 'mac.insert.smartArt.relationship', label: 'Relationship' },
      { id: 'mac.insert.smartArt.matrix', label: 'Matrix' },
      { id: 'mac.insert.smartArt.pyramid', label: 'Pyramid' },
      { id: 'mac.insert.smartArt.picture', label: 'Picture' },
    ],
    NOT_IMPLEMENTED,
  ),
  'mac.insert.screenshot': menu([
    { id: 'mac.insert.screenshot.clipping', label: 'Screen Clipping' },
  ]),
  // With no data selected the menu carries only the explanation; otherwise the recommendation.
  'mac.insert.recommendedChart': [
    {
      id: 'mac.insert.recommendedChart.none',
      label: galleryLabel('mac.insert.recommendedChart.none'),
      own: true,
      disabled: true,
    },
    {
      id: 'mac.insert.recommendedChart.apply',
      label: galleryLabel('mac.insert.recommendedChart.apply'),
      own: true,
    },
  ],
  'mac.insert.chartColumn': chartGallery.column,
  'mac.insert.chartBar': menu([leaf('mac.insert.chartBar', 'chartBar')]),
  'mac.insert.chartLine': chartGallery.line,
  'mac.insert.chartArea': menu([leaf('mac.insert.chartArea', 'chartArea')]),
  'mac.insert.chartPie': chartGallery.pie,
  'mac.insert.chartScatter': chartGallery.scatter,
  'mac.insert.chartHierarchy': chartGallery.hierarchy,
  'mac.insert.chartStatistical': chartGallery.statistical,
  'mac.insert.chartWaterfall': chartGallery.waterfall,
  'mac.insert.chartCombo': chartGallery.combo,
  'mac.insert.chartMap': chartGallery.map,
  // Each type opens the sparkline dialog with that type chosen.
  'mac.insert.sparkline': menu([
    { id: 'mac.insert.sparkline.line', label: 'Line' },
    { id: 'mac.insert.sparkline.column', label: 'Column' },
    { id: 'mac.insert.sparkline.winLoss', label: 'Win/Loss' },
  ]),
  'mac.insert.textBox': unavailable(
    [
      { id: 'mac.insert.textBox.horizontal', label: 'Horizontal Text Box' },
      { id: 'mac.insert.textBox.vertical', label: 'Vertical Text Box' },
    ],
    EMBEDDED_UNSUPPORTED,
  ),
  // The WordArt gallery carries styles, not named entries.
  'mac.insert.wordArt': [],
  'mac.insert.equation': unavailable(
    [{ id: 'mac.insert.equation.insert', label: 'Equation' }],
    NOT_IMPLEMENTED,
  ),
  // Only the whole-stroke eraser exists; point and segment erasing are not built.
  'mac.draw.eraser': menu([
    { id: 'mac.draw.eraser.stroke', label: 'Stroke Eraser' },
    {
      id: 'mac.draw.eraser.small',
      label: 'Small Eraser',
      disabled: true,
      disabledReason: NOT_IMPLEMENTED.en,
    },
    {
      id: 'mac.draw.eraser.medium',
      label: 'Medium Eraser',
      disabled: true,
      disabledReason: NOT_IMPLEMENTED.en,
    },
    {
      id: 'mac.draw.eraser.segment',
      label: 'Segment Eraser',
      disabled: true,
      disabledReason: NOT_IMPLEMENTED.en,
    },
  ]),
  // Adding a pen kind to the strip is not built; the four pens have fixed buttons.
  'mac.draw.add': unavailable(
    [
      { id: 'mac.draw.add.pen', label: 'Pen' },
      { id: 'mac.draw.add.highlighter', label: 'Highlighter' },
      { id: 'mac.draw.add.pencil', label: 'Pencil' },
    ],
    NOT_IMPLEMENTED,
  ),
  // Office and Grayscale are the themes the host can apply; the commands below them
  // are the desktop's online-gallery, browse and save-to-file entries, which it has no backing for.
  'mac.page.theme': menu([
    { id: 'mac.page.theme.office', label: 'Office' },
    { id: 'mac.page.theme.gray', label: 'Grayscale' },
    {
      id: 'mac.page.theme.enableUpdates',
      label: 'Enable Content Updates',
      disabled: true,
      disabledReason: OFFICE_REQUIRED.en,
    },
    {
      id: 'mac.page.theme.browse',
      label: 'Browse for Themes...',
      disabled: true,
      disabledReason: NOT_IMPLEMENTED.en,
    },
    {
      id: 'mac.page.theme.save',
      label: 'Save Current Theme...',
      disabled: true,
      disabledReason: NOT_IMPLEMENTED.en,
    },
  ]),
  'mac.page.color': menu([
    { id: 'mac.page.color.office', label: 'Office' },
    { id: 'mac.page.color.gray', label: 'Grayscale' },
  ]),
  'mac.page.font': menu([
    { id: 'mac.page.font.office', label: 'Office' },
    { id: 'mac.page.font.gray', label: 'Grayscale' },
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
    { id: 'mac.page.break.insert', label: 'Insert Page Break' },
    { id: 'mac.page.break.remove', label: 'Remove Page Break' },
    {
      id: 'mac.page.break.reset',
      label: 'Reset All Page Breaks',
    },
  ]),
  'mac.formulas.autoSum': menu([
    { id: 'mac.autosum.SUM', label: 'Sum' },
    { id: 'mac.autosum.AVERAGE', label: 'Average' },
    { id: 'mac.autosum.COUNT', label: 'Count Numbers' },
    { id: 'mac.autosum.MAX', label: 'Max' },
    { id: 'mac.autosum.MIN', label: 'Min' },
    { id: 'mac.formulas.category.all', label: 'More Functions...' },
  ]),
  'mac.formulas.recent': menu([{ id: 'mac.formulas.category.recent', label: 'More Functions...' }]),
  'mac.formulas.financial': functionMenu('financial'),
  'mac.formulas.logical': functionMenu('logical'),
  'mac.formulas.text': functionMenu('text'),
  'mac.formulas.dateTime': functionMenu('datetime'),
  'mac.formulas.lookup': functionMenu('lookup'),
  'mac.formulas.math': functionMenu('math'),
  'mac.formulas.more': menu([]),
  // Python is a cloud feature the host does not provide, so its menu stays greyed out.
  'mac.formulas.reset': unavailable(
    [
      { id: 'mac.formulas.reset.clear', label: 'Reset' },
      { id: 'mac.formulas.reset.runtime', label: 'Reset Runtime' },
      separator,
      {
        id: 'mac.formulas.reset.version',
        label: 'Version',
        submenu: [
          { id: 'mac.formulas.reset.versionStable', label: 'Stable' },
          { id: 'mac.formulas.reset.versionLatest', label: 'Latest' },
          separator,
          { id: 'mac.formulas.reset.versionLearnMore', label: 'Learn More' },
        ],
      },
    ],
    OFFICE_REQUIRED,
  ),
  'mac.formulas.defineName': menu([
    { id: 'mac.formulas.defineName.define', label: 'Define Name...' },
    {
      id: 'mac.formulas.defineName.apply',
      label: 'Apply...',
      disabled: true,
      disabledReason: NOT_IMPLEMENTED.en,
    },
  ]),
  'mac.formulas.useInFormula': menu([
    { id: 'mac.formulas.useInFormula.paste', label: 'Paste Names...' },
  ]),
  'mac.formulas.calcOptions': menu([
    { id: 'mac.formulas.calc.auto', label: 'Automatic' },
    {
      id: 'mac.formulas.calc.autoNoTable',
      label: 'Partial',
    },
    { id: 'mac.formulas.calc.manual', label: 'Manual' },
    separator,
    { id: 'mac.formulas.calc.formatStale', label: 'Format Stale Values', check: 'checkbox' },
    separator,
    {
      id: 'mac.formulas.calc.compatVersion',
      label: 'Compatibility Version',
      submenu: [
        ...COMPATIBILITY_VERSIONS.map(
          (version): MacRibbonMenuItem => ({
            id: `mac.formulas.calc.compatVersion.${version}`,
            label: `Version ${version}`,
            check: 'radio',
          }),
        ),
        separator,
        {
          id: 'mac.formulas.calc.compatVersion.learnMore',
          label: 'Learn More',
          disabled: true,
          disabledReason: NOT_IMPLEMENTED.en,
        },
      ],
    },
  ]),
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
    // Its entries are the active sheet's cycle members, listed when the menu opens.
    { id: 'mac.formulas.errorCheck.circular', label: 'Circular References', submenu: [] },
  ]),
  'mac.data.validation': menu([
    { id: 'mac.data.validation.settings', label: 'Data Validation...' },
    { id: 'mac.data.validation.circleInvalid', label: 'Circle Invalid Data' },
    { id: 'mac.data.validation.clearCircles', label: 'Clear Validation Circles' },
    { id: 'mac.data.validation.clearRules', label: 'Clear Validation' },
  ]),
  'mac.data.whatIf': menu([
    {
      id: 'mac.data.scenarioManager',
      label: 'Scenario Manager',
      disabled: true,
      disabledReason: NOT_IMPLEMENTED.en,
    },
    { id: 'mac.data.goalSeek', label: 'Goal Seek' },
    { id: 'mac.data.dataTable', label: 'Data Table' },
  ]),
  // Alt Text and Accessibility Help have nothing behind them.
  'mac.review.accessibility': menu([
    { id: 'mac.review.accessibility.check', label: 'Check Accessibility' },
    ...menu([{ id: 'mac.review.accessibility.altText', label: 'Alt Text' }]).map((item) => ({
      ...item,
      disabled: true,
      disabledReason: NOT_IMPLEMENTED.en,
    })),
    { id: 'mac.review.accessibility.formatTable', label: 'Format as Table' },
    { id: 'mac.review.accessibility.unmerge', label: 'Unmerge Cells' },
    ...menu([{ id: 'mac.review.accessibility.help', label: 'Accessibility Help' }]).map((item) => ({
      ...item,
      disabled: true,
      disabledReason: NOT_IMPLEMENTED.en,
    })),
  ]),
  // Ink cannot be hidden or deleted in this host, so every entry is greyed out.
  'mac.review.hideInk': [
    ...unavailable([{ id: 'mac.review.hideInk.hide', label: 'Hide Ink' }], EMBEDDED_UNSUPPORTED),
    ...unavailable(
      [
        { id: 'mac.review.hideInk.deleteSheet', label: 'Delete All Ink on Sheet' },
        { id: 'mac.review.hideInk.deleteWorkbook', label: 'Delete All Ink in Workbook' },
      ],
      NOT_IMPLEMENTED,
    ),
  ],
  'mac.review.notes': menu([
    { id: 'mac.review.newNote', label: 'New Note' },
    { id: 'mac.review.previousNote', label: 'Previous Note' },
    { id: 'mac.review.nextNote', label: 'Next Note' },
    { id: 'mac.review.showHideNote', label: 'Show/Hide Note' },
    { id: 'mac.review.showNotes', label: 'Show All Notes' },
    { id: 'mac.review.convertNotes', label: 'Convert to Comments' },
  ]),
  // Only the measured default colour is built; Focus Cell Color stays greyed out.
  'mac.view.cellFocus': [
    { id: 'mac.view.cellFocus.focusCell', label: 'Focus Cell', check: 'checkbox' },
    ...unavailable(
      [{ id: 'mac.view.cellFocus.color', label: 'Focus Cell Color' }],
      NOT_IMPLEMENTED,
    ),
    { id: 'mac.view.cellFocus.autoHighlight', label: 'Show Auto-Highlight', check: 'checkbox' },
  ],
  'mac.view.freeze': menu([
    { id: 'mac.view.freeze.panes', label: 'Freeze Panes' },
    { id: 'mac.view.freeze.firstRow', label: 'Freeze Top Row' },
    { id: 'mac.view.freeze.firstColumn', label: 'Freeze First Column' },
  ]),
  // Office Scripts is a cloud feature the host does not provide, so New Script's menu stays greyed out.
  'mac.automate.newScript': unavailable(
    [
      { id: 'mac.automate.newScript.editor', label: 'Create in Code Editor' },
      { id: 'mac.automate.newScript.recording', label: 'Create from Recording' },
    ],
    OFFICE_REQUIRED,
  ),
  // Samples opens the script gallery; the other two lists have no script storage behind them.
  'mac.automate.showScripts': [
    ...unavailable(
      [
        { id: 'mac.automate.showScripts.recent', label: 'Recent Scripts' },
        { id: 'mac.automate.showScripts.workbook', label: 'Scripts from this Workbook' },
      ],
      NOT_IMPLEMENTED,
    ),
    ...menu([{ id: 'mac.automate.showScripts.samples', label: 'Samples' }]),
  ],
};

export const macRibbonMenuIdForCommand = (commandId: string): string | null => {
  if (!MAC_RIBBON_MENU_COMMAND_SET.has(commandId)) return null;
  return MAC_RIBBON_MENU_DOM_IDS[commandId] ?? null;
};
