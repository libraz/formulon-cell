import {
  type CatalogFunctionCategory,
  supportedFunctionNames,
} from '../../../commands/function-categories.js';
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
  labelJa?: string;
  disabled?: boolean;
  disabledReason?: string;
  icon?: string;
  /** A section caption rather than a command. */
  heading?: boolean;
  /** The label is the gallery's own wording, not the command's ribbon label. */
  own?: boolean;
}

/** Root command ids whose primary face opens a Mac menu. The mount can use
 * this explicit list to apply menu-first behavior without changing the
 * generic ribbon activation manifest. */
export const MAC_RIBBON_MENU_COMMANDS = [
  'mac.insert.forms',
  'mac.insert.pictureFromData',
  'mac.insert.photo',
  'mac.insert.shapes',
  'mac.insert.threeDModel',
  'mac.insert.smartArt',
  'mac.insert.screenshot',
  'mac.insert.recommendedChart',
  'mac.insert.chartColumn',
  'mac.insert.chartBar',
  'mac.insert.chartLine',
  'mac.insert.chartArea',
  'mac.insert.chartPie',
  'mac.insert.chartScatter',
  'mac.insert.chartHierarchy',
  'mac.insert.chartStatistical',
  'mac.insert.chartWaterfall',
  'mac.insert.chartCombo',
  'mac.insert.chartMap',
  'mac.insert.sparkline',
  'mac.insert.textBox',
  'mac.insert.wordArt',
  'mac.insert.equation',
  'mac.draw.eraser',
  'mac.draw.add',
  'mac.page.theme',
  'mac.page.color',
  'mac.page.font',
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
  'mac.formulas.reset',
  'mac.formulas.defineName',
  'mac.formulas.useInFormula',
  'mac.formulas.calcOptions',
  'mac.formulas.removeArrows',
  'mac.formulas.errorCheck',
  'mac.data.dataTypes',
  'mac.data.pictureFromData',
  'mac.data.sortFilter',
  'mac.data.validation',
  'mac.data.whatIf',
  'mac.review.accessibility',
  'mac.review.hideInk',
  'mac.review.notes',
  'mac.view.cellFocus',
  'mac.view.freeze',
  'mac.automate.newScript',
  'mac.automate.showScripts',
] as const;

export const MAC_RIBBON_MENU_COMMAND_SET: ReadonlySet<string> = new Set(MAC_RIBBON_MENU_COMMANDS);

const menu = (items: readonly MacRibbonMenuItem[]): readonly MacRibbonMenuItem[] =>
  items.map((item) => {
    const label = labelFor(item.id);
    return label ? { ...item, label: label.en, labelJa: label.ja } : item;
  });
/** A menu whose entries the host cannot run: every row greyed out with the same reason. */
const unavailable = (
  items: readonly MacRibbonMenuItem[],
  reason: Label,
): readonly MacRibbonMenuItem[] =>
  menu(items).map((item) => ({ ...item, disabled: true, disabledReason: reason.en }));
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

/** A gallery section caption. */
const section = (id: string, label: string, labelJa: string): MacRibbonMenuItem => ({
  id: `mac.insert.heading.${id}`,
  label,
  labelJa,
  heading: true,
  own: true,
});

/** A gallery entry; `enabled` entries run their chart command, the rest stay greyed out. */
const galleryItem = (
  id: string,
  label: string,
  labelJa: string,
  enabled = false,
): MacRibbonMenuItem => ({
  id,
  label,
  labelJa,
  own: true,
  disabled: !enabled,
  disabledReason: enabled ? undefined : NOT_IMPLEMENTED.en,
});

const chartGallery = {
  column: [
    section('column2d', '2-D Column', '2-D 縦棒'),
    galleryItem('mac.insert.chartColumn', 'Clustered Column', '集合縦棒', true),
    galleryItem('mac.insert.chartColumnStacked', 'Stacked Column', '積み上げ縦棒', true),
    galleryItem('mac.insert.chartColumn100', '100% Stacked Column', '100% 積み上げ縦棒', true),
    section('column3d', '3-D Column', '3-D 縦棒'),
    galleryItem('mac.insert.chartColumn3d', '3-D Clustered Column', '3-D 集合縦棒'),
    galleryItem('mac.insert.chartColumnStacked3d', '3-D Stacked Column', '3-D 積み上げ縦棒'),
    galleryItem('mac.insert.chartColumn1003d', '3-D 100% Stacked Column', '3-D 100% 積み上げ縦棒'),
    galleryItem('mac.insert.chartColumnTrue3d', '3-D Column', '3-D 縦棒グラフ'),
    section('bar2d', '2-D Bar', '2-D 横棒'),
    galleryItem('mac.insert.chartBar', 'Clustered Bar', '集合横棒', true),
    galleryItem('mac.insert.chartBarStacked', 'Stacked Bar', '積み上げ横棒', true),
    galleryItem('mac.insert.chartBar100', '100% Stacked Bar', '100% 積み上げ横棒', true),
    section('bar3d', '3-D Bar', '3-D 横棒'),
    galleryItem('mac.insert.chartBar3d', '3-D Clustered Bar', '3-D 集合横棒'),
    galleryItem('mac.insert.chartBarStacked3d', '3-D Stacked Bar', '3-D 積み上げ横棒'),
    galleryItem('mac.insert.chartBar1003d', '3-D 100% Stacked Bar', '3-D 100% 積み上げ横棒'),
  ],
  line: [
    section('line2d', '2-D Line', '2-D 折れ線'),
    galleryItem('mac.insert.chartLine', 'Line', '折れ線', true),
    galleryItem('mac.insert.chartLineStacked', 'Stacked Line', '積み上げ折れ線', true),
    galleryItem('mac.insert.chartLine100', '100% Stacked Line', '100% 積み上げ折れ線', true),
    galleryItem('mac.insert.chartLineMarkers', 'Line with Markers', 'マーカー付き折れ線'),
    galleryItem(
      'mac.insert.chartLineStackedMarkers',
      'Stacked Line with Markers',
      'マーカー付き積み上げ折れ線',
    ),
    galleryItem(
      'mac.insert.chartLine100Markers',
      '100% Stacked Line with Markers',
      'マーカー付き 100% 積み上げ折れ線',
    ),
    section('line3d', '3-D Line', '3-D 折れ線'),
    galleryItem('mac.insert.chartLine3d', '3-D Line', '3-D 折れ線'),
    section('area2d', '2-D Area', '2-D 面'),
    galleryItem('mac.insert.chartArea', 'Area', '面', true),
    galleryItem('mac.insert.chartAreaStacked', 'Stacked Area', '積み上げ面', true),
    galleryItem('mac.insert.chartArea100', '100% Stacked Area', '100% 積み上げ面', true),
    section('area3d', '3-D Area', '3-D 面'),
    galleryItem('mac.insert.chartArea3d', '3-D Area', '3-D 面'),
    galleryItem('mac.insert.chartAreaStacked3d', '3-D Stacked Area', '3-D 積み上げ面'),
    galleryItem('mac.insert.chartArea1003d', '3-D 100% Stacked Area', '3-D 100% 積み上げ面'),
  ],
  pie: [
    section('pie2d', '2-D Pie', '2-D 円'),
    galleryItem('mac.insert.chartPie', 'Pie', '円', true),
    galleryItem('mac.insert.chartPieOfPie', 'Pie of Pie', '補助円グラフ付き円'),
    galleryItem('mac.insert.chartBarOfPie', 'Bar of Pie', '補助縦棒付き円'),
    section('pie3d', '3-D Pie', '3-D 円'),
    galleryItem('mac.insert.chartPie3d', '3-D Pie', '3-D 円'),
    section('doughnut', 'Doughnut', 'ドーナツ'),
    galleryItem('mac.insert.chartDoughnut', 'Doughnut', 'ドーナツ', true),
  ],
  scatter: [
    section('scatter', 'Scatter', '散布図'),
    galleryItem('mac.insert.chartScatter', 'Scatter', '散布図', true),
    galleryItem('mac.insert.chartScatterSmooth', 'Scatter with Smooth Lines', '散布図 (平滑線)'),
    galleryItem(
      'mac.insert.chartScatterLinesMarkers',
      'Scatter with Straight Lines and Markers',
      '散布図 (直線とマーカー)',
    ),
    galleryItem('mac.insert.chartScatterLines', 'Scatter with Straight Lines', '散布図 (直線)'),
    section('bubble', 'Bubble', 'バブル'),
    galleryItem('mac.insert.chartBubble', 'Bubble', 'バブル'),
    galleryItem('mac.insert.chartBubble3d', '3-D Bubble', '3-D 効果付きバブル'),
  ],
  hierarchy: [
    section('treemap', 'Treemap', 'ツリーマップ'),
    galleryItem('mac.insert.chartTreemap', 'Treemap', 'ツリーマップ'),
    section('sunburst', 'Sunburst', 'サンバースト'),
    galleryItem('mac.insert.chartSunburst', 'Sunburst', 'サンバースト'),
  ],
  statistical: [
    section('histogramGroup', 'Histogram', 'ヒストグラム'),
    galleryItem('mac.insert.chartHistogram', 'Histogram', 'ヒストグラム'),
    galleryItem('mac.insert.chartPareto', 'Pareto', 'パレート図'),
    section('boxWhisker', 'Box and Whisker', '箱ひげ図'),
    galleryItem('mac.insert.chartBoxWhisker', 'Box and Whisker', '箱ひげ図'),
  ],
  waterfall: [
    section('waterfall', 'Waterfall', 'ウォーターフォール'),
    galleryItem('mac.insert.chartWaterfall.plot', 'Waterfall', 'ウォーターフォール'),
    section('funnel', 'Funnel', 'じょうご'),
    galleryItem('mac.insert.chartFunnel', 'Funnel', 'じょうご'),
    section('stock', 'Stock', '株価'),
    galleryItem(
      'mac.insert.chartStockHlc',
      'Stock: High-Low-Close',
      '株価チャート (高値-安値-終値)',
    ),
    galleryItem(
      'mac.insert.chartStockOhlc',
      'Stock: Open-High-Low-Close',
      '株価チャート (始値-高値-安値-終値)',
    ),
    galleryItem(
      'mac.insert.chartStockVhlc',
      'Stock: Volume-High-Low-Close',
      '株価チャート (出来高-高値-安値-終値)',
    ),
    galleryItem(
      'mac.insert.chartStockVohlc',
      'Stock: Volume-Open-High-Low-Close',
      '株価チャート (出来高-始値-高値-安値-終値)',
    ),
    section('surface', 'Surface', '等高線'),
    galleryItem('mac.insert.chartSurface', '3-D Surface', '3-D の表面'),
    galleryItem(
      'mac.insert.chartSurfaceWireframe',
      'Wireframe 3-D Surface',
      'ワイヤーフレーム 3-D 等高線',
    ),
    galleryItem('mac.insert.chartContour', 'Contour', '等高線'),
    galleryItem('mac.insert.chartContourWireframe', 'Wireframe Contour', 'ワイヤーフレーム等高線'),
    section('radar', 'Radar', 'レーダー'),
    galleryItem('mac.insert.chartRadar', 'Radar', 'レーダー チャート'),
    galleryItem('mac.insert.chartRadarMarkers', 'Radar with Markers', 'マーカー付きレーダー'),
    galleryItem('mac.insert.chartRadarFilled', 'Filled Radar', '塗りつぶしレーダー'),
  ],
  combo: [
    section('combo', 'Combo', '組み合わせ'),
    galleryItem('mac.insert.chartComboColumnLine', 'Clustered Column - Line', '集合縦棒 - 折れ線'),
    galleryItem(
      'mac.insert.chartComboColumnLineSecondary',
      'Clustered Column - Line on Secondary Axis',
      '集合縦棒 - 第 2 軸の折れ線',
    ),
    galleryItem(
      'mac.insert.chartComboAreaColumn',
      'Stacked Area - Clustered Column',
      '積み上げ面 - 集合縦棒',
    ),
  ],
  map: [
    section('filledMap', 'Filled Map', '塗り分けマップ'),
    galleryItem('mac.insert.chartFilledMap', 'Filled Map', '塗り分けマップ'),
  ],
} as const satisfies Record<string, readonly MacRibbonMenuItem[]>;

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
      label: 'No charts are recommended for the selected data.',
      labelJa: '選択したデータに対するおすすめグラフはありません。',
      own: true,
      disabled: true,
    },
    {
      id: 'mac.insert.recommendedChart.apply',
      label: 'Recommended Charts',
      labelJa: 'おすすめグラフ',
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
    { id: 'mac.page.theme.office', label: 'Office', labelJa: 'Office' },
    { id: 'mac.page.theme.gray', label: 'Grayscale', labelJa: 'グレースケール' },
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
    { id: 'mac.page.color.office', label: 'Office', labelJa: 'Office' },
    { id: 'mac.page.color.gray', label: 'Grayscale', labelJa: 'グレースケール' },
  ]),
  'mac.page.font': menu([
    { id: 'mac.page.font.office', label: 'Office', labelJa: 'Office' },
    { id: 'mac.page.font.gray', label: 'Grayscale', labelJa: 'グレースケール' },
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
  // Python is a cloud feature the host does not provide, so its menu stays greyed out.
  'mac.formulas.reset': menu([
    { id: 'mac.formulas.reset.clear', label: 'Reset' },
    { id: 'mac.formulas.reset.runtime', label: 'Reset Runtime' },
    { id: 'mac.formulas.reset.versionStable', label: 'Stable' },
    { id: 'mac.formulas.reset.versionLatest', label: 'Latest' },
    { id: 'mac.formulas.reset.versionLearnMore', label: 'Learn More' },
  ]).map((item) => ({ ...item, disabled: true, disabledReason: OFFICE_REQUIRED.en })),
  'mac.formulas.defineName': menu([
    { id: 'mac.formulas.defineName.define', label: 'Define Name...' },
    {
      id: 'mac.formulas.defineName.apply',
      label: 'Apply...',
      disabled: true,
      disabledReason: NOT_IMPLEMENTED.en,
    },
  ]),
  'mac.formulas.useInFormula': menu([{ id: 'mac.formulas.useInFormula.paste', label: 'Paste...' }]),
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
  // The host has no linked data types, so the gallery lists Excel's entries greyed out.
  'mac.data.dataTypes': [
    galleryItem('mac.data.dataTypes.stocks', 'Stocks (English)', '株式 (English)'),
    galleryItem('mac.data.dataTypes.currencies', 'Currencies (English)', '通貨 (English)'),
    galleryItem('mac.data.dataTypes.geography', 'Geography (English)', '地理 (English)'),
  ],
  'mac.data.pictureFromData': [
    galleryItem('mac.data.pictureFromData.file', 'From File...', 'ファイルからの画像...'),
    galleryItem('mac.data.pictureFromData.clipboard', 'From Clipboard', 'クリップボードからの画像'),
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
  // Only the checker is built; the other entries are the desktop's and have nothing behind them.
  'mac.review.accessibility': menu([
    { id: 'mac.review.accessibility.check', label: 'Check Accessibility' },
    ...menu([
      { id: 'mac.review.accessibility.altText', label: 'Alt Text' },
      { id: 'mac.review.accessibility.formatTable', label: 'Format as Table' },
      { id: 'mac.review.accessibility.unmerge', label: 'Unmerge Cells' },
      { id: 'mac.review.accessibility.help', label: 'Accessibility Help' },
    ]).map((item) => ({ ...item, disabled: true, disabledReason: NOT_IMPLEMENTED.en })),
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
  // Focus Cell is not built yet; the menu carries the desktop's three entries greyed out.
  'mac.view.cellFocus': menu([
    { id: 'mac.view.cellFocus.focusCell', label: 'Focus Cell' },
    { id: 'mac.view.cellFocus.color', label: 'Focus Cell Color' },
    { id: 'mac.view.cellFocus.autoHighlight', label: 'Show Auto-Highlight' },
  ]).map((item) => ({ ...item, disabled: true, disabledReason: NOT_IMPLEMENTED.en })),
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

/** DOM id of the Mac Formulas "More Functions" menu. */
export const MAC_FORMULAS_MORE_MENU_ID = 'menu-mac-formulas-more';

export const macRibbonMenuIdForCommand = (commandId: string): string | null => {
  if (!MAC_RIBBON_MENU_COMMAND_SET.has(commandId)) return null;
  const ids: Readonly<Record<string, string>> = {
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
  return ids[commandId] ?? null;
};
