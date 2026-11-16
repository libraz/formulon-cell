import type {
  RibbonCommand,
  RibbonGroupModel,
  RibbonTabModel,
  ToolbarLang,
} from '../../ribbon-model.js';
import { labels, macGroupTitle } from './labels.js';
import { MAC_OBJECT_TAB_DISABLED_COMMAND_IDS } from './object-tab-ids.js';

/** Wording of the Mac table and chart tabs: Japanese is measured, other languages read as English. */
const L = (ja: string, en: string): Readonly<Record<ToolbarLang, string>> => ({
  ja,
  en,
  de: en,
  fr: en,
  zh: en,
  ko: en,
  th: en,
});

const LABELS = {
  tableTab: L('テーブル', 'Table'),
  tableName: L('テーブル名:', 'Table Name:'),
  summarizePivot: L('ピボットテーブルで集計', 'Summarize with PivotTable'),
  removeDuplicates: L('重複を削除', 'Remove Duplicates'),
  convertRange: L('範囲に変換', 'Convert to Range'),
  insertSlicer: L('スライサーの挿入', 'Insert Slicer'),
  refresh: L('更新', 'Refresh'),
  headerRow: L('見出し行', 'Header Row'),
  totalRow: L('集計行', 'Total Row'),
  bandedRows: L('縞模様 (行)', 'Banded Rows'),
  firstColumn: L('最初の列', 'First Column'),
  lastColumn: L('最後の列', 'Last Column'),
  bandedColumns: L('縞模様 (列)', 'Banded Columns'),
  filterButton: L('フィルター ボタン', 'Filter Button'),
  quickStyles: L('クイック スタイル', 'Quick Styles'),
  addElement: L('グラフ要素を追加', 'Add Chart Element'),
  quickLayout: L('クイック レイアウト', 'Quick Layout'),
  changeColors: L('色の変更', 'Change Colors'),
  switchRowColumn: L('行/列の切り替え', 'Switch Row/Column'),
  selectData: L('グラフ データの選択', 'Select Data'),
  changeType: L('グラフの種類の変更', 'Change Chart Type'),
  moveChart: L('グラフの移動', 'Move Chart'),
  chartElement: L('グラフ要素', 'Chart Elements'),
  resetStyle: L('リセットしてスタイルに合わせる', 'Reset to Match Style'),
  shapes: L('図形', 'Shapes'),
  changeShape: L('図形の変更', 'Change Shape'),
  shapeFill: L('図形の塗りつぶし', 'Shape Fill'),
  shapeOutline: L('図形の枠線', 'Shape Outline'),
  shapeEffects: L('図形の効果', 'Shape Effects'),
  textFill: L('文字の塗りつぶし', 'Text Fill'),
  textOutline: L('文字の輪郭', 'Text Outline'),
  textEffects: L('文字の効果', 'Text Effects'),
  bringForward: L('前面へ移動', 'Bring Forward'),
  sendBackward: L('背面へ移動', 'Send Backward'),
  selectionPane: L('選択ウインドウ', 'Selection Pane'),
  arrangeObjects: L('オブジェクトの並べ替え', 'Arrange Objects'),
  align: L('整列', 'Align'),
  group: L('グループ化', 'Group'),
  rotate: L('回転', 'Rotate'),
  lockAspect: L('縦横比を固定', 'Lock Aspect Ratio'),
  formatPane: L('書式ウィンドウ', 'Format Pane'),
} as const;

/** A shared contextual tab under the name the Mac draws, where the Mac names it differently. */
export const withMacTabLabel = (tab: RibbonTabModel, lang: ToolbarLang): RibbonTabModel => {
  const label = labels[tab.id]?.[lang];
  return label ? { ...tab, label } : tab;
};

/** The Excel for Mac table tab and chart tabs, in the groups and order Excel shows them. */
export const buildMacObjectTabs = (
  lang: ToolbarLang,
  tabLabels: { chartDesign: string; chartFormat: string },
): RibbonTabModel[] => {
  const text = (key: keyof typeof LABELS): string => LABELS[key][lang];
  const command = (
    id: string,
    key: keyof typeof LABELS,
    opts: { kind?: RibbonCommand['kind'] } = {},
  ): RibbonCommand => ({
    id,
    title: text(key),
    label: text(key),
    kind: opts.kind ?? 'wide',
    ...(MAC_OBJECT_TAB_DISABLED_COMMAND_IDS.includes(id) ? { disabled: true } : {}),
  });
  const group = (commands: RibbonCommand[], title?: string): RibbonGroupModel => ({
    title: title === undefined ? '' : macGroupTitle(title, lang),
    commands,
    variant: 'compact',
  });
  return [
    {
      id: 'tableDesign',
      label: text('tableTab'),
      contextual: true,
      groups: [
        group([command('tableName', 'tableName', { kind: 'input' })], 'Properties'),
        group(
          [
            command('tableSummarizePivot', 'summarizePivot'),
            command('mac.data.removeDuplicates', 'removeDuplicates'),
            command('tableConvertRange', 'convertRange'),
            command('mac.insert.slicer', 'insertSlicer', { kind: 'large' }),
          ],
          'Tools (Table)',
        ),
        group([command('tableRefresh', 'refresh', { kind: 'large' })], 'External Table Data'),
        group(
          [
            command('tableHeaderRow', 'headerRow'),
            command('tableTotalRow', 'totalRow'),
            command('tableBandedRows', 'bandedRows'),
            command('tableFirstColumn', 'firstColumn'),
            command('tableLastColumn', 'lastColumn'),
            command('tableBandedColumns', 'bandedColumns'),
            command('tableFilterButton', 'filterButton'),
          ],
          'Table Style Options',
        ),
        group([command('tableQuickStyles', 'quickStyles')], 'Table Styles'),
      ],
    },
    {
      id: 'chartDesign',
      label: tabLabels.chartDesign,
      contextual: true,
      groups: [
        group(
          [command('chartAddElement', 'addElement'), command('chartQuickLayout', 'quickLayout')],
          'Chart Layouts',
        ),
        group(
          [command('chartColors', 'changeColors'), command('chartQuickStyles', 'quickStyles')],
          'Chart Styles',
        ),
        group(
          [
            command('chartSwitchRowColumn', 'switchRowColumn'),
            command('chartSelectData', 'selectData'),
          ],
          'Data',
        ),
        group([command('chartChangeType', 'changeType')], 'Type'),
        group([command('chartMove', 'moveChart')], 'Location'),
      ],
    },
    {
      id: 'chartFormat',
      label: tabLabels.chartFormat,
      contextual: true,
      groups: [
        group(
          [command('chartElementSelect', 'chartElement'), command('chartResetStyle', 'resetStyle')],
          'Current Selection',
        ),
        group(
          [command('chartShapes', 'shapes'), command('chartChangeShape', 'changeShape')],
          'Insert Shapes',
        ),
        group(
          [
            command('chartShapeQuickStyles', 'quickStyles'),
            command('chartShapeFill', 'shapeFill'),
            command('chartShapeOutline', 'shapeOutline'),
            command('chartShapeEffects', 'shapeEffects'),
          ],
          'Shape Styles',
        ),
        group(
          [
            command('chartTextQuickStyles', 'quickStyles'),
            command('chartTextFill', 'textFill'),
            command('chartTextOutline', 'textOutline'),
            command('chartTextEffects', 'textEffects'),
          ],
          'WordArt Styles',
        ),
        group(
          [
            command('objectBringForward', 'bringForward'),
            command('objectSendBackward', 'sendBackward'),
            command('objectSelectionPane', 'selectionPane'),
            command('chartArrangeObjects', 'arrangeObjects'),
            command('chartAlign', 'align'),
            command('chartGroup', 'group'),
            command('chartRotate', 'rotate'),
          ],
          'Arrange',
        ),
        group([command('chartLockAspect', 'lockAspect')], 'Size'),
        group([command('chartFormatSelection', 'formatPane')], 'Format'),
      ],
    },
  ];
};
