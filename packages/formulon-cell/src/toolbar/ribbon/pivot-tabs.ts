import type {
  RibbonCommand,
  RibbonGroupModel,
  RibbonProfile,
  RibbonTabModel,
  ToolbarLang,
} from '../ribbon-model.js';
import { GROUP_LABELS } from './mac/labels.js';
import { PIVOT_DISABLED_COMMAND_IDS } from './pivot-command-ids.js';

/** Wording of the PivotTable tabs: Japanese is measured, other languages read as English. */
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
  name: L('ピボットテーブル名:', 'PivotTable Name:'),
  options: L('オプション', 'Options'),
  activeField: L('アクティブなフィールド:', 'Active Field:'),
  fieldSettings: L('フィールドの設定', 'Field Settings'),
  expandField: L('フィールドの展開', 'Expand Field'),
  collapseField: L('フィールドの折りたたみ', 'Collapse Field'),
  groupSelection: L('グループの選択', 'Group Selection'),
  insertSlicer: L('スライサーの挿入', 'Insert Slicer'),
  insertTimeline: L('タイムラインの挿入', 'Insert Timeline'),
  filterConnections: L('フィルターの接続', 'Filter Connections'),
  refresh: L('更新', 'Refresh'),
  changeSource: L('データ ソースの変更', 'Change Data Source'),
  clear: L('クリア', 'Clear'),
  select: L('選択', 'Select'),
  move: L('ピボットテーブルの移動', 'Move PivotTable'),
  fieldsItemsSets: L('フィールド/項目/セット', 'Fields, Items & Sets'),
  pivotChart: L('ピボットグラフ', 'PivotChart'),
  fieldList: L('フィールド リスト', 'Field List'),
  buttons: L('+/- ボタン', '+/- Buttons'),
  headers: L('フィールドのヘッダー', 'Field Headers'),
  subtotals: L('小計', 'Subtotals'),
  grandTotals: L('集計', 'Grand Totals'),
  reportLayout: L('レポートのレイアウト', 'Report Layout'),
  blankRows: L('空白行', 'Blank Rows'),
  rowHeaders: L('行見出し', 'Row Headers'),
  colHeaders: L('列見出し', 'Column Headers'),
  rowStripes: L('縞模様 (行)', 'Banded Rows'),
  colStripes: L('縞模様 (列)', 'Banded Columns'),
  quickStyles: L('クイック スタイル', 'Quick Styles'),
} as const;

/** The two tabs a PivotTable report adds. Commands the host cannot run stay disabled in place.
 *  The Mac profile titles the groups and lays out the Design tab the way the desktop draws them;
 *  elsewhere the groups stay untitled. */
export const buildPivotContextualTabs = (
  lang: ToolbarLang,
  tabLabels: { pivotAnalyze: string; pivotDesign: string },
  profile: RibbonProfile = 'default',
): RibbonTabModel[] => {
  const mac = profile === 'excel365Mac';
  const text = (key: keyof typeof LABELS): string => LABELS[key][lang] ?? LABELS[key].en;
  const command = (
    id: string,
    key: keyof typeof LABELS,
    opts: { kind?: RibbonCommand['kind']; icon?: string } = {},
  ): RibbonCommand => ({
    id,
    title: text(key),
    label: text(key),
    icon: opts.icon,
    kind: opts.kind ?? 'wide',
    ...(PIVOT_DISABLED_COMMAND_IDS.includes(id) ? { disabled: true } : {}),
  });
  const group = (commands: RibbonCommand[], title?: string): RibbonGroupModel => ({
    title: mac && title !== undefined ? (GROUP_LABELS[title]?.[lang] ?? title) : '',
    commands,
    variant: 'compact',
  });
  const totals = [
    command('pivotSubtotals', 'subtotals'),
    command('pivotGrandTotals', 'grandTotals'),
  ];
  const layout = [
    command('pivotReportLayout', 'reportLayout'),
    command('pivotBlankRows', 'blankRows'),
  ];
  return [
    {
      id: 'pivotAnalyze',
      label: tabLabels.pivotAnalyze,
      contextual: true,
      groups: [
        group(
          [command('pivotName', 'name', { kind: 'input' }), command('pivotOptions', 'options')],
          'PivotTable',
        ),
        group(
          [
            command('pivotActiveField', 'activeField', { kind: 'input' }),
            command('pivotFieldSettings', 'fieldSettings'),
            command('pivotExpandField', 'expandField'),
            command('pivotCollapseField', 'collapseField'),
          ],
          'Active Field',
        ),
        group([command('pivotGroupSelection', 'groupSelection')], 'Group (PivotTable)'),
        group(
          [
            command('pivotInsertSlicer', 'insertSlicer'),
            command('pivotInsertTimeline', 'insertTimeline'),
            command('pivotFilterConnections', 'filterConnections'),
          ],
          'Filter',
        ),
        group(
          [command('pivotRefresh', 'refresh'), command('pivotChangeSource', 'changeSource')],
          'Data (PivotTable)',
        ),
        group(
          [
            command('pivotClear', 'clear'),
            command('pivotSelect', 'select'),
            command('pivotMove', 'move'),
          ],
          'Actions',
        ),
        group([command('pivotFieldsItemsSets', 'fieldsItemsSets')], 'Calculations'),
        group([command('pivotChart', 'pivotChart')], 'Tools (PivotTable)'),
        group(
          [
            command('pivotFieldList', 'fieldList'),
            command('pivotButtons', 'buttons'),
            command('pivotHeaders', 'headers'),
          ],
          'Show (PivotTable)',
        ),
      ],
    },
    {
      id: 'pivotDesign',
      label: tabLabels.pivotDesign,
      contextual: true,
      groups: [
        ...(mac ? [group([...totals, ...layout], 'Layout')] : [group(totals), group(layout)]),
        group(
          [
            command('pivotRowHeaders', 'rowHeaders'),
            command('pivotColHeaders', 'colHeaders'),
            command('pivotRowStripes', 'rowStripes'),
            command('pivotColStripes', 'colStripes'),
          ],
          'PivotTable Style Options',
        ),
        group([command('pivotQuickStyles', 'quickStyles')], 'PivotTable Styles'),
      ],
    },
  ];
};
