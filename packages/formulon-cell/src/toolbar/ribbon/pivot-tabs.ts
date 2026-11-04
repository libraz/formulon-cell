import type {
  RibbonCommand,
  RibbonGroupModel,
  RibbonTabModel,
  ToolbarLang,
} from '../ribbon-model.js';
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

/** The two tabs a PivotTable report adds. Commands the host cannot run stay disabled in place. */
export const buildPivotContextualTabs = (
  lang: ToolbarLang,
  tabLabels: { pivotAnalyze: string; pivotDesign: string },
): RibbonTabModel[] => {
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
  const group = (commands: RibbonCommand[]): RibbonGroupModel => ({
    title: '',
    commands,
    variant: 'compact',
  });
  return [
    {
      id: 'pivotAnalyze',
      label: tabLabels.pivotAnalyze,
      contextual: true,
      groups: [
        group([
          command('pivotName', 'name', { kind: 'input' }),
          command('pivotOptions', 'options'),
        ]),
        group([
          command('pivotActiveField', 'activeField', { kind: 'input' }),
          command('pivotFieldSettings', 'fieldSettings'),
          command('pivotExpandField', 'expandField'),
          command('pivotCollapseField', 'collapseField'),
        ]),
        group([command('pivotGroupSelection', 'groupSelection')]),
        group([
          command('pivotInsertSlicer', 'insertSlicer'),
          command('pivotInsertTimeline', 'insertTimeline'),
          command('pivotFilterConnections', 'filterConnections'),
        ]),
        group([command('pivotRefresh', 'refresh'), command('pivotChangeSource', 'changeSource')]),
        group([
          command('pivotClear', 'clear'),
          command('pivotSelect', 'select'),
          command('pivotMove', 'move'),
        ]),
        group([command('pivotFieldsItemsSets', 'fieldsItemsSets')]),
        group([command('pivotChart', 'pivotChart')]),
        group([
          command('pivotFieldList', 'fieldList'),
          command('pivotButtons', 'buttons'),
          command('pivotHeaders', 'headers'),
        ]),
      ],
    },
    {
      id: 'pivotDesign',
      label: tabLabels.pivotDesign,
      contextual: true,
      groups: [
        group([
          command('pivotSubtotals', 'subtotals'),
          command('pivotGrandTotals', 'grandTotals'),
          command('pivotReportLayout', 'reportLayout'),
          command('pivotBlankRows', 'blankRows'),
        ]),
        group([
          command('pivotRowHeaders', 'rowHeaders'),
          command('pivotColHeaders', 'colHeaders'),
          command('pivotRowStripes', 'rowStripes'),
          command('pivotColStripes', 'colStripes'),
        ]),
        group([command('pivotStyleGallery', 'quickStyles', { kind: 'input' })]),
      ],
    },
  ];
};
