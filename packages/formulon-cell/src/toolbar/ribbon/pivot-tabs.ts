import type {
  RibbonCommand,
  RibbonGroupModel,
  RibbonProfile,
  RibbonTabModel,
  ToolbarLang,
} from '../ribbon-model.js';
import { macBundleText } from './mac/bundle-text.js';
import { macGroupTitle } from './mac/labels.js';
import { PIVOT_DISABLED_COMMAND_IDS } from './pivot-command-ids.js';

/** A PivotTable tab wording, generated as the `pivot-tab:` label of this key. */
type PivotTabText =
  | 'name'
  | 'options'
  | 'activeField'
  | 'fieldSettings'
  | 'expandField'
  | 'collapseField'
  | 'groupSelection'
  | 'insertSlicer'
  | 'insertTimeline'
  | 'filterConnections'
  | 'refresh'
  | 'changeSource'
  | 'clear'
  | 'select'
  | 'move'
  | 'fieldsItemsSets'
  | 'pivotChart'
  | 'fieldList'
  | 'buttons'
  | 'headers'
  | 'subtotals'
  | 'grandTotals'
  | 'reportLayout'
  | 'blankRows'
  | 'rowHeaders'
  | 'colHeaders'
  | 'rowStripes'
  | 'colStripes'
  | 'quickStyles';

/** The two tabs a PivotTable report adds. Commands the host cannot run stay disabled in place.
 *  The Mac profile titles the groups and lays out the Design tab the way the desktop draws them;
 *  elsewhere the groups stay untitled. */
export const buildPivotContextualTabs = (
  lang: ToolbarLang,
  tabLabels: { pivotAnalyze: string; pivotDesign: string },
  profile: RibbonProfile = 'default',
): RibbonTabModel[] => {
  const mac = profile === 'excel365Mac';
  const text = (key: PivotTabText): string => macBundleText(`pivot-tab:${key}`, lang);
  const command = (
    id: string,
    key: PivotTabText,
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
    title: mac && title !== undefined ? macGroupTitle(title, lang) : '',
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
