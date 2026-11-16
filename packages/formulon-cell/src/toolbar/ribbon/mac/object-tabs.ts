import type {
  RibbonCommand,
  RibbonGroupModel,
  RibbonTabModel,
  ToolbarLang,
} from '../../ribbon-model.js';
import { macBundleText } from './bundle-text.js';
import { labels, macGroupTitle } from './labels.js';
import { MAC_OBJECT_TAB_DISABLED_COMMAND_IDS } from './object-tab-ids.js';

/** A table or chart tab wording, generated as the `object-tab:` label of this key. */
type ObjectTabText =
  | 'tableTab'
  | 'tableName'
  | 'summarizePivot'
  | 'removeDuplicates'
  | 'convertRange'
  | 'insertSlicer'
  | 'refresh'
  | 'headerRow'
  | 'totalRow'
  | 'bandedRows'
  | 'firstColumn'
  | 'lastColumn'
  | 'bandedColumns'
  | 'filterButton'
  | 'quickStyles'
  | 'addElement'
  | 'quickLayout'
  | 'changeColors'
  | 'switchRowColumn'
  | 'selectData'
  | 'changeType'
  | 'moveChart'
  | 'chartElement'
  | 'resetStyle'
  | 'shapes'
  | 'changeShape'
  | 'shapeFill'
  | 'shapeOutline'
  | 'shapeEffects'
  | 'textFill'
  | 'textOutline'
  | 'textEffects'
  | 'bringForward'
  | 'sendBackward'
  | 'selectionPane'
  | 'arrangeObjects'
  | 'align'
  | 'group'
  | 'rotate'
  | 'lockAspect'
  | 'formatPane';

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
  const text = (key: ObjectTabText): string => macBundleText(`object-tab:${key}`, lang);
  const command = (
    id: string,
    key: ObjectTabText,
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
