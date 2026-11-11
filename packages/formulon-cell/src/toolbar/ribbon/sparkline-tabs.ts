import type {
  RibbonCommand,
  RibbonGroupModel,
  RibbonTabModel,
  ToolbarText,
} from '../ribbon-model.js';

/** The ribbon command ids of the Sparkline tab, by the control they name. */
export const SPARKLINE_COMMANDS = {
  editData: 'sparklineEditData',
  styleGallery: 'sparklineStyleGallery',
  color: 'sparklineColor',
  markerColor: 'sparklineMarkerColor',
  axis: 'sparklineAxis',
  group: 'sparklineGroup',
  ungroup: 'sparklineUngroup',
  clear: 'sparklineClear',
} as const;

/** Menu ids of the Sparkline tab's dropdowns. */
export const SPARKLINE_MENU_IDS = {
  editData: 'menu-sparkline-edit-data',
  style: 'menu-sparkline-style',
  color: 'menu-sparkline-color',
  markerColor: 'menu-sparkline-marker-color',
  axis: 'menu-sparkline-axis',
  clear: 'menu-sparkline-clear',
} as const;

/** The Plot Data Right-to-Left axis option stays greyed until the engine models it. */
export const SPARKLINE_AXIS_DISABLED_ITEMS: readonly string[] = ['rtl'];

/** The Sparkline contextual tab, shown while the active cell hosts a sparkline. */
export const buildSparklineDesignTab = (tr: ToolbarText): RibbonTabModel => {
  const t = tr.sparkline;
  const command = (id: string, label: string, icon: string): RibbonCommand => ({
    id,
    title: label,
    label,
    icon,
    kind: 'wide',
  });
  const group = (title: string, commands: RibbonCommand[]): RibbonGroupModel => ({
    title,
    commands,
    variant: 'tiles',
  });
  return {
    id: 'sparklineDesign',
    label: tr.tabs.sparklineDesign,
    contextual: true,
    groups: [
      group(t.title, [command(SPARKLINE_COMMANDS.editData, t.editData, 'chartSelectData')]),
      group(t.type, [
        command('sparklineTypeLine', t.line, 'sparkline'),
        command('sparklineTypeColumn', t.column, 'sparklineColumn'),
        command('sparklineTypeWinLoss', t.winLoss, 'sparklineWinLoss'),
      ]),
      group(t.show, [
        command('sparklineShowHigh', t.highPoint, 'sparklineHigh'),
        command('sparklineShowLow', t.lowPoint, 'sparklineLow'),
        command('sparklineShowNegative', t.negativePoints, 'sparklineNegative'),
        command('sparklineShowFirst', t.firstPoint, 'sparklineFirst'),
        command('sparklineShowLast', t.lastPoint, 'sparklineLast'),
        command('sparklineShowMarkers', t.markers, 'sparklineMarkers'),
      ]),
      group(t.style, [
        command(SPARKLINE_COMMANDS.styleGallery, t.style, 'chartStyles'),
        command(SPARKLINE_COMMANDS.color, t.color, 'shapeOutline'),
        command(SPARKLINE_COMMANDS.markerColor, t.markerColor, 'fillColor'),
      ]),
      group(t.group, [
        command(SPARKLINE_COMMANDS.axis, t.axis, 'sparklineAxis'),
        command(SPARKLINE_COMMANDS.group, t.group, 'outlineGroup'),
        command(SPARKLINE_COMMANDS.ungroup, t.ungroup, 'outlineUngroup'),
        command(SPARKLINE_COMMANDS.clear, t.clear, 'clearAll'),
      ]),
    ],
  };
};
