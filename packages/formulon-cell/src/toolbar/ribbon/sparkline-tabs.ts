import type {
  RibbonCommand,
  RibbonGroupModel,
  RibbonProfile,
  RibbonTabModel,
  ToolbarLang,
  ToolbarText,
} from '../ribbon-model.js';
import { macGroupTitle } from './mac/labels.js';

/** The ribbon command ids of the Sparkline tab, by the control they name. */
export const SPARKLINE_COMMANDS = {
  editData: 'sparklineEditData',
  styleGallery: 'sparklineStyleGallery',
  color: 'sparklineColor',
  markerColor: 'sparklineMarkerColor',
  axis: 'sparklineAxis',
  /** The Mac app's Axis button, which opens the Axis dialog instead of a menu. */
  axisDialog: 'sparklineAxisDialog',
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

/** Group variants of the Sparkline tab: Show stacks its toggles in columns of three, and the last
 *  group stacks Group, Ungroup and Clear beside the large Axis button. */
export const SPARKLINE_STACKED_LAYOUT_GROUP_VARIANTS = ['sparkline-show'] as const;
export const SPARKLINE_MIXED_LAYOUT_GROUP_VARIANTS = ['sparkline-group'] as const;

/** The Sparkline contextual tab, shown while the active cell hosts a sparkline. The Mac profile's
 *  Axis button opens the Axis dialog, and its group titles are the desktop's Mac group labels in
 *  `lang`; elsewhere Axis opens the Axis menu. */
export const buildSparklineDesignTab = (
  tr: ToolbarText,
  profile: RibbonProfile = 'default',
  lang: ToolbarLang = 'en',
): RibbonTabModel => {
  const t = tr.sparkline;
  const mac = profile === 'excel365Mac';
  /** The group title: the Mac label of `macKey` on the Mac profile, else `fallback`. */
  const title = (macKey: string, fallback: string): string =>
    mac ? macGroupTitle(macKey, lang) : fallback;
  const large = (id: string, label: string, icon: string): RibbonCommand => ({
    id,
    title: label,
    label,
    icon,
    kind: 'large',
  });
  const stacked = (id: string, label: string, icon: string, className?: string): RibbonCommand => ({
    id,
    title: label,
    label,
    icon,
    kind: 'wide',
    layout: 'stacked',
    ...(className ? { className } : {}),
  });
  const check = (id: string, label: string, icon: string): RibbonCommand =>
    stacked(id, label, icon, 'fc-tb__rb--mac-checkbox');
  const group = (
    title: string,
    commands: RibbonCommand[],
    variant = 'tiles',
  ): RibbonGroupModel => ({ title, commands, variant });
  return {
    id: 'sparklineDesign',
    label: tr.tabs.sparklineDesign,
    contextual: true,
    groups: [
      group(title('Sparkline', t.title), [
        large(SPARKLINE_COMMANDS.editData, t.editData, 'chartSelectData'),
      ]),
      group(title('Type (Sparkline)', t.type), [
        large('sparklineTypeLine', t.line, 'sparkline'),
        large('sparklineTypeColumn', t.column, 'sparklineColumn'),
        large('sparklineTypeWinLoss', t.winLoss, 'sparklineWinLoss'),
      ]),
      group(
        title('Show (Sparkline)', t.show),
        [
          check('sparklineShowHigh', t.highPoint, 'sparklineHigh'),
          check('sparklineShowLow', t.lowPoint, 'sparklineLow'),
          check('sparklineShowNegative', t.negativePoints, 'sparklineNegative'),
          check('sparklineShowFirst', t.firstPoint, 'sparklineFirst'),
          check('sparklineShowLast', t.lastPoint, 'sparklineLast'),
          check('sparklineShowMarkers', t.markers, 'sparklineMarkers'),
        ],
        SPARKLINE_STACKED_LAYOUT_GROUP_VARIANTS[0],
      ),
      group(title('Style', t.style), [
        large(SPARKLINE_COMMANDS.styleGallery, t.style, 'chartStyles'),
        large(SPARKLINE_COMMANDS.color, t.color, 'shapeOutline'),
        large(SPARKLINE_COMMANDS.markerColor, t.markerColor, 'fillColor'),
      ]),
      group(
        title('Group (Sparkline)', t.group),
        [
          large(
            mac ? SPARKLINE_COMMANDS.axisDialog : SPARKLINE_COMMANDS.axis,
            t.axis,
            'sparklineAxis',
          ),
          stacked(SPARKLINE_COMMANDS.group, t.groupButton, 'outlineGroup'),
          stacked(SPARKLINE_COMMANDS.ungroup, t.ungroup, 'outlineUngroup'),
          stacked(SPARKLINE_COMMANDS.clear, t.clear, 'clearAll'),
        ],
        SPARKLINE_MIXED_LAYOUT_GROUP_VARIANTS[0],
      ),
    ],
  };
};
