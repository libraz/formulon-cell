import { commentAt } from '../commands/comment.js';
import type { BorderPreset as CommandBorderPreset } from '../commands/format.js';
import { tableForCell } from '../commands/format-as-table.js';
import type { Range } from '../engine/types.js';
import { mergeAt } from '../commands/merge.js';
import type { MarginPreset } from '../commands/page-setup.js';
import { marginPresetOf, pageSetupForSheet } from '../commands/page-setup.js';
import { hiddenInSelection } from '../commands/structure.js';
import type { Strings } from '../i18n/strings/types.js';
import type { SpreadsheetInstance } from '../mount.js';
import type {
  CellBorderStyle,
  ConditionalRule,
  NumFmt,
  PageOrientation,
  PaperSize,
  WorkbookViewMode,
} from '../store/types.js';

/**
 * Ribbon active-state contract shared by every framework wrapper.
 *
 * Both `@libraz/formulon-cell-react` and `@libraz/formulon-cell-vue`
 * derive their ribbon button states (bold pressed?, current font, page
 * orientation, …) from the same projection over the spreadsheet store,
 * so the canonical shape, defaults, and projection live here in core.
 * Wrappers re-export these so consumers keep one source of truth even
 * when mixing frameworks in the same app.
 */
export interface ActiveState {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  alignLeft: boolean;
  alignCenter: boolean;
  alignRight: boolean;
  vAlignTop: boolean;
  vAlignMiddle: boolean;
  vAlignBottom: boolean;
  wrapText: boolean;
  merged: boolean;
  mergeCenter: boolean;
  conditionalFormatting: boolean;
  formatAsTable: boolean;
  cellStyle: string | null;
  textOrientation:
    | 'angleCounterclockwise'
    | 'angleClockwise'
    | 'rotateTextUp'
    | 'rotateTextDown'
    | 'horizontalText';
  currency: boolean;
  percent: boolean;
  commaStyle: boolean;
  numberFormat: string;
  frozen: boolean;
  filterOn: boolean;
  rowsHidden: boolean;
  colsHidden: boolean;
  protected: boolean;
  zoom: number;
  formatPainterArmed: boolean;
  hasComment: boolean;
  fontFamily: string;
  fontSize: number;
  fontColor: string;
  fillColor: string;
  pageOrientation: PageOrientation;
  paperSize: PaperSize;
  pageScale: number;
  fitWidth: number | null;
  fitHeight: number | null;
  gridlinesVisible: boolean;
  headingsVisible: boolean;
  printGridlines: boolean;
  printHeadings: boolean;
  formulasVisible: boolean;
  workbookView: WorkbookViewMode;
  r1c1: boolean;
  calcMode: 0 | 1 | 2 | null;
  /** Closest named preset for the active sheet's margins, or `null`
   *  when the user has set custom values via the Page Setup dialog. */
  marginPreset: MarginPreset | null;
}

export const EMPTY_ACTIVE_STATE: ActiveState = {
  bold: false,
  italic: false,
  underline: false,
  strike: false,
  alignLeft: false,
  alignCenter: false,
  alignRight: false,
  vAlignTop: false,
  vAlignMiddle: false,
  vAlignBottom: true,
  wrapText: false,
  merged: false,
  mergeCenter: false,
  conditionalFormatting: false,
  formatAsTable: false,
  cellStyle: null,
  textOrientation: 'horizontalText',
  currency: false,
  percent: false,
  commaStyle: false,
  numberFormat: 'general',
  frozen: false,
  filterOn: false,
  rowsHidden: false,
  colsHidden: false,
  protected: false,
  zoom: 1,
  formatPainterArmed: false,
  hasComment: false,
  fontFamily: 'Aptos',
  fontSize: 11,
  fontColor: '#201f1e',
  fillColor: '#ffffff',
  pageOrientation: 'portrait',
  paperSize: 'A4',
  pageScale: 1,
  fitWidth: null,
  fitHeight: null,
  gridlinesVisible: true,
  headingsVisible: true,
  printGridlines: false,
  printHeadings: false,
  formulasVisible: false,
  workbookView: 'normal',
  r1c1: false,
  calcMode: null,
  marginPreset: 'normal',
};

const numberFormatOf = (fmt: NumFmt | undefined): string => {
  if (!fmt) return 'general';
  switch (fmt.kind) {
    case 'fixed':
      return 'fixed';
    case 'currency':
      return 'currency';
    case 'accounting':
      return 'accounting';
    case 'date':
      return fmt.pattern.includes('mmmm') || fmt.pattern.includes('"年"')
        ? 'longDate'
        : 'shortDate';
    case 'time':
      return 'time';
    case 'percent':
      return 'percent';
    case 'custom':
      return fmt.pattern.includes('?/?') ? 'fraction' : 'general';
    case 'scientific':
      return 'scientific';
    case 'text':
      return 'text';
    case 'datetime':
      return 'shortDate';
    default:
      return 'general';
  }
};

const rangesIntersect = (a: Range, b: ConditionalRule['range']): boolean =>
  a.sheet === b.sheet && !(a.r1 < b.r0 || a.r0 > b.r1 || a.c1 < b.c0 || a.c0 > b.c1);

type BorderPresetLabelKey =
  | 'noBorder'
  | 'outsideBorders'
  | 'thickOutsideBorders'
  | 'allBorders'
  | 'insideBorders'
  | 'insideHorizontalBorder'
  | 'insideVerticalBorder'
  | 'topBorder'
  | 'bottomBorder'
  | 'leftBorder'
  | 'rightBorder'
  | 'doubleBottomBorder'
  | 'thickBottomBorder'
  | 'topAndBottomBorder'
  | 'topAndThickBottomBorder'
  | 'topAndDoubleBottomBorder'
  | 'diagonalDownBorder'
  | 'diagonalUpBorder';

type BorderStyleLabelKey =
  | 'thin'
  | 'medium'
  | 'thick'
  | 'dashed'
  | 'dotted'
  | 'double'
  | 'hair'
  | 'mediumDashed'
  | 'dashDot'
  | 'mediumDashDot'
  | 'dashDotDot'
  | 'mediumDashDotDot'
  | 'slantDashDot';

export const BORDER_PRESETS: {
  value: CommandBorderPreset;
  label: string;
  labelKey: BorderPresetLabelKey;
}[] = [
  { value: 'none', label: 'No Border', labelKey: 'noBorder' },
  { value: 'outline', label: 'Outside Borders', labelKey: 'outsideBorders' },
  { value: 'thickOutline', label: 'Thick Outside Borders', labelKey: 'thickOutsideBorders' },
  { value: 'all', label: 'All Borders', labelKey: 'allBorders' },
  { value: 'inside', label: 'Inside Borders', labelKey: 'insideBorders' },
  {
    value: 'insideHorizontal',
    label: 'Inside Horizontal Border',
    labelKey: 'insideHorizontalBorder',
  },
  { value: 'insideVertical', label: 'Inside Vertical Border', labelKey: 'insideVerticalBorder' },
  { value: 'top', label: 'Top Border', labelKey: 'topBorder' },
  { value: 'bottom', label: 'Bottom Border', labelKey: 'bottomBorder' },
  { value: 'left', label: 'Left Border', labelKey: 'leftBorder' },
  { value: 'right', label: 'Right Border', labelKey: 'rightBorder' },
  { value: 'doubleBottom', label: 'Double Bottom', labelKey: 'doubleBottomBorder' },
  { value: 'thickBottom', label: 'Thick Bottom Border', labelKey: 'thickBottomBorder' },
  { value: 'topAndBottom', label: 'Top and Bottom Border', labelKey: 'topAndBottomBorder' },
  {
    value: 'topAndThickBottom',
    label: 'Top and Thick Bottom Border',
    labelKey: 'topAndThickBottomBorder',
  },
  {
    value: 'topAndDoubleBottom',
    label: 'Top and Bottom Double Border',
    labelKey: 'topAndDoubleBottomBorder',
  },
  { value: 'diagonalDown', label: 'Diagonal Down Border', labelKey: 'diagonalDownBorder' },
  { value: 'diagonalUp', label: 'Diagonal Up Border', labelKey: 'diagonalUpBorder' },
];

export const BORDER_STYLES: {
  value: CellBorderStyle;
  label: string;
  labelKey: BorderStyleLabelKey;
}[] = [
  { value: 'thin', label: 'Thin', labelKey: 'thin' },
  { value: 'medium', label: 'Medium', labelKey: 'medium' },
  { value: 'thick', label: 'Thick', labelKey: 'thick' },
  { value: 'dashed', label: 'Dashed', labelKey: 'dashed' },
  { value: 'dotted', label: 'Dotted', labelKey: 'dotted' },
  { value: 'double', label: 'Double', labelKey: 'double' },
  { value: 'hair', label: 'Hairline', labelKey: 'hair' },
  { value: 'mediumDashed', label: 'Medium Dashed', labelKey: 'mediumDashed' },
  { value: 'dashDot', label: 'Dash Dot', labelKey: 'dashDot' },
  { value: 'mediumDashDot', label: 'Medium Dash Dot', labelKey: 'mediumDashDot' },
  { value: 'dashDotDot', label: 'Dash Dot Dot', labelKey: 'dashDotDot' },
  { value: 'mediumDashDotDot', label: 'Medium Dash Dot Dot', labelKey: 'mediumDashDotDot' },
  { value: 'slantDashDot', label: 'Slant Dash Dot', labelKey: 'slantDashDot' },
];

export const localizeBorderPresets = (
  ribbon: Strings['ribbon'],
): { value: CommandBorderPreset; label: string }[] =>
  BORDER_PRESETS.map((preset) => ({ value: preset.value, label: ribbon[preset.labelKey] }));

export const localizeBorderStyles = (
  ribbon: Strings['ribbon'],
): { value: CellBorderStyle; label: string }[] =>
  BORDER_STYLES.map((style) => ({ value: style.value, label: ribbon[style.labelKey] }));

export const projectActiveState = (inst: SpreadsheetInstance): ActiveState => {
  const s = inst.store.getState();
  const a = s.selection.active;
  const r = s.selection.range;
  const activeMerge = mergeAt(s, a);
  const formatAddr = activeMerge
    ? { sheet: activeMerge.sheet, row: activeMerge.r0, col: activeMerge.c0 }
    : a;
  const f = s.format.formats.get(`${formatAddr.sheet}:${formatAddr.row}:${formatAddr.col}`);
  const hasConditionalFormatting = s.conditional.rules.some((rule) =>
    rangesIntersect(r, rule.range),
  );
  const activeTable = tableForCell(s.tables.tables, a.sheet, a.row, a.col);
  const setup = pageSetupForSheet(s, s.data.sheetIndex);
  return {
    bold: !!f?.bold,
    italic: !!f?.italic,
    underline: !!f?.underline,
    strike: !!f?.strike,
    alignLeft: f?.align === 'left',
    alignCenter: f?.align === 'center',
    alignRight: f?.align === 'right',
    vAlignTop: f?.vAlign === 'top',
    vAlignMiddle: f?.vAlign === 'middle',
    vAlignBottom: f?.vAlign == null || f.vAlign === 'bottom',
    wrapText: f?.wrap === true,
    merged: activeMerge != null,
    mergeCenter: activeMerge != null && f?.align === 'center',
    conditionalFormatting: hasConditionalFormatting,
    formatAsTable: activeTable != null,
    cellStyle: f?.cellStyle ?? null,
    textOrientation:
      f?.rotation === 45
        ? 'angleCounterclockwise'
        : f?.rotation === -45
          ? 'angleClockwise'
          : f?.rotation === 90
            ? 'rotateTextUp'
            : f?.rotation === -90
              ? 'rotateTextDown'
              : 'horizontalText',
    currency: f?.numFmt?.kind === 'currency',
    percent: f?.numFmt?.kind === 'percent',
    commaStyle: f?.numFmt?.kind === 'fixed' && f.numFmt.thousands === true,
    numberFormat: numberFormatOf(f?.numFmt),
    frozen: s.layout.freezeRows > 0 || s.layout.freezeCols > 0,
    filterOn: s.ui.filterRange != null,
    rowsHidden: hiddenInSelection(s.layout, 'row', r.r0, r.r1).length > 0,
    colsHidden: hiddenInSelection(s.layout, 'col', r.c0, r.c1).length > 0,
    protected: inst.isSheetProtected(),
    zoom: s.viewport.zoom,
    formatPainterArmed: !!inst.formatPainter?.isActive(),
    hasComment: commentAt(s, a) != null,
    fontFamily: f?.fontFamily ?? 'Aptos',
    fontSize: f?.fontSize ?? 11,
    fontColor: f?.color ?? '#201f1e',
    fillColor: f?.fill ?? '#ffffff',
    pageOrientation: setup.orientation,
    paperSize: setup.paperSize,
    pageScale: setup.scale ?? 1,
    fitWidth: setup.fitWidth ?? null,
    fitHeight: setup.fitHeight ?? null,
    gridlinesVisible: s.ui.showGridLines !== false,
    headingsVisible: s.ui.showHeaders !== false,
    printGridlines: setup.showGridlines === true,
    printHeadings: setup.showHeadings === true,
    formulasVisible: !!s.ui.showFormulas,
    workbookView: s.ui.workbookView,
    r1c1: !!s.ui.r1c1,
    calcMode: inst.workbook.calcMode(),
    marginPreset: marginPresetOf(setup.margins),
  };
};
