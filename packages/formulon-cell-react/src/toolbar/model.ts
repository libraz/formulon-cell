import {
  type CellBorderStyle,
  commentAt,
  FONT_FAMILIES,
  FONT_SIZES,
  hiddenInSelection,
  type MarginPreset,
  marginPresetOf,
  type PageOrientation,
  type PaperSize,
  pageSetupForSheet,
  RIBBON_TAB_LABELS,
  type RibbonTab,
  type SpreadsheetInstance,
} from '@libraz/formulon-cell';

export interface SpreadsheetToolbarProps {
  instance: SpreadsheetInstance | null;
  activeTab: RibbonTab;
  onTabChange: (tab: RibbonTab) => void;
  locale: string;
}

export { RIBBON_TAB_LABELS, type RibbonTab };

export interface ActiveState {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  alignLeft: boolean;
  alignCenter: boolean;
  alignRight: boolean;
  currency: boolean;
  percent: boolean;
  frozen: boolean;
  filterOn: boolean;
  rowsHidden: boolean;
  colsHidden: boolean;
  protected: boolean;
  zoom: number;
  fontFamily: string;
  fontSize: number;
  fontColor: string;
  fillColor: string;
  formatPainterArmed: boolean;
  hasComment: boolean;
  pageOrientation: PageOrientation;
  paperSize: PaperSize;
  /** Closest named preset for the active sheet's margins, or `null` when
   *  the user has set custom values via the Page Setup dialog. */
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
  currency: false,
  percent: false,
  frozen: false,
  filterOn: false,
  rowsHidden: false,
  colsHidden: false,
  protected: false,
  zoom: 1,
  fontFamily: 'Aptos',
  fontSize: 11,
  fontColor: '#201f1e',
  fillColor: '#ffffff',
  formatPainterArmed: false,
  hasComment: false,
  pageOrientation: 'portrait',
  paperSize: 'A4',
  marginPreset: 'normal',
};

export { FONT_FAMILIES, FONT_SIZES };

export type BorderPreset =
  | 'none'
  | 'outline'
  | 'all'
  | 'top'
  | 'bottom'
  | 'left'
  | 'right'
  | 'doubleBottom';

export const BORDER_STYLES: { value: CellBorderStyle; label: string }[] = [
  { value: 'thin', label: 'Thin' },
  { value: 'medium', label: 'Medium' },
  { value: 'thick', label: 'Thick' },
  { value: 'dashed', label: 'Dashed' },
  { value: 'dotted', label: 'Dotted' },
  { value: 'double', label: 'Double' },
];

export const BORDER_PRESETS: { value: BorderPreset; label: string }[] = [
  { value: 'none', label: 'No Border' },
  { value: 'outline', label: 'Outside Borders' },
  { value: 'all', label: 'All Borders' },
  { value: 'top', label: 'Top Border' },
  { value: 'bottom', label: 'Bottom Border' },
  { value: 'left', label: 'Left Border' },
  { value: 'right', label: 'Right Border' },
  { value: 'doubleBottom', label: 'Double Bottom' },
];

export const projectActiveState = (inst: SpreadsheetInstance): ActiveState => {
  const s = inst.store.getState();
  const a = s.selection.active;
  const r = s.selection.range;
  const f = s.format.formats.get(`${a.sheet}:${a.row}:${a.col}`);
  const setup = pageSetupForSheet(s, s.data.sheetIndex);
  return {
    bold: !!f?.bold,
    italic: !!f?.italic,
    underline: !!f?.underline,
    strike: !!f?.strike,
    alignLeft: f?.align === 'left',
    alignCenter: f?.align === 'center',
    alignRight: f?.align === 'right',
    currency: f?.numFmt?.kind === 'currency',
    percent: f?.numFmt?.kind === 'percent',
    frozen: s.layout.freezeRows > 0 || s.layout.freezeCols > 0,
    filterOn: s.ui.filterRange != null,
    rowsHidden: hiddenInSelection(s.layout, 'row', r.r0, r.r1).length > 0,
    colsHidden: hiddenInSelection(s.layout, 'col', r.c0, r.c1).length > 0,
    protected: inst.isSheetProtected(),
    zoom: s.viewport.zoom,
    fontFamily: f?.fontFamily ?? 'Aptos',
    fontSize: f?.fontSize ?? 11,
    fontColor: f?.color ?? '#201f1e',
    fillColor: f?.fill ?? '#ffffff',
    formatPainterArmed: !!inst.formatPainter?.isActive(),
    hasComment: commentAt(s, a) != null,
    pageOrientation: setup.orientation,
    paperSize: setup.paperSize,
    marginPreset: marginPresetOf(setup.margins),
  };
};
