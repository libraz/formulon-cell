import { commentAt } from '../commands/comment.js';
import type { BorderPreset as CommandBorderPreset } from '../commands/format.js';
import type { MarginPreset } from '../commands/page-setup.js';
import { marginPresetOf, pageSetupForSheet } from '../commands/page-setup.js';
import { hiddenInSelection } from '../commands/structure.js';
import type { SpreadsheetInstance } from '../mount.js';
import type { CellBorderStyle, PageOrientation, PaperSize } from '../store/types.js';

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
  currency: boolean;
  percent: boolean;
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
  currency: false,
  percent: false,
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
  marginPreset: 'normal',
};

export const BORDER_PRESETS: { value: CommandBorderPreset; label: string }[] = [
  { value: 'none', label: 'No Border' },
  { value: 'outline', label: 'Outside Borders' },
  { value: 'all', label: 'All Borders' },
  { value: 'top', label: 'Top Border' },
  { value: 'bottom', label: 'Bottom Border' },
  { value: 'left', label: 'Left Border' },
  { value: 'right', label: 'Right Border' },
  { value: 'doubleBottom', label: 'Double Bottom' },
];

export const BORDER_STYLES: { value: CellBorderStyle; label: string }[] = [
  { value: 'thin', label: 'Thin' },
  { value: 'medium', label: 'Medium' },
  { value: 'thick', label: 'Thick' },
  { value: 'dashed', label: 'Dashed' },
  { value: 'dotted', label: 'Dotted' },
  { value: 'double', label: 'Double' },
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
    formatPainterArmed: !!inst.formatPainter?.isActive(),
    hasComment: commentAt(s, a) != null,
    fontFamily: f?.fontFamily ?? 'Aptos',
    fontSize: f?.fontSize ?? 11,
    fontColor: f?.color ?? '#201f1e',
    fillColor: f?.fill ?? '#ffffff',
    pageOrientation: setup.orientation,
    paperSize: setup.paperSize,
    marginPreset: marginPresetOf(setup.margins),
  };
};
