import { DEFAULT_THEME_ACCENTS } from '../charts/series-colors.js';
import type { CellBorderStyle } from '../store/types/cell-format.js';
import type { StyleElementFormat, StyleElementFormats } from './table-style-elements.js';
import { cssColorFromArgb, tintCssColor } from './theme-color.js';

/**
 * The elements of the 60 built-in table styles (TableStyleLight1-21,
 * TableStyleMedium1-28, TableStyleDark1-11), as the desktop app defines them,
 * with colours kept as theme references so they follow the workbook theme.
 * Read from the desktop app's object model; its whole-table element is not
 * readable there, so whole-table borders (e.g. the outer rules of Light1, the
 * grid of Light15-21) are not part of these definitions.
 */

/** A theme colour: its clrScheme slot (0 dk1, 1 lt1, 4-9 accent1-6) and tint. */
type ThemeRef = readonly [slot: number, tint: number];
type Side = readonly [style: CellBorderStyle, color: ThemeRef];

interface ElementDef {
  fill?: ThemeRef;
  color?: ThemeRef;
  bold?: true;
  top?: Side;
  bottom?: Side;
  left?: Side;
  right?: Side;
}

type StyleDef = Partial<Record<keyof StyleElementFormats, ElementDef>>;

// The tints the built-in styles use, as the file stores them.
const L80 = 0.7999816888943144;
const L60 = 0.5999938962981048;
const L25 = 0.249977111117893;
const L15 = 0.1499984740745262;
const D15 = -0.1499984740745262;
const D25 = -0.249977111117893;
const D35 = -0.3499862666707358;
const D50 = -0.499984740745262;

const STYLES: Readonly<Record<string, StyleDef>> = {
  TableStyleLight1: {
    headerRow: { color: [0, 0], bold: true, bottom: ['thin', [0, 0]] },
    totalRow: { color: [0, 0], bold: true, top: ['thin', [0, 0]] },
    firstColumn: { color: [0, 0], bold: true },
    lastColumn: { color: [0, 0], bold: true },
    firstRowStripe: { fill: [1, D15] },
    firstColumnStripe: { fill: [1, D15] },
  },
  TableStyleLight2: {
    headerRow: { color: [4, D25], bold: true, bottom: ['thin', [4, 0]] },
    totalRow: { color: [4, D25], bold: true, top: ['thin', [4, 0]] },
    firstColumn: { color: [4, D25], bold: true },
    lastColumn: { color: [4, D25], bold: true },
    firstRowStripe: { fill: [4, L80] },
    firstColumnStripe: { fill: [4, L80] },
  },
  TableStyleLight3: {
    headerRow: { color: [5, D25], bold: true, bottom: ['thin', [5, 0]] },
    totalRow: { color: [5, D25], bold: true, top: ['thin', [5, 0]] },
    firstColumn: { color: [5, D25], bold: true },
    lastColumn: { color: [5, D25], bold: true },
    firstRowStripe: { fill: [5, L80] },
    firstColumnStripe: { fill: [5, L80] },
  },
  TableStyleLight4: {
    headerRow: { color: [6, D25], bold: true, bottom: ['thin', [6, 0]] },
    totalRow: { color: [6, D25], bold: true, top: ['thin', [6, 0]] },
    firstColumn: { color: [6, D25], bold: true },
    lastColumn: { color: [6, D25], bold: true },
    firstRowStripe: { fill: [6, L80] },
    firstColumnStripe: { fill: [6, L80] },
  },
  TableStyleLight5: {
    headerRow: { color: [7, D25], bold: true, bottom: ['thin', [7, 0]] },
    totalRow: { color: [7, D25], bold: true, top: ['thin', [7, 0]] },
    firstColumn: { color: [7, D25], bold: true },
    lastColumn: { color: [7, D25], bold: true },
    firstRowStripe: { fill: [7, L80] },
    firstColumnStripe: { fill: [7, L80] },
  },
  TableStyleLight6: {
    headerRow: { color: [8, D25], bold: true, bottom: ['thin', [8, 0]] },
    totalRow: { color: [8, D25], bold: true, top: ['thin', [8, 0]] },
    firstColumn: { color: [8, D25], bold: true },
    lastColumn: { color: [8, D25], bold: true },
    firstRowStripe: { fill: [8, L80] },
    firstColumnStripe: { fill: [8, L80] },
  },
  TableStyleLight7: {
    headerRow: { color: [9, D25], bold: true, bottom: ['thin', [9, 0]] },
    totalRow: { color: [9, D25], bold: true, top: ['thin', [9, 0]] },
    firstColumn: { color: [9, D25], bold: true },
    lastColumn: { color: [9, D25], bold: true },
    firstRowStripe: { fill: [9, L80] },
    firstColumnStripe: { fill: [9, L80] },
  },
  TableStyleLight8: {
    headerRow: { fill: [0, 0], color: [1, 0], bold: true },
    totalRow: { color: [0, 0], bold: true, top: ['double', [0, 0]] },
    firstColumn: { color: [0, 0], bold: true },
    lastColumn: { color: [0, 0], bold: true },
    firstRowStripe: { top: ['thin', [0, 0]] },
    secondRowStripe: { top: ['thin', [0, 0]] },
    firstColumnStripe: { left: ['thin', [0, 0]] },
    secondColumnStripe: { left: ['thin', [0, 0]] },
  },
  TableStyleLight9: {
    headerRow: { fill: [4, 0], color: [1, 0], bold: true },
    totalRow: { color: [0, 0], bold: true, top: ['double', [4, 0]] },
    firstColumn: { color: [0, 0], bold: true },
    lastColumn: { color: [0, 0], bold: true },
    firstRowStripe: { top: ['thin', [4, 0]] },
    secondRowStripe: { top: ['thin', [4, 0]] },
    firstColumnStripe: { left: ['thin', [4, 0]] },
    secondColumnStripe: { left: ['thin', [4, 0]] },
  },
  TableStyleLight10: {
    headerRow: { fill: [5, 0], color: [1, 0], bold: true },
    totalRow: { color: [0, 0], bold: true, top: ['double', [5, 0]] },
    firstColumn: { color: [0, 0], bold: true },
    lastColumn: { color: [0, 0], bold: true },
    firstRowStripe: { top: ['thin', [5, 0]] },
    secondRowStripe: { top: ['thin', [5, 0]] },
    firstColumnStripe: { left: ['thin', [5, 0]] },
    secondColumnStripe: { left: ['thin', [5, 0]] },
  },
  TableStyleLight11: {
    headerRow: { fill: [6, 0], color: [1, 0], bold: true },
    totalRow: { color: [0, 0], bold: true, top: ['double', [6, 0]] },
    firstColumn: { color: [0, 0], bold: true },
    lastColumn: { color: [0, 0], bold: true },
    firstRowStripe: { top: ['thin', [6, 0]] },
    secondRowStripe: { top: ['thin', [6, 0]] },
    firstColumnStripe: { left: ['thin', [6, 0]] },
    secondColumnStripe: { left: ['thin', [6, 0]] },
  },
  TableStyleLight12: {
    headerRow: { fill: [7, 0], color: [1, 0], bold: true },
    totalRow: { color: [0, 0], bold: true, top: ['double', [7, 0]] },
    firstColumn: { color: [0, 0], bold: true },
    lastColumn: { color: [0, 0], bold: true },
    firstRowStripe: { top: ['thin', [7, 0]] },
    secondRowStripe: { top: ['thin', [7, 0]] },
    firstColumnStripe: { left: ['thin', [7, 0]] },
    secondColumnStripe: { left: ['thin', [7, 0]] },
  },
  TableStyleLight13: {
    headerRow: { fill: [8, 0], color: [1, 0], bold: true },
    totalRow: { color: [0, 0], bold: true, top: ['double', [8, 0]] },
    firstColumn: { color: [0, 0], bold: true },
    lastColumn: { color: [0, 0], bold: true },
    firstRowStripe: { top: ['thin', [8, 0]] },
    secondRowStripe: { top: ['thin', [8, 0]] },
    firstColumnStripe: { left: ['thin', [8, 0]] },
    secondColumnStripe: { left: ['thin', [8, 0]] },
  },
  TableStyleLight14: {
    headerRow: { fill: [9, 0], color: [1, 0], bold: true },
    totalRow: { color: [0, 0], bold: true, top: ['double', [9, 0]] },
    firstColumn: { color: [0, 0], bold: true },
    lastColumn: { color: [0, 0], bold: true },
    firstRowStripe: { top: ['thin', [9, 0]] },
    secondRowStripe: { top: ['thin', [9, 0]] },
    firstColumnStripe: { left: ['thin', [9, 0]] },
    secondColumnStripe: { left: ['thin', [9, 0]] },
  },
  TableStyleLight15: {
    headerRow: { color: [0, 0], bold: true, bottom: ['medium', [0, 0]] },
    totalRow: { color: [0, 0], bold: true, top: ['double', [0, 0]] },
    firstColumn: { color: [0, 0], bold: true },
    lastColumn: { color: [0, 0], bold: true },
    firstRowStripe: { fill: [1, D15] },
    firstColumnStripe: { fill: [1, D15] },
  },
  TableStyleLight16: {
    headerRow: { color: [0, 0], bold: true, bottom: ['medium', [4, 0]] },
    totalRow: { color: [0, 0], bold: true, top: ['double', [4, 0]] },
    firstColumn: { color: [0, 0], bold: true },
    lastColumn: { color: [0, 0], bold: true },
    firstRowStripe: { fill: [4, L80] },
    firstColumnStripe: { fill: [4, L80] },
  },
  TableStyleLight17: {
    headerRow: { color: [0, 0], bold: true, bottom: ['medium', [5, 0]] },
    totalRow: { color: [0, 0], bold: true, top: ['double', [5, 0]] },
    firstColumn: { color: [0, 0], bold: true },
    lastColumn: { color: [0, 0], bold: true },
    firstRowStripe: { fill: [5, L80] },
    firstColumnStripe: { fill: [5, L80] },
  },
  TableStyleLight18: {
    headerRow: { color: [0, 0], bold: true, bottom: ['medium', [6, 0]] },
    totalRow: { color: [0, 0], bold: true, top: ['double', [6, 0]] },
    firstColumn: { color: [0, 0], bold: true },
    lastColumn: { color: [0, 0], bold: true },
    firstRowStripe: { fill: [6, L80] },
    firstColumnStripe: { fill: [6, L80] },
  },
  TableStyleLight19: {
    headerRow: { color: [0, 0], bold: true, bottom: ['medium', [7, 0]] },
    totalRow: { color: [0, 0], bold: true, top: ['double', [7, 0]] },
    firstColumn: { color: [0, 0], bold: true },
    lastColumn: { color: [0, 0], bold: true },
    firstRowStripe: { fill: [7, L80] },
    firstColumnStripe: { fill: [7, L80] },
  },
  TableStyleLight20: {
    headerRow: { color: [0, 0], bold: true, bottom: ['medium', [8, 0]] },
    totalRow: { color: [0, 0], bold: true, top: ['double', [8, 0]] },
    firstColumn: { color: [0, 0], bold: true },
    lastColumn: { color: [0, 0], bold: true },
    firstRowStripe: { fill: [8, L80] },
    firstColumnStripe: { fill: [8, L80] },
  },
  TableStyleLight21: {
    headerRow: { color: [0, 0], bold: true, bottom: ['medium', [9, 0]] },
    totalRow: { color: [0, 0], bold: true, top: ['double', [9, 0]] },
    firstColumn: { color: [0, 0], bold: true },
    lastColumn: { color: [0, 0], bold: true },
    firstRowStripe: { fill: [9, L80] },
    firstColumnStripe: { fill: [9, L80] },
  },
  TableStyleMedium1: {
    headerRow: { fill: [0, 0], color: [1, 0], bold: true },
    totalRow: { color: [0, 0], bold: true, top: ['double', [0, 0]] },
    firstColumn: { color: [0, 0], bold: true },
    lastColumn: { color: [0, 0], bold: true },
    firstRowStripe: { fill: [1, D15] },
    firstColumnStripe: { fill: [1, D15] },
  },
  TableStyleMedium2: {
    headerRow: { fill: [4, 0], color: [1, 0], bold: true },
    totalRow: { color: [0, 0], bold: true, top: ['double', [4, 0]] },
    firstColumn: { color: [0, 0], bold: true },
    lastColumn: { color: [0, 0], bold: true },
    firstRowStripe: { fill: [4, L80] },
    firstColumnStripe: { fill: [4, L80] },
  },
  TableStyleMedium3: {
    headerRow: { fill: [5, 0], color: [1, 0], bold: true },
    totalRow: { color: [0, 0], bold: true, top: ['double', [5, 0]] },
    firstColumn: { color: [0, 0], bold: true },
    lastColumn: { color: [0, 0], bold: true },
    firstRowStripe: { fill: [5, L80] },
    firstColumnStripe: { fill: [5, L80] },
  },
  TableStyleMedium4: {
    headerRow: { fill: [6, 0], color: [1, 0], bold: true },
    totalRow: { color: [0, 0], bold: true, top: ['double', [6, 0]] },
    firstColumn: { color: [0, 0], bold: true },
    lastColumn: { color: [0, 0], bold: true },
    firstRowStripe: { fill: [6, L80] },
    firstColumnStripe: { fill: [6, L80] },
  },
  TableStyleMedium5: {
    headerRow: { fill: [7, 0], color: [1, 0], bold: true },
    totalRow: { color: [0, 0], bold: true, top: ['double', [7, 0]] },
    firstColumn: { color: [0, 0], bold: true },
    lastColumn: { color: [0, 0], bold: true },
    firstRowStripe: { fill: [7, L80] },
    firstColumnStripe: { fill: [7, L80] },
  },
  TableStyleMedium6: {
    headerRow: { fill: [8, 0], color: [1, 0], bold: true },
    totalRow: { color: [0, 0], bold: true, top: ['double', [8, 0]] },
    firstColumn: { color: [0, 0], bold: true },
    lastColumn: { color: [0, 0], bold: true },
    firstRowStripe: { fill: [8, L80] },
    firstColumnStripe: { fill: [8, L80] },
  },
  TableStyleMedium7: {
    headerRow: { fill: [9, 0], color: [1, 0], bold: true },
    totalRow: { color: [0, 0], bold: true, top: ['double', [9, 0]] },
    firstColumn: { color: [0, 0], bold: true },
    lastColumn: { color: [0, 0], bold: true },
    firstRowStripe: { fill: [9, L80] },
    firstColumnStripe: { fill: [9, L80] },
  },
  TableStyleMedium8: {
    headerRow: { fill: [0, 0], color: [1, 0], bold: true, bottom: ['thick', [1, 0]] },
    totalRow: { fill: [0, 0], color: [1, 0], bold: true, top: ['thick', [1, 0]] },
    firstColumn: { fill: [0, 0], color: [1, 0], bold: true },
    lastColumn: { fill: [0, 0], color: [1, 0], bold: true },
    firstRowStripe: { fill: [1, D35] },
    firstColumnStripe: { fill: [1, D35] },
  },
  TableStyleMedium9: {
    headerRow: { fill: [4, 0], color: [1, 0], bold: true, bottom: ['thick', [1, 0]] },
    totalRow: { fill: [4, 0], color: [1, 0], bold: true, top: ['thick', [1, 0]] },
    firstColumn: { fill: [4, 0], color: [1, 0], bold: true },
    lastColumn: { fill: [4, 0], color: [1, 0], bold: true },
    firstRowStripe: { fill: [4, L60] },
    firstColumnStripe: { fill: [4, L60] },
  },
  TableStyleMedium10: {
    headerRow: { fill: [5, 0], color: [1, 0], bold: true, bottom: ['thick', [1, 0]] },
    totalRow: { fill: [5, 0], color: [1, 0], bold: true, top: ['thick', [1, 0]] },
    firstColumn: { fill: [5, 0], color: [1, 0], bold: true },
    lastColumn: { fill: [5, 0], color: [1, 0], bold: true },
    firstRowStripe: { fill: [5, L60] },
    firstColumnStripe: { fill: [5, L60] },
  },
  TableStyleMedium11: {
    headerRow: { fill: [6, 0], color: [1, 0], bold: true, bottom: ['thick', [1, 0]] },
    totalRow: { fill: [6, 0], color: [1, 0], bold: true, top: ['thick', [1, 0]] },
    firstColumn: { fill: [6, 0], color: [1, 0], bold: true },
    lastColumn: { fill: [6, 0], color: [1, 0], bold: true },
    firstRowStripe: { fill: [6, L60] },
    firstColumnStripe: { fill: [6, L60] },
  },
  TableStyleMedium12: {
    headerRow: { fill: [7, 0], color: [1, 0], bold: true, bottom: ['thick', [1, 0]] },
    totalRow: { fill: [7, 0], color: [1, 0], bold: true, top: ['thick', [1, 0]] },
    firstColumn: { fill: [7, 0], color: [1, 0], bold: true },
    lastColumn: { fill: [7, 0], color: [1, 0], bold: true },
    firstRowStripe: { fill: [7, L60] },
    firstColumnStripe: { fill: [7, L60] },
  },
  TableStyleMedium13: {
    headerRow: { fill: [8, 0], color: [1, 0], bold: true, bottom: ['thick', [1, 0]] },
    totalRow: { fill: [8, 0], color: [1, 0], bold: true, top: ['thick', [1, 0]] },
    firstColumn: { fill: [8, 0], color: [1, 0], bold: true },
    lastColumn: { fill: [8, 0], color: [1, 0], bold: true },
    firstRowStripe: { fill: [8, L60] },
    firstColumnStripe: { fill: [8, L60] },
  },
  TableStyleMedium14: {
    headerRow: { fill: [9, 0], color: [1, 0], bold: true, bottom: ['thick', [1, 0]] },
    totalRow: { fill: [9, 0], color: [1, 0], bold: true, top: ['thick', [1, 0]] },
    firstColumn: { fill: [9, 0], color: [1, 0], bold: true },
    lastColumn: { fill: [9, 0], color: [1, 0], bold: true },
    firstRowStripe: { fill: [9, L60] },
    firstColumnStripe: { fill: [9, L60] },
  },
  TableStyleMedium15: {
    headerRow: { fill: [0, 0], color: [1, 0], bold: true, bottom: ['medium', [0, 0]] },
    totalRow: { top: ['double', [0, 0]] },
    firstColumn: { fill: [0, 0], color: [1, 0], bold: true },
    lastColumn: { fill: [0, 0], color: [1, 0], bold: true },
    firstRowStripe: { fill: [1, D15] },
    firstColumnStripe: { fill: [1, D15] },
  },
  TableStyleMedium16: {
    headerRow: { fill: [4, 0], color: [1, 0], bold: true, bottom: ['medium', [0, 0]] },
    totalRow: { top: ['double', [0, 0]] },
    firstColumn: { fill: [4, 0], color: [1, 0], bold: true },
    lastColumn: { fill: [4, 0], color: [1, 0], bold: true },
    firstRowStripe: { fill: [1, D15] },
    firstColumnStripe: { fill: [1, D15] },
  },
  TableStyleMedium17: {
    headerRow: { fill: [5, 0], color: [1, 0], bold: true, bottom: ['medium', [0, 0]] },
    totalRow: { top: ['double', [0, 0]] },
    firstColumn: { fill: [5, 0], color: [1, 0], bold: true },
    lastColumn: { fill: [5, 0], color: [1, 0], bold: true },
    firstRowStripe: { fill: [1, D15] },
    firstColumnStripe: { fill: [1, D15] },
  },
  TableStyleMedium18: {
    headerRow: { fill: [6, 0], color: [1, 0], bold: true, bottom: ['medium', [0, 0]] },
    totalRow: { top: ['double', [0, 0]] },
    firstColumn: { fill: [6, 0], color: [1, 0], bold: true },
    lastColumn: { fill: [6, 0], color: [1, 0], bold: true },
    firstRowStripe: { fill: [1, D15] },
    firstColumnStripe: { fill: [1, D15] },
  },
  TableStyleMedium19: {
    headerRow: { fill: [7, 0], color: [1, 0], bold: true, bottom: ['medium', [0, 0]] },
    totalRow: { top: ['double', [0, 0]] },
    firstColumn: { fill: [7, 0], color: [1, 0], bold: true },
    lastColumn: { fill: [7, 0], color: [1, 0], bold: true },
    firstRowStripe: { fill: [1, D15] },
    firstColumnStripe: { fill: [1, D15] },
  },
  TableStyleMedium20: {
    headerRow: { fill: [8, 0], color: [1, 0], bold: true, bottom: ['medium', [0, 0]] },
    totalRow: { top: ['double', [0, 0]] },
    firstColumn: { fill: [8, 0], color: [1, 0], bold: true },
    lastColumn: { fill: [8, 0], color: [1, 0], bold: true },
    firstRowStripe: { fill: [1, D15] },
    firstColumnStripe: { fill: [1, D15] },
  },
  TableStyleMedium21: {
    headerRow: { fill: [9, 0], color: [1, 0], bold: true, bottom: ['medium', [0, 0]] },
    totalRow: { top: ['double', [0, 0]] },
    firstColumn: { fill: [9, 0], color: [1, 0], bold: true },
    lastColumn: { fill: [9, 0], color: [1, 0], bold: true },
    firstRowStripe: { fill: [1, D15] },
    firstColumnStripe: { fill: [1, D15] },
  },
  TableStyleMedium22: {
    headerRow: { color: [0, 0], bold: true },
    totalRow: { color: [0, 0], bold: true, top: ['medium', [0, 0]] },
    firstColumn: { color: [0, 0], bold: true },
    lastColumn: { color: [0, 0], bold: true },
    firstRowStripe: { fill: [1, D35] },
    firstColumnStripe: { fill: [1, D35] },
  },
  TableStyleMedium23: {
    headerRow: { color: [0, 0], bold: true },
    totalRow: { color: [0, 0], bold: true, top: ['medium', [4, 0]] },
    firstColumn: { color: [0, 0], bold: true },
    lastColumn: { color: [0, 0], bold: true },
    firstRowStripe: { fill: [4, L60] },
    firstColumnStripe: { fill: [4, L60] },
  },
  TableStyleMedium24: {
    headerRow: { color: [0, 0], bold: true },
    totalRow: { color: [0, 0], bold: true, top: ['medium', [5, 0]] },
    firstColumn: { color: [0, 0], bold: true },
    lastColumn: { color: [0, 0], bold: true },
    firstRowStripe: { fill: [5, L60] },
    firstColumnStripe: { fill: [5, L60] },
  },
  TableStyleMedium25: {
    headerRow: { color: [0, 0], bold: true },
    totalRow: { color: [0, 0], bold: true, top: ['medium', [6, 0]] },
    firstColumn: { color: [0, 0], bold: true },
    lastColumn: { color: [0, 0], bold: true },
    firstRowStripe: { fill: [6, L60] },
    firstColumnStripe: { fill: [6, L60] },
  },
  TableStyleMedium26: {
    headerRow: { color: [0, 0], bold: true },
    totalRow: { color: [0, 0], bold: true, top: ['medium', [7, 0]] },
    firstColumn: { color: [0, 0], bold: true },
    lastColumn: { color: [0, 0], bold: true },
    firstRowStripe: { fill: [7, L60] },
    firstColumnStripe: { fill: [7, L60] },
  },
  TableStyleMedium27: {
    headerRow: { color: [0, 0], bold: true },
    totalRow: { color: [0, 0], bold: true, top: ['medium', [8, 0]] },
    firstColumn: { color: [0, 0], bold: true },
    lastColumn: { color: [0, 0], bold: true },
    firstRowStripe: { fill: [8, L60] },
    firstColumnStripe: { fill: [8, L60] },
  },
  TableStyleMedium28: {
    headerRow: { color: [0, 0], bold: true },
    totalRow: { color: [0, 0], bold: true, top: ['medium', [9, 0]] },
    firstColumn: { color: [0, 0], bold: true },
    lastColumn: { color: [0, 0], bold: true },
    firstRowStripe: { fill: [9, L60] },
    firstColumnStripe: { fill: [9, L60] },
  },
  TableStyleDark1: {
    headerRow: { fill: [0, 0], color: [1, 0], bold: true, bottom: ['medium', [1, 0]] },
    totalRow: { fill: [0, L15], color: [1, 0], bold: true, top: ['medium', [1, 0]] },
    firstColumn: { fill: [0, L25], color: [1, 0], bold: true, right: ['medium', [1, 0]] },
    lastColumn: { fill: [0, L25], color: [1, 0], bold: true, left: ['medium', [1, 0]] },
    firstRowStripe: { fill: [0, L25] },
    firstColumnStripe: { fill: [0, L25] },
  },
  TableStyleDark2: {
    headerRow: { fill: [0, 0], color: [1, 0], bold: true, bottom: ['medium', [1, 0]] },
    totalRow: { fill: [4, D50], color: [1, 0], bold: true, top: ['medium', [1, 0]] },
    firstColumn: { fill: [4, D25], color: [1, 0], bold: true, right: ['medium', [1, 0]] },
    lastColumn: { fill: [4, D25], color: [1, 0], bold: true, left: ['medium', [1, 0]] },
    firstRowStripe: { fill: [4, D25] },
    firstColumnStripe: { fill: [4, D25] },
  },
  TableStyleDark3: {
    headerRow: { fill: [0, 0], color: [1, 0], bold: true, bottom: ['medium', [1, 0]] },
    totalRow: { fill: [5, D50], color: [1, 0], bold: true, top: ['medium', [1, 0]] },
    firstColumn: { fill: [5, D25], color: [1, 0], bold: true, right: ['medium', [1, 0]] },
    lastColumn: { fill: [5, D25], color: [1, 0], bold: true, left: ['medium', [1, 0]] },
    firstRowStripe: { fill: [5, D25] },
    firstColumnStripe: { fill: [5, D25] },
  },
  TableStyleDark4: {
    headerRow: { fill: [0, 0], color: [1, 0], bold: true, bottom: ['medium', [1, 0]] },
    totalRow: { fill: [6, D50], color: [1, 0], bold: true, top: ['medium', [1, 0]] },
    firstColumn: { fill: [6, D25], color: [1, 0], bold: true, right: ['medium', [1, 0]] },
    lastColumn: { fill: [6, D25], color: [1, 0], bold: true, left: ['medium', [1, 0]] },
    firstRowStripe: { fill: [6, D25] },
    firstColumnStripe: { fill: [6, D25] },
  },
  TableStyleDark5: {
    headerRow: { fill: [0, 0], color: [1, 0], bold: true, bottom: ['medium', [1, 0]] },
    totalRow: { fill: [7, D50], color: [1, 0], bold: true, top: ['medium', [1, 0]] },
    firstColumn: { fill: [7, D25], color: [1, 0], bold: true, right: ['medium', [1, 0]] },
    lastColumn: { fill: [7, D25], color: [1, 0], bold: true, left: ['medium', [1, 0]] },
    firstRowStripe: { fill: [7, D25] },
    firstColumnStripe: { fill: [7, D25] },
  },
  TableStyleDark6: {
    headerRow: { fill: [0, 0], color: [1, 0], bold: true, bottom: ['medium', [1, 0]] },
    totalRow: { fill: [8, D50], color: [1, 0], bold: true, top: ['medium', [1, 0]] },
    firstColumn: { fill: [8, D25], color: [1, 0], bold: true, right: ['medium', [1, 0]] },
    lastColumn: { fill: [8, D25], color: [1, 0], bold: true, left: ['medium', [1, 0]] },
    firstRowStripe: { fill: [8, D25] },
    firstColumnStripe: { fill: [8, D25] },
  },
  TableStyleDark7: {
    headerRow: { fill: [0, 0], color: [1, 0], bold: true, bottom: ['medium', [1, 0]] },
    totalRow: { fill: [9, D50], color: [1, 0], bold: true, top: ['medium', [1, 0]] },
    firstColumn: { fill: [9, D25], color: [1, 0], bold: true, right: ['medium', [1, 0]] },
    lastColumn: { fill: [9, D25], color: [1, 0], bold: true, left: ['medium', [1, 0]] },
    firstRowStripe: { fill: [9, D25] },
    firstColumnStripe: { fill: [9, D25] },
  },
  TableStyleDark8: {
    headerRow: { fill: [0, 0], color: [1, 0] },
    totalRow: { color: [0, 0], bold: true, top: ['double', [0, 0]] },
    firstColumn: { color: [0, 0], bold: true },
    lastColumn: { color: [0, 0], bold: true },
    firstRowStripe: { fill: [1, D35] },
    firstColumnStripe: { fill: [1, D35] },
  },
  TableStyleDark9: {
    headerRow: { fill: [5, 0], color: [1, 0] },
    totalRow: { color: [0, 0], bold: true, top: ['double', [0, 0]] },
    firstColumn: { color: [0, 0], bold: true },
    lastColumn: { color: [0, 0], bold: true },
    firstRowStripe: { fill: [4, L60] },
    firstColumnStripe: { fill: [4, L60] },
  },
  TableStyleDark10: {
    headerRow: { fill: [7, 0], color: [1, 0] },
    totalRow: { color: [0, 0], bold: true, top: ['double', [0, 0]] },
    firstColumn: { color: [0, 0], bold: true },
    lastColumn: { color: [0, 0], bold: true },
    firstRowStripe: { fill: [6, L60] },
    firstColumnStripe: { fill: [6, L60] },
  },
  TableStyleDark11: {
    headerRow: { fill: [9, 0], color: [1, 0] },
    totalRow: { color: [0, 0], bold: true, top: ['double', [0, 0]] },
    firstColumn: { color: [0, 0], bold: true },
    lastColumn: { color: [0, 0], bold: true },
    firstRowStripe: { fill: [8, L60] },
    firstColumnStripe: { fill: [8, L60] },
  },
};

/** The CSS colour of a clrScheme slot; the default Office palette without a workbook theme. */
const slotColor = (slot: number, themeColors?: readonly number[]): string => {
  const argb = themeColors?.[slot];
  if (argb !== undefined) return cssColorFromArgb(argb);
  if (slot === 0) return '#000000';
  if (slot === 1) return '#ffffff';
  return DEFAULT_THEME_ACCENTS[slot - 4] ?? '#000000';
};

/** The elements of the built-in table style `name`; null for any other name. */
export function builtinTableStyleElements(
  name: string,
  themeColors?: readonly number[],
): StyleElementFormats | null {
  const style = STYLES[name];
  if (!style) return null;
  const resolve = ([slot, tint]: ThemeRef): string => {
    const base = slotColor(slot, themeColors);
    return tint === 0 ? base : (tintCssColor(base, tint) ?? base);
  };
  const out: StyleElementFormats = {};
  for (const [type, def] of Object.entries(style) as [keyof StyleElementFormats, ElementDef][]) {
    const element: StyleElementFormat = {};
    if (def.fill) element.fill = resolve(def.fill);
    if (def.color) element.color = resolve(def.color);
    if (def.bold) element.bold = true;
    const borders: NonNullable<StyleElementFormat['borders']> = {};
    for (const side of ['top', 'bottom', 'left', 'right'] as const) {
      const edge = def[side];
      if (edge) borders[side] = { style: edge[0], color: resolve(edge[1]) };
    }
    if (Object.keys(borders).length > 0) element.borders = borders;
    out[type] = element;
  }
  return out;
}
