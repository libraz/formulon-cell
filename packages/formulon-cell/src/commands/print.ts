import type { PaginationPage, PaginationResult } from '@libraz/formulon';
import { colLetter, formatA1Cell, parseAddrKey } from '../engine/address.js';
import { effectiveCellFormatWithPending } from '../engine/cell-format-sync.js';
import { numFmtToFormatCode } from '../engine/format-writeback.js';
import type { CellValue } from '../engine/types.js';
// Print / PDF export.
//
// `buildPrintDocument` renders the current sheet as an HTML table with inline
// styles for cell formatting (bold/italic/colors/borders/alignment). The
// resulting fragment is wrapped in a self-contained HTML document with an
// `@page` block driven by the active `PageSetup` so the browser's native
// print dialog (Print → "Save as PDF") inherits orientation, paper size,
// margins, and header/footer text.
//
// `printSheet` lifts that document into a hidden iframe, calls
// `iframe.contentWindow.print()`, then removes the iframe once focus
// returns to the host window. We never spawn third-party PDF libs — the
// browser's print pipeline is the entire export path.
import { formatCell } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { numberFormatPaint } from '../format/number-format.js';
import { gradientFillCss } from '../render/gradient-fill-css.js';
import { cellFontFamily } from '../render/painters/text.js';
import { formatWithPending } from '../store/pending-format.js';
import { hiddenRowsOf } from '../store/row-visibility.js';
import {
  type CellFormat,
  defaultPageSetup,
  getPageSetup,
  type PageMargins,
  type PageSetup,
  type PrintCellErrorsMode,
  type SpreadsheetStore,
} from '../store/store.js';
import {
  computeFitToPagesScale,
  effectivePrintMargins,
  PAPER_DIMENSIONS,
  printColumnsForRegion,
  splitPrintRegionIntoTiles,
} from './page-geometry.js';
import { paginationFor } from './pagination.js';
import {
  orderPrintRegionsForPageOrder,
  type PrintAreaBounds,
  parsePrintAreas,
  parsePrintTitleCols,
  parsePrintTitleRows,
} from './print-ranges.js';
import { type PrinterProfile, resolvePrinterProfileBounds } from './printer-profile.js';
import { formatA1FormulaAsR1C1 } from './refs.js';

export type { AxisBand, PrintableMarginAdjustment, SplitAxisOptions } from './page-geometry.js';
export {
  computeFitToPagesScale,
  effectivePrintMargins,
  PRINT_PX_PER_INCH,
  paperInches,
  printableMarginAdjustments,
  printablePagePixels,
  splitAxisIntoBands,
} from './page-geometry.js';
export type { PrintAreaBounds } from './print-ranges.js';
export {
  orderPrintRegionsForPageOrder,
  parsePrintArea,
  parsePrintAreas,
  parsePrintTitleCols,
  parsePrintTitleRows,
} from './print-ranges.js';

/** Output of `buildPrintDocument`. `html` is a complete HTML document string
 *  ready to feed into an iframe via `document.write`. `cssVars` carries
 *  individual values (paper size token, margins, scale) so callers / tests
 *  can verify the page-setup wired through without parsing the HTML. */
export interface PrintDocument {
  html: string;
  cssVars: Record<string, string>;
}

export interface BuildPrintDocumentOptions {
  printableBounds?: PageMargins | null;
  /** Print the Outline and Shadow font effects, which only the Mac profile
   *  draws. */
  fontEffects?: boolean;
}

export interface PrintSheetOptions {
  printerProfiles?: readonly PrinterProfile[];
  printerProfileId?: string;
}

/** Refuse a print request before it can allocate a sheet-sized HTML table. */
export const MAX_PRINT_MATERIALIZED_CELLS = 1_000_000;
/** Keep the number of table sections bounded even when each section is sparse. */
export const MAX_PRINT_GENERATED_PAGES = 10_000;

export class PrintResourceLimitError extends Error {
  readonly code = 'PRINT_RESOURCE_LIMIT';

  constructor(
    public readonly cellCount: number,
    public readonly limit = MAX_PRINT_MATERIALIZED_CELLS,
  ) {
    super(
      `Print range contains ${cellCount.toLocaleString()} cells; reduce the print area below ${limit.toLocaleString()} cells and try again.`,
    );
    this.name = 'PrintResourceLimitError';
  }
}

const escapeHtml = (s: string): string =>
  s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');

const splitFontFamily = (family: string): string[] => {
  const parts: string[] = [];
  let start = 0;
  let quote = '';
  let depth = 0;
  for (let i = 0; i < family.length; i += 1) {
    const ch = family[i] ?? '';
    if (quote) {
      if (ch === quote && family[i - 1] !== '\\') quote = '';
      continue;
    }
    if (ch === '"' || ch === "'") {
      quote = ch;
      continue;
    }
    if (ch === '(') depth += 1;
    else if (ch === ')' && depth > 0) depth -= 1;
    else if (ch === ',' && depth === 0) {
      parts.push(family.slice(start, i).trim());
      start = i + 1;
    }
  }
  parts.push(family.slice(start).trim());
  return parts.filter(Boolean);
};

/** Escape a dynamic value for a CSS string. `<` and `>` are escaped even
 * inside quotes because the complete stylesheet is an HTML raw-text node. */
const cssString = (value: string): string =>
  `"${value
    .replaceAll('\\', '\\\\')
    .replaceAll('"', '\\"')
    .replace(/\r\n?|\n/g, '\\A ')
    .replaceAll('&', '\\26 ')
    .replaceAll('<', '\\3C ')
    .replaceAll('>', '\\3E ')}"`;

const safeFontFamily = (family: string): string =>
  splitFontFamily(family)
    .map((part) => {
      const unquoted =
        part.length >= 2 &&
        ((part.startsWith('"') && part.endsWith('"')) ||
          (part.startsWith("'") && part.endsWith("'")))
          ? part.slice(1, -1)
          : part;
      return /^[a-zA-Z][a-zA-Z0-9-]*$/.test(unquoted) ? unquoted : cssString(unquoted);
    })
    .join(', ');

const PRINT_FONT_FALLBACK = 'system-ui, -apple-system, "Segoe UI", sans-serif';

const cellDisplay = (
  value: CellValue,
  formula: string | null,
  showFormulas: boolean,
  addr?: { row: number; col: number },
  r1c1 = false,
  cellErrorsAs: PrintCellErrorsMode = 'displayed',
  nativeR1C1Formula?: string | null,
  nativeDisplay?: string | null,
  format?: CellFormat,
): string => {
  if (showFormulas && formula) {
    if (!r1c1) return formula;
    return nativeR1C1Formula ?? (addr ? formatA1FormulaAsR1C1(formula, addr) : formula);
  }
  if (value.kind === 'blank') return format?.hyperlinkDisplay ?? '';
  if (value.kind === 'error') {
    if (cellErrorsAs === 'blank') return '';
    if (cellErrorsAs === 'dash') return '--';
    if (cellErrorsAs === 'na') return '#N/A';
  }
  if (nativeDisplay !== null && nativeDisplay !== undefined) return nativeDisplay;
  return formatCell(value);
};

const fillPatternCss = (pattern: CellFormat['fillPattern'], color = '#000000'): string => {
  switch (pattern) {
    case 'gray0625':
      return `radial-gradient(${color} 0.4px, transparent 0.4px)`;
    case 'gray125':
      return `radial-gradient(${color} 0.6px, transparent 0.6px)`;
    case 'gray25':
      return `radial-gradient(${color} 1px, transparent 1px)`;
    case 'gray50':
      return `repeating-linear-gradient(45deg, ${color} 0 2px, transparent 2px 4px)`;
    case 'gray75':
      return `repeating-linear-gradient(45deg, ${color} 0 3px, transparent 3px 4px)`;
    case 'horizontal':
    case 'darkHorizontal':
      return `repeating-linear-gradient(0deg, ${color} 0 1px, transparent 1px 4px)`;
    case 'lightHorizontal':
      return `repeating-linear-gradient(0deg, ${color} 0 1px, transparent 1px 7px)`;
    case 'vertical':
    case 'darkVertical':
      return `repeating-linear-gradient(90deg, ${color} 0 1px, transparent 1px 4px)`;
    case 'lightVertical':
      return `repeating-linear-gradient(90deg, ${color} 0 1px, transparent 1px 7px)`;
    case 'diagonalDown':
    case 'darkDown':
      return `repeating-linear-gradient(45deg, ${color} 0 1px, transparent 1px 5px)`;
    case 'diagonalUp':
    case 'darkUp':
      return `repeating-linear-gradient(135deg, ${color} 0 1px, transparent 1px 5px)`;
    case 'lightDown':
      return `repeating-linear-gradient(45deg, ${color} 0 1px, transparent 1px 9px)`;
    case 'lightUp':
      return `repeating-linear-gradient(135deg, ${color} 0 1px, transparent 1px 9px)`;
    case 'darkGrid':
    case 'lightGrid': {
      const step = pattern === 'darkGrid' ? 4 : 7;
      return `repeating-linear-gradient(0deg, ${color} 0 1px, transparent 1px ${step}px), repeating-linear-gradient(90deg, ${color} 0 1px, transparent 1px ${step}px)`;
    }
    case 'darkTrellis':
    case 'lightTrellis': {
      const step = pattern === 'darkTrellis' ? 6 : 9;
      return `repeating-linear-gradient(45deg, ${color} 0 1px, transparent 1px ${step}px), repeating-linear-gradient(135deg, ${color} 0 1px, transparent 1px ${step}px)`;
    }
    default:
      return '';
  }
};

/** Translate a `CellFormat` into an inline-style string. Only the subset that
 *  matters for print is included — borders are emitted on the cell itself
 *  so colspan/rowspan merges retain their outline. */
function inlineCellStyle(
  fmt: CellFormat | undefined,
  value: CellValue,
  showGridlines: boolean,
  blackAndWhite: boolean,
  scale = 1,
  fontEffects = false,
): string {
  const parts: string[] = [];
  if (showGridlines) {
    // Default hairline grid for unformatted cells. Per-side overrides below.
    parts.push(`border:${scale}px solid #c8c8c8`);
  }
  if (!fmt) return parts.join(';');
  if (fmt.bold) parts.push('font-weight:600');
  if (fmt.italic) parts.push('font-style:italic');
  if (fontEffects && fmt.fontShadow)
    parts.push(`text-shadow:${scale}px ${scale}px 0 rgba(0,0,0,0.35)`);
  if (fontEffects && fmt.fontOutline) {
    parts.push(`-webkit-text-stroke:${0.75 * scale}px currentColor`);
    parts.push('-webkit-text-fill-color:transparent');
  }
  if (fmt.underline || fmt.strike) {
    const decos: string[] = [];
    if (fmt.underline) decos.push('underline');
    if (fmt.strike) decos.push('line-through');
    parts.push(`text-decoration:${decos.join(' ')}`);
  }
  if (fmt.align) {
    const align =
      fmt.align === 'centerContinuous'
        ? 'center'
        : fmt.align === 'justify' || fmt.align === 'distributed'
          ? 'justify'
          : fmt.align === 'fill'
            ? 'left'
            : fmt.align;
    parts.push(`text-align:${align}`);
  }
  if (fmt.vAlign) {
    const v = fmt.vAlign === 'justify' || fmt.vAlign === 'distributed' ? 'middle' : fmt.vAlign;
    parts.push(`vertical-align:${v}`);
  }
  if (fmt.wrap) parts.push('white-space:pre-wrap');
  else parts.push('white-space:nowrap');
  if (typeof fmt.indent === 'number' && fmt.indent > 0) {
    parts.push(`padding-left:${(4 + fmt.indent * 8) * scale}px`);
  }
  if (typeof fmt.rotation === 'number' && fmt.rotation !== 0) {
    parts.push(`transform:rotate(${-fmt.rotation}deg);transform-origin:center`);
  }
  if (fmt.textDirection === 'rtl') parts.push('direction:rtl');
  else if (fmt.textDirection === 'ltr') parts.push('direction:ltr');
  const formatCode = value.kind === 'number' ? numFmtToFormatCode(fmt.numFmt) : null;
  const formatColor =
    formatCode && value.kind === 'number'
      ? numberFormatPaint(formatCode, value.value).color
      : undefined;
  const color = formatColor ?? fmt.color;
  if (color && !blackAndWhite) parts.push(`color:${color}`);
  if (fmt.fillGradient && !blackAndWhite) {
    parts.push(`background:${gradientFillCss(fmt.fillGradient)}`);
  } else if (fmt.fill && !blackAndWhite) parts.push(`background:${fmt.fill}`);
  if (fmt.fillPattern && !fmt.fillGradient && !blackAndWhite) {
    const image = fillPatternCss(fmt.fillPattern, fmt.fillPatternColor);
    if (image) parts.push(`background-image:${image}`);
    if (fmt.fillPattern === 'gray125' || fmt.fillPattern === 'gray25') {
      parts.push(`background-size:${4 * scale}px ${4 * scale}px`);
    }
  }
  if (fmt.fontFamily) {
    parts.push(
      `font-family:${safeFontFamily(cellFontFamily(fmt.fontFamily, PRINT_FONT_FALLBACK))}`,
    );
  }
  if (typeof fmt.fontSize === 'number') parts.push(`font-size:${fmt.fontSize * scale}px`);
  if (fmt.borders) {
    const sideToCss = (side: 'top' | 'right' | 'bottom' | 'left'): void => {
      const b = fmt.borders?.[side];
      if (!b) return;
      if (b === true) {
        parts.push(`border-${side}:${scale}px solid #000`);
        return;
      }
      const width =
        b.style === 'thick'
          ? 2
          : b.style === 'medium'
            ? 1.5
            : b.style === 'double'
              ? 3
              : b.style === 'dashed' || b.style === 'mediumDashed' || b.style === 'dashDot'
                ? 1
                : b.style === 'dotted'
                  ? 1
                  : 1;
      const style =
        b.style === 'double'
          ? 'double'
          : b.style === 'dashed' || b.style === 'mediumDashed' || b.style === 'dashDot'
            ? 'dashed'
            : 'solid';
      parts.push(
        `border-${side}:${width * scale}px ${style} ${blackAndWhite ? '#000' : (b.color ?? '#000')}`,
      );
    };
    sideToCss('top');
    sideToCss('right');
    sideToCss('bottom');
    sideToCss('left');
  }
  return parts.join(';');
}

/** Build the @page CSS string for a given setup. Browsers honour orientation
 *  and (for print preview / PDF) the size keyword. */
/** Context used to expand header/footer field codes into concrete text. */
export interface HeaderFooterContext {
  sheetName: string;
  fileName: string;
  date: string;
  time: string;
}

/**
 * Expand an Excel header/footer field string into a CSS `content` value for a
 * `@page` margin box. Static codes (`&D` date, `&T` time, `&F` file, `&A` sheet)
 * become quoted literals; page-relative codes (`&P` page, `&N` pages) become
 * `counter(page)` / `counter(pages)` so the browser fills the real numbers on
 * every printed page. `&&` is a literal ampersand; formatting codes (`&B`,
 * `&"font"`, size runs, section markers) are stripped. Returns `''` when the
 * field is empty.
 */
export function headerFieldToCssContent(raw: string, ctx: HeaderFooterContext): string {
  if (!raw) return '';
  const parts: string[] = [];
  let buf = '';
  const flush = (): void => {
    if (buf) {
      parts.push(cssString(buf));
      buf = '';
    }
  };
  for (let i = 0; i < raw.length; i += 1) {
    const ch = raw[i] ?? '';
    if (ch !== '&') {
      buf += ch;
      continue;
    }
    const next = raw[i + 1] ?? '';
    const upper = next.toUpperCase();
    if (next === '&') {
      buf += '&';
      i += 1;
    } else if (upper === 'P' || upper === 'N') {
      flush();
      parts.push(upper === 'P' ? 'counter(page)' : 'counter(pages)');
      i += 1;
      // Drop an Excel page offset like `&P+1` — CSS counters cannot offset.
      const rest = raw.slice(i + 1);
      const off = /^[+-]\d+/.exec(rest);
      if (off) i += off[0].length;
    } else if (upper === 'D') {
      buf += ctx.date;
      i += 1;
    } else if (upper === 'T') {
      buf += ctx.time;
      i += 1;
    } else if (upper === 'F') {
      buf += ctx.fileName;
      i += 1;
    } else if (upper === 'A') {
      buf += ctx.sheetName;
      i += 1;
    } else if (next === '"') {
      // &"font name" — skip the quoted font spec entirely.
      const close = raw.indexOf('"', i + 2);
      i = close >= 0 ? close : i + 1;
    } else if (/[0-9]/.test(next)) {
      // &nn font-size run — drop the digits.
      let j = i + 1;
      while (j < raw.length && /[0-9]/.test(raw[j] ?? '')) j += 1;
      i = j - 1;
    } else {
      // Single-letter formatting / section code (&B, &I, &U, &L, &C, &R, &Z…).
      i += 1;
    }
  }
  flush();
  return parts.length ? parts.join(' ') : '';
}

function buildPageRule(
  setup: PageSetup,
  ctx: HeaderFooterContext,
  native?: PaginationResult,
): string {
  const size = native
    ? `${native.paper.widthPt}pt ${native.paper.heightPt}pt`
    : `${PAPER_DIMENSIONS[setup.paperSize] ?? 'A4'} ${setup.orientation}`;
  const m = native
    ? {
        top: `${native.margins.top}pt`,
        right: `${native.margins.right}pt`,
        bottom: `${native.margins.bottom}pt`,
        left: `${native.margins.left}pt`,
      }
    : (() => {
        const margins = effectivePrintMargins(setup);
        return {
          top: `${margins.top}in`,
          right: `${margins.right}in`,
          bottom: `${margins.bottom}in`,
          left: `${margins.left}in`,
        };
      })();
  const boxes: [string, string | undefined][] = [
    ['top-left', setup.headerLeft],
    ['top-center', setup.headerCenter],
    ['top-right', setup.headerRight],
    ['bottom-left', setup.footerLeft],
    ['bottom-center', setup.footerCenter],
    ['bottom-right', setup.footerRight],
  ];
  const marginBoxes = boxes
    .map(([box, raw]) => {
      const content = headerFieldToCssContent(raw ?? '', ctx);
      if (!content) return '';
      return `@${box} { content: ${content}; font-size: 9pt; color: #555; white-space: pre; }`;
    })
    .filter(Boolean)
    .join(' ');
  const base = `@page { size: ${size}; margin: ${m.top} ${m.right} ${m.bottom} ${m.left};`;
  return marginBoxes ? `${base} ${marginBoxes} }` : `${base} }`;
}

/** Snapshot of one cell ready to render. We collect into an intermediate map
 *  so the table builder can apply merges + print-titles without re-walking the
 *  workbook. */
interface CellSnap {
  value: CellValue;
  formula: string | null;
  format?: CellFormat;
}

interface PrintCommentSnap {
  row: number;
  col: number;
  text: string;
}

/** Render the print document HTML for `sheet`. The selection / viewport are
 *  ignored — printing always emits the full populated range. Hidden rows and
 *  columns honour the layout slice. */
export function buildPrintDocument(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
  title = 'Print',
  options: BuildPrintDocumentOptions = {},
): PrintDocument {
  const state = store.getState();
  const baseSetup = getPageSetup(state, sheet);
  const setup: PageSetup =
    options.printableBounds !== undefined
      ? {
          ...baseSetup,
          printableBounds: options.printableBounds ? { ...options.printableBounds } : undefined,
        }
      : baseSetup;
  const showFormulas = state.ui.showFormulas;
  const hiddenRows = hiddenRowsOf(state.layout);
  const hiddenCols = state.layout.hiddenCols;
  // The native paginator is authoritative when the worksheet's print setup
  // has been synchronized to the engine. A printer profile changes the
  // session-only printable insets, so keep the local splitter for that path.
  const nativePagination =
    options.printableBounds === undefined
      ? paginationFor(state, sheet, { engine: wb, purpose: 'print' })
      : null;
  const native = nativePagination?.native;

  // Collect populated cells. Track max row/col so the table doesn't run past
  // the last data cell.
  const cellMap = new Map<string, CellSnap>();
  const formatAt = (row: number, col: number): CellFormat | undefined => {
    const direct = formatWithPending(state, { sheet, row, col });
    if (wb.capabilities.effectiveStyle === true && !wb.isStub) {
      return (effectiveCellFormatWithPending(wb, state, { sheet, row, col }) ?? direct) as
        | CellFormat
        | undefined;
    }
    return direct;
  };
  let maxRow = -1;
  let maxCol = -1;
  for (const e of wb.cells(sheet)) {
    const key = `${e.addr.row}:${e.addr.col}`;
    const fmt = formatAt(e.addr.row, e.addr.col);
    cellMap.set(key, { value: e.value, formula: e.formula, format: fmt });
    if (e.addr.row > maxRow) maxRow = e.addr.row;
    if (e.addr.col > maxCol) maxCol = e.addr.col;
  }
  for (const [formatKey] of state.format.formats) {
    const addr = parseAddrKey(formatKey);
    if (!addr || addr.sheet !== sheet) continue;
    const { row, col } = addr;
    const key = `${row}:${col}`;
    if (!cellMap.has(key)) {
      cellMap.set(key, {
        value: { kind: 'blank' },
        formula: null,
        format: formatAt(row, col),
      });
    }
    if (row > maxRow) maxRow = row;
    if (col > maxCol) maxCol = col;
  }

  // Find merge extents before walking their interior. A single far-corner
  // merge must be rejected by the same preflight as a far-corner value.
  for (const range of state.merges.byAnchor.values()) {
    if (range.sheet !== sheet) continue;
    if (range.r1 > maxRow) maxRow = range.r1;
    if (range.c1 > maxCol) maxCol = range.c1;
  }
  if (maxRow < 0) {
    // Empty sheet — print at least one empty row so the user gets a real page.
    maxRow = 0;
    maxCol = 0;
  }

  const nativeHasPages = (native?.pages.length ?? 0) > 0;
  const titleRowRange = native?.printTitles.hasRows
    ? ([native.printTitles.firstRow, native.printTitles.lastRow] as [number, number])
    : parsePrintTitleRows(setup.printTitleRows);
  const titleColRange = native?.printTitles.hasCols
    ? ([native.printTitles.firstCol, native.printTitles.lastCol] as [number, number])
    : parsePrintTitleCols(setup.printTitleCols);
  const nativePrintAreas = native?.printArea.length
    ? native.printArea.map((area) => ({
        row0: area.firstRow,
        col0: area.firstCol,
        row1: area.lastRow,
        col1: area.lastCol,
      }))
    : null;
  const printAreas = nativePrintAreas ?? parsePrintAreas(setup.printArea);
  const rangeLength = (range: [number, number] | null): number =>
    range && Number.isFinite(range[0]) && Number.isFinite(range[1]) && range[1] >= range[0]
      ? range[1] - range[0] + 1
      : 0;
  const titleRowCount = rangeLength(titleRowRange);
  const titleColCount = rangeLength(titleColRange);
  const cappedAdd = (left: number, right: number): number =>
    Math.min(MAX_PRINT_MATERIALIZED_CELLS + 1, left + right);
  const cappedMultiply = (left: number, right: number): number => {
    if (left <= 0 || right <= 0) return 0;
    if (left > Math.floor((MAX_PRINT_MATERIALIZED_CELLS + 1) / right)) {
      return MAX_PRINT_MATERIALIZED_CELLS + 1;
    }
    return left * right;
  };
  // This is intentionally conservative: title rows/columns are emitted in
  // every table section, and headings add one row plus one corner column.
  const generatedCellsForRegion = (region: PrintAreaBounds): number => {
    const rows = cappedAdd(
      Math.max(0, region.row1 - region.row0 + 1),
      titleRowCount + (setup.showHeadings ? 1 : 0),
    );
    const cols = cappedAdd(
      Math.max(0, region.col1 - region.col0 + 1),
      titleColCount + (setup.showHeadings ? 1 : 0),
    );
    return cappedMultiply(rows, cols);
  };
  const throwIfOverPrintBudget = (estimatedCells: number): void => {
    if (estimatedCells > MAX_PRINT_MATERIALIZED_CELLS) {
      throw new PrintResourceLimitError(estimatedCells);
    }
  };

  if (nativeHasPages) {
    // Native pages may represent sparse content without materializing every
    // sheet cell. Budget the rectangles that this HTML path will actually
    // emit, while leaving a small sparse page valid.
    if (native && native.pages.length > MAX_PRINT_GENERATED_PAGES) {
      throw new PrintResourceLimitError(MAX_PRINT_MATERIALIZED_CELLS + 1);
    }
    let estimatedCells = 0;
    for (const page of native?.pages ?? []) {
      estimatedCells = cappedAdd(
        estimatedCells,
        generatedCellsForRegion({
          row0: page.firstRow,
          col0: page.firstCol,
          row1: page.lastRow,
          col1: page.lastCol,
        }),
      );
      if (estimatedCells > MAX_PRINT_MATERIALIZED_CELLS) break;
    }
    throwIfOverPrintBudget(estimatedCells);
  } else {
    // Local/printer-profile pagination materializes every cell in each tile.
    // Check the aggregate rectangle before merge maps, tiles, or HTML rows are
    // allocated so a sparse far-corner cell has an explicit bounded outcome.
    const areas = printAreas ?? [{ row0: 0, col0: 0, row1: maxRow, col1: maxCol }];
    let estimatedCells = 0;
    for (const area of areas) {
      const row1 = Math.min(maxRow, area.row1);
      const col1 = Math.min(maxCol, area.col1);
      if (row1 < area.row0 || col1 < area.col0) continue;
      estimatedCells = cappedAdd(
        estimatedCells,
        generatedCellsForRegion({ row0: area.row0, col0: area.col0, row1, col1 }),
      );
      if (estimatedCells > MAX_PRINT_MATERIALIZED_CELLS) break;
    }
    throwIfOverPrintBudget(estimatedCells);
  }

  // Merges shape colspan/rowspan + cell-skip set (cells inside but not at
  // the anchor are absorbed into the anchor's span).
  let mergeInteriorCells = 0;
  for (const range of state.merges.byAnchor.values()) {
    if (range.sheet !== sheet) continue;
    const mergeCells = cappedMultiply(
      Math.max(0, range.r1 - range.r0 + 1),
      Math.max(0, range.c1 - range.c0 + 1),
    );
    mergeInteriorCells = cappedAdd(mergeInteriorCells, Math.max(0, mergeCells - 1));
    if (mergeInteriorCells > MAX_PRINT_MATERIALIZED_CELLS) break;
  }
  // Every native page clones this skip set before rendering. Bound that
  // repeated allocation as well as the initial interior-cell walk.
  const mergeMaterializationCount = nativeHasPages
    ? cappedMultiply(mergeInteriorCells, native?.pages.length ?? 0)
    : mergeInteriorCells;
  throwIfOverPrintBudget(mergeMaterializationCount);
  const skipKeys = new Set<string>();
  const spanByAnchor = new Map<string, { rowspan: number; colspan: number }>();
  for (const [, range] of state.merges.byAnchor) {
    if (range.sheet !== sheet) continue;
    if (range.r1 > maxRow) maxRow = range.r1;
    if (range.c1 > maxCol) maxCol = range.c1;
    spanByAnchor.set(`${range.r0}:${range.c0}`, {
      rowspan: range.r1 - range.r0 + 1,
      colspan: range.c1 - range.c0 + 1,
    });
    for (let r = range.r0; r <= range.r1; r += 1) {
      for (let c = range.c0; c <= range.c1; c += 1) {
        if (r === range.r0 && c === range.c0) continue;
        skipKeys.add(`${r}:${c}`);
      }
    }
  }

  const manualRowBreaks = new Set(setup.manualPageBreakRows ?? []);
  const manualColBreaks = new Set(setup.manualPageBreakCols ?? []);
  const printRegions: PrintAreaBounds[] = orderPrintRegionsForPageOrder(
    (printAreas ?? [{ row0: 0, col0: 0, row1: maxRow, col1: maxCol }]).map((area) => {
      const row1 = Math.min(maxRow, area.row1);
      const col1 = Math.min(maxCol, area.col1);
      return {
        row0: area.row0,
        col0: area.col0,
        row1: row1 < area.row0 ? area.row0 : row1,
        col1: col1 < area.col0 ? area.col0 : col1,
      };
    }),
    setup.pageOrder,
  );
  const cellInPrintRegions = (row: number, col: number): boolean =>
    printRegions.some(
      (area) => row >= area.row0 && row <= area.row1 && col >= area.col0 && col <= area.col1,
    );
  const columnsForRegion = (region: PrintAreaBounds): number[] =>
    printColumnsForRegion(region, titleColRange);
  const isTitleCol = (col: number): boolean =>
    titleColRange !== null && col >= titleColRange[0] && col <= titleColRange[1];

  const commentEntries: PrintCommentSnap[] =
    setup.comments === 'none'
      ? []
      : [...state.format.formats.entries()]
          .map(([formatKey, fmt]): PrintCommentSnap | null => {
            if (typeof fmt.comment !== 'string' || fmt.comment.length === 0) return null;
            const addr = parseAddrKey(formatKey);
            if (!addr || addr.sheet !== sheet) return null;
            const { row, col } = addr;
            if (!cellInPrintRegions(row, col)) return null;
            if (hiddenRows.has(row) || hiddenCols.has(col)) return null;
            return { row, col, text: fmt.comment };
          })
          .filter((entry): entry is PrintCommentSnap => entry !== null)
          .sort((a, b) => a.row - b.row || a.col - b.col);
  const commentByCell = new Map(
    commentEntries.map((entry) => [`${entry.row}:${entry.col}`, entry]),
  );

  const r1c1FormulaFor = (row: number, col: number, formula: string | null): string | null => {
    if (!state.ui.r1c1 || !formula || wb.capabilities.formulaR1C1 !== true) return null;
    try {
      return wb.cellFormulaR1C1({ sheet, row, col });
    } catch {
      return null;
    }
  };

  const effectiveFormatFor = (
    row: number,
    col: number,
    snap: CellSnap | undefined,
  ): CellFormat | undefined => {
    const merged = snap?.format ?? formatAt(row, col) ?? {};
    return Object.keys(merged).length > 0 ? merged : undefined;
  };

  const nativeDisplayFor = (
    row: number,
    col: number,
    snap: CellSnap | undefined,
    format: CellFormat | undefined,
  ): string | null => {
    if (state.ui.showFormulas === true || !snap || snap.value.kind === 'blank') return null;
    // A store-side number-format edit can briefly lead the mirror. Format the
    // value directly in that case; otherwise the workbook renderer accounts
    // for inherited styles, date systems, and display overflow.
    const ownFormatCode = format?.numFmt ? numFmtToFormatCode(format.numFmt, wb) : null;
    if (ownFormatCode && typeof wb.formatValueForDisplay === 'function') {
      return wb.formatValueForDisplay(snap.value, ownFormatCode);
    }
    if (format?.hyperlinkDisplay || typeof wb.getDisplayText !== 'function') return null;
    const displayed = wb.getDisplayText(sheet, row, col);
    if (displayed === '') return null;
    return displayed;
  };

  const nativeScale = native ? Math.max(Number.isFinite(native.scale) ? native.scale : 1, 0.01) : 1;
  // Chromium rounds collapsed table borders and fractional point row heights
  // independently. Leave a small physical reserve so the final row cannot be
  // clipped when the native extent lands exactly on a browser pixel boundary.
  const NATIVE_LAYOUT_EPSILON_PT = 3;
  const rowHeightPt = (row: number): number =>
    ((state.layout.rowHeights.get(row) ?? state.layout.defaultRowHeight) * 72) / 96;
  const colWidthPt = (col: number): number =>
    ((state.layout.colWidths.get(col) ?? state.layout.defaultColWidth) * 72) / 96;

  const nativePageMetrics = (page: PaginationPage) => {
    if (!native) return null;
    const titleColsOffset = Math.max(0, page.originXPt - native.printable.x);
    const titleRowsOffset = Math.max(0, page.originYPt - native.printable.y);
    return {
      titleColsOffset,
      titleRowsOffset,
      // `printable` already accounts for the header/footer encroachment that
      // the native Windows geometry contract applies to the body height.
      // Using paper-minus-margins here would let a browser lay out one extra
      // row below a native footer band.
      contentWidth: Math.max(1, native.printable.width),
      contentHeight: Math.max(1, native.printable.height),
      tableWidth: Math.max(1, titleColsOffset + page.widthPt),
      tableHeight: Math.max(1, titleRowsOffset + page.heightPt),
    };
  };

  // Build column letters used for the optional `<colgroup>` headings strip.
  const renderCell = (
    row: number,
    col: number,
    tag: 'td' | 'th',
    regionMaxCol: number,
    renderedSkipKeys: Set<string>,
    cellScale = 1,
    nativeClip = false,
  ): string => {
    const key = `${row}:${col}`;
    if (renderedSkipKeys.has(key)) return '';
    const snap = cellMap.get(key);
    const value: CellValue = snap?.value ?? { kind: 'blank' };
    const format = effectiveFormatFor(row, col, snap);
    const text = cellDisplay(
      value,
      snap?.formula ?? null,
      showFormulas,
      { row, col },
      state.ui.r1c1 === true,
      setup.cellErrorsAs,
      r1c1FormulaFor(row, col, snap?.formula ?? null),
      nativeDisplayFor(row, col, snap, format),
      format,
    );
    const span = (() => {
      const mergeSpan = spanByAnchor.get(key);
      if (mergeSpan) return mergeSpan;
      if (format?.align !== 'centerContinuous' || !text) return undefined;
      let colspan = 1;
      for (let nextCol = col + 1; nextCol <= regionMaxCol; nextCol += 1) {
        if (hiddenCols.has(nextCol)) continue;
        const nextKey = `${row}:${nextCol}`;
        if (renderedSkipKeys.has(nextKey)) break;
        const nextSnap = cellMap.get(nextKey);
        const nextFormat = effectiveFormatFor(row, nextCol, nextSnap);
        const nextText = cellDisplay(
          nextSnap?.value ?? { kind: 'blank' },
          nextSnap?.formula ?? null,
          showFormulas,
          { row, col: nextCol },
          state.ui.r1c1 === true,
          setup.cellErrorsAs,
          r1c1FormulaFor(row, nextCol, nextSnap?.formula ?? null),
          nativeDisplayFor(row, nextCol, nextSnap, nextFormat),
          nextFormat,
        );
        if (nextFormat?.align !== 'centerContinuous' || nextText) break;
        renderedSkipKeys.add(nextKey);
        colspan += 1;
      }
      return colspan > 1 ? { rowspan: 1, colspan } : undefined;
    })();
    const style = inlineCellStyle(
      format,
      value,
      setup.showGridlines === true,
      setup.blackAndWhite === true,
      cellScale,
      options.fontEffects === true,
    );
    const rowspanAttr = span && span.rowspan > 1 ? ` rowspan="${span.rowspan}"` : '';
    const colspanAttr = span && span.colspan > 1 ? ` colspan="${span.colspan}"` : '';
    const styleAttr = style ? ` style="${escapeHtml(style)}"` : '';
    const comment = setup.comments === 'asDisplayed' ? commentByCell.get(key) : undefined;
    const commentHtml = comment
      ? `<div class="fc-print__comment-note">${escapeHtml(comment.text)}</div>`
      : '';
    const verticalTextStyle = format?.fontVertAlign
      ? `font-size:${
          typeof format.fontSize === 'number' && Number.isFinite(format.fontSize)
            ? `${format.fontSize * cellScale * 0.7}px`
            : '70%'
        };vertical-align:${format.fontVertAlign === 'superscript' ? 'super' : 'sub'}`
      : '';
    const textHtml = verticalTextStyle
      ? `<span style="${verticalTextStyle}">${escapeHtml(text)}</span>`
      : escapeHtml(text);
    const verticalAlignment = format?.vAlign;
    const contentJustify =
      verticalAlignment === 'top'
        ? 'flex-start'
        : verticalAlignment && verticalAlignment !== 'bottom'
          ? 'center'
          : 'flex-end';
    const cellContent = nativeClip
      ? `<div class="fc-print__cell-content" style="justify-content:${contentJustify}">${textHtml}${commentHtml}</div>`
      : `${textHtml}${commentHtml}`;
    const classes = [
      manualColBreaks.has(col) && !native ? 'fc-print__manual-break-col' : '',
      isTitleCol(col) ? 'fc-print__title-col-cell' : '',
    ].filter(Boolean);
    const classAttr = classes.length ? ` class="${classes.join(' ')}"` : '';
    return `<${tag}${classAttr}${rowspanAttr}${colspanAttr}${styleAttr}>${cellContent}</${tag}>`;
  };

  const renderRow = (
    row: number,
    isTitle: boolean,
    region: PrintAreaBounds,
    renderedSkipKeys: Set<string>,
    nativePage?: PaginationPage,
    nativeHeightPt?: number,
  ): string => {
    if (hiddenRows.has(row)) return '';
    const cells: string[] = [];
    if (setup.showHeadings) {
      cells.push(`<th class="fc-print__rowhead">${row + 1}</th>`);
    }
    const cols = columnsForRegion(region);
    const regionMaxCol = Math.max(...cols, region.col1);
    for (const c of cols) {
      if (hiddenCols.has(c)) continue;
      cells.push(
        renderCell(
          row,
          c,
          isTitle ? 'th' : 'td',
          regionMaxCol,
          renderedSkipKeys,
          nativePage ? nativeScale : 1,
          nativePage !== undefined,
        ),
      );
    }
    const cls =
      manualRowBreaks.has(row) && !isTitle && !nativePage
        ? ' class="fc-print__manual-break-row"'
        : '';
    const nativeHeight = nativePage
      ? ` style="height:${nativeHeightPt ?? rowHeightPt(row)}pt"`
      : '';
    return `<tr${cls}${nativeHeight}>${cells.join('')}</tr>`;
  };

  // Optional column-letter heading row. Repeated inside <thead> when print-
  // title rows are configured so it appears at the top of every page.
  const headingRow = (region: PrintAreaBounds): string => {
    if (!setup.showHeadings) return '';
    const cells: string[] = [`<th class="fc-print__corner"></th>`];
    for (const c of columnsForRegion(region)) {
      if (hiddenCols.has(c)) continue;
      const cls =
        manualColBreaks.has(c) && !native
          ? 'fc-print__colhead fc-print__manual-break-col'
          : 'fc-print__colhead';
      cells.push(`<th class="${cls}">${colLetter(c)}</th>`);
    }
    return `<tr>${cells.join('')}</tr>`;
  };

  const renderTableSection = (
    region: PrintAreaBounds,
    sectionIndex: number,
    nativePage?: PaginationPage,
  ): string => {
    const nativeMetrics = nativePage ? nativePageMetrics(nativePage) : null;
    const renderedSkipKeys = new Set(skipKeys);
    const titleSet = new Set<number>();
    if (titleRowRange) {
      for (let r = titleRowRange[0]; r <= titleRowRange[1]; r += 1) titleSet.add(r);
    }
    const titleRowsInRegion = [...titleSet].filter(
      (row) => row >= region.row0 && row <= region.row1 && !hiddenRows.has(row),
    );
    const actualTitleHeight = titleRowsInRegion.reduce((sum, row) => sum + rowHeightPt(row), 0);
    const bodyNaturalHeight = Array.from(
      { length: Math.max(0, region.row1 - region.row0 + 1) },
      (_, i) => region.row0 + i,
    )
      .filter((row) => !titleSet.has(row) && !hiddenRows.has(row))
      .reduce((sum, row) => sum + rowHeightPt(row), 0);
    // Native page extents are already scaled physical points. Keep the table
    // in that coordinate system so Chromium paginates the same boxes it
    // paints. The first native page includes its title rows in `heightPt`;
    // later pages report only their body block and reserve repeated titles in
    // `originYPt`.
    const nativeBodyHeight = nativePage
      ? Math.max(
          0.01,
          nativePage.heightPt -
            actualTitleHeight * nativeScale -
            (bodyNaturalHeight > 0 ? NATIVE_LAYOUT_EPSILON_PT : 0),
        )
      : 0;
    const nativeBodyScale =
      nativePage && bodyNaturalHeight > 0 ? nativeBodyHeight / bodyNaturalHeight : 1;
    const nativeRowHeightFor = (row: number, isTitle: boolean): number | undefined => {
      if (!nativePage) return undefined;
      if (isTitle) return rowHeightPt(row) * nativeScale;
      return rowHeightPt(row) * nativeBodyScale;
    };
    const theadRows: string[] = [];
    const heading = headingRow(region);
    if (heading) theadRows.push(heading);
    if (titleRowRange) {
      for (let r = titleRowRange[0]; r <= titleRowRange[1]; r += 1) {
        const html = renderRow(
          r,
          true,
          region,
          renderedSkipKeys,
          nativePage,
          nativeRowHeightFor(r, true),
        );
        if (html) theadRows.push(html);
      }
      if (nativePage && nativeMetrics) {
        const titleReserve = nativeMetrics.titleRowsOffset;
        const renderedTitleHeight = actualTitleHeight * nativeScale;
        const extraTitleHeight = Math.max(0, titleReserve - renderedTitleHeight);
        if (extraTitleHeight > 0.01) {
          const colCount = Math.max(
            1,
            columnsForRegion(region).filter((c) => !hiddenCols.has(c)).length,
          );
          theadRows.push(
            `<tr class="fc-print__native-title-spacer"><th colspan="${colCount}" style="height:${extraTitleHeight}pt;padding:0;border:0"></th></tr>`,
          );
        }
      }
    }
    const thead = theadRows.length ? `<thead>${theadRows.join('')}</thead>` : '';

    const bodyRows: string[] = [];
    for (let r = region.row0; r <= region.row1; r += 1) {
      if (titleSet.has(r)) continue;
      const html = renderRow(
        r,
        false,
        region,
        renderedSkipKeys,
        nativePage,
        nativeRowHeightFor(r, false),
      );
      if (html) bodyRows.push(html);
    }
    const tbody = `<tbody>${bodyRows.join('')}</tbody>`;

    const colgroup = (() => {
      // The fallback leaves column sizing to the browser unless repeated
      // title columns need a stable `<col>` strip. Native pages always get a
      // fixed colgroup: their page rectangle was measured in print points,
      // and an auto-layout table would otherwise reflow it to the longest
      // cell text and spill past the native page edge.
      if (!nativePage || !nativeMetrics) {
        if (!titleColRange) return '';
        const cols: string[] = [];
        if (setup.showHeadings) cols.push('<col>');
        for (const c of columnsForRegion(region)) {
          if (hiddenCols.has(c)) continue;
          cols.push(isTitleCol(c) ? '<col class="fc-print__title-col">' : '<col>');
        }
        return `<colgroup>${cols.join('')}</colgroup>`;
      }
      const cols: string[] = [];
      if (setup.showHeadings) cols.push('<col>');
      const visibleCols = columnsForRegion(region).filter((c) => !hiddenCols.has(c));
      const titleColsOffset = nativeMetrics?.titleColsOffset ?? 0;
      const hasSeparateTitleCols =
        nativePage !== undefined &&
        titleColRange !== null &&
        nativePage.firstCol > titleColRange[1];
      const titleNaturalWidth = visibleCols
        .filter((c) => hasSeparateTitleCols && isTitleCol(c))
        .reduce((sum, c) => sum + colWidthPt(c), 0);
      const bodyNaturalWidth = visibleCols
        .filter((c) => !hasSeparateTitleCols || !isTitleCol(c))
        .reduce((sum, c) => sum + colWidthPt(c), 0);
      const widthForCol = (col: number): number => {
        if (!nativePage || !nativeMetrics) return 0;
        if (hasSeparateTitleCols && isTitleCol(col)) {
          return titleNaturalWidth > 0
            ? (titleColsOffset * colWidthPt(col)) / titleNaturalWidth
            : 0;
        }
        const target = hasSeparateTitleCols ? nativePage.widthPt : nativeMetrics.tableWidth;
        return bodyNaturalWidth > 0 ? (target * colWidthPt(col)) / bodyNaturalWidth : 0;
      };
      for (const c of visibleCols) {
        const width = widthForCol(c);
        const style = width > 0 ? ` style="width:${width}pt"` : '';
        cols.push(isTitleCol(c) ? `<col class="fc-print__title-col"${style}>` : `<col${style}>`);
      }
      if (cols.length === (setup.showHeadings ? 1 : 0)) cols.push('<col style="width:100%">');
      return `<colgroup>${cols.join('')}</colgroup>`;
    })();

    const sectionClass =
      sectionIndex === 0 ? 'fc-print__area' : 'fc-print__area fc-print__area--break';
    if (!nativePage || !nativeMetrics) {
      return `<section class="${sectionClass}"><table class="fc-print__table"${tableTransform}>${colgroup}${thead}${tbody}</table></section>`;
    }
    const pageStyle = ` style="position:relative;box-sizing:border-box;width:${nativeMetrics.contentWidth}pt;height:${nativeMetrics.contentHeight}pt;overflow:hidden;break-inside:avoid"`;
    const nativeLeft = setup.centerHorizontally
      ? Math.max(0, (nativeMetrics.contentWidth - nativeMetrics.tableWidth) / 2)
      : 0;
    const nativeTop = setup.centerVertically
      ? Math.max(0, (nativeMetrics.contentHeight - nativeMetrics.tableHeight) / 2)
      : 0;
    // `PaginationPage` values are already scaled physical extents. A CSS
    // transform would make the browser paginate the pre-scale table and only
    // shrink its paint, dropping rows at page boundaries. Give layout the
    // final point sizes directly instead.
    const tableStyle = ` style="position:absolute;left:${nativeLeft}pt;top:${nativeTop}pt;width:${nativeMetrics.tableWidth}pt;height:${nativeMetrics.tableHeight}pt;table-layout:fixed;margin:0"`;
    const placementAttrs = ` data-native-origin-x-pt="${nativePage.originXPt}" data-native-origin-y-pt="${nativePage.originYPt}" data-native-width-pt="${nativePage.widthPt}" data-native-height-pt="${nativePage.heightPt}"`;
    return `<section class="${sectionClass} fc-print__native-page"${pageStyle}${placementAttrs}><table class="fc-print__table"${tableStyle}>${colgroup}${thead}${tbody}</table></section>`;
  };

  // Header / footer are rendered by the browser's paged-media engine via
  // `@page` margin boxes (see buildPageRule) so they repeat on every page and
  // their Excel field codes (&P, &N, &D, &F, &A…) expand to real values —
  // rather than a single body-flow strip that printed the codes verbatim.
  const now = new Date();
  const headerFooterCtx: HeaderFooterContext = {
    sheetName: wb.sheetName(sheet),
    fileName: title,
    date: now.toLocaleDateString(),
    time: now.toLocaleTimeString(),
  };

  const commentsHtml =
    setup.comments === 'endOfSheet' && commentEntries.length
      ? `<section class="fc-print__comments"><h2>Comments and Notes</h2><table><tbody>${commentEntries
          .map(
            (entry) =>
              `<tr><th>${formatA1Cell(entry.row, entry.col)}</th><td>${escapeHtml(
                entry.text,
              )}</td></tr>`,
          )
          .join('')}</tbody></table></section>`
      : '';

  // Fit-to-pages now derives a real scale from the content's natural size
  // instead of being ignored (which silently printed at 100%).
  const scale =
    native?.scale ??
    computeFitToPagesScale(setup, printRegions, state.layout, hiddenRows, hiddenCols);
  const tableTransform =
    scale === 1 ? '' : ` style="transform:scale(${scale});transform-origin:top left;"`;

  const paginatedRegions = native?.pages.length
    ? native.pages.map((page) => ({
        region: {
          row0: page.firstRow,
          col0: page.firstCol,
          row1: page.lastRow,
          col1: page.lastCol,
        },
        nativePage: page,
      }))
    : printRegions.flatMap((region) =>
        splitPrintRegionIntoTiles(
          region,
          setup,
          state.layout,
          hiddenRows,
          hiddenCols,
          titleRowRange,
          titleColRange,
          scale,
        ).map((tile) => ({ region: tile, nativePage: undefined })),
      );

  // The native preflight above covers the engine's rectangles. Re-check the
  // concrete sections after local tiling so repeated title rows/columns and a
  // large number of generated tables cannot grow the HTML without a bound.
  if (paginatedRegions.length > MAX_PRINT_GENERATED_PAGES) {
    throw new PrintResourceLimitError(MAX_PRINT_MATERIALIZED_CELLS + 1);
  }
  let generatedCells = 0;
  for (const { region } of paginatedRegions) {
    generatedCells = cappedAdd(generatedCells, generatedCellsForRegion(region));
    if (generatedCells > MAX_PRINT_MATERIALIZED_CELLS) break;
  }
  throwIfOverPrintBudget(generatedCells);

  const tablesHtml = paginatedRegions
    .map(({ region, nativePage }, index) => renderTableSection(region, index, nativePage))
    .join('');

  const effectiveMargins = native
    ? {
        top: native.margins.top / 72,
        right: native.margins.right / 72,
        bottom: native.margins.bottom / 72,
        left: native.margins.left / 72,
      }
    : effectivePrintMargins(setup);
  const cssVars: Record<string, string> = {
    '--fc-print-paper': PAPER_DIMENSIONS[setup.paperSize] ?? 'A4',
    '--fc-print-orient': setup.orientation,
    '--fc-print-margin-top': `${setup.margins.top}in`,
    '--fc-print-margin-right': `${setup.margins.right}in`,
    '--fc-print-margin-bottom': `${setup.margins.bottom}in`,
    '--fc-print-margin-left': `${setup.margins.left}in`,
    '--fc-print-effective-margin-top': `${effectiveMargins.top}in`,
    '--fc-print-effective-margin-right': `${effectiveMargins.right}in`,
    '--fc-print-effective-margin-bottom': `${effectiveMargins.bottom}in`,
    '--fc-print-effective-margin-left': `${effectiveMargins.left}in`,
    '--fc-print-printable-top': `${setup.printableBounds?.top ?? 0}in`,
    '--fc-print-printable-right': `${setup.printableBounds?.right ?? 0}in`,
    '--fc-print-printable-bottom': `${setup.printableBounds?.bottom ?? 0}in`,
    '--fc-print-printable-left': `${setup.printableBounds?.left ?? 0}in`,
    '--fc-print-header-margin': `${setup.headerMargin ?? 0.3}in`,
    '--fc-print-footer-margin': `${setup.footerMargin ?? 0.3}in`,
    '--fc-print-scale': String(scale),
    '--fc-print-fit-width': String(setup.fitWidth ?? 0),
    '--fc-print-fit-height': String(setup.fitHeight ?? 0),
    '--fc-print-quality': setup.printQuality ?? 'automatic',
    '--fc-print-first-page': String(setup.firstPageNumber ?? 'auto'),
    '--fc-print-header-footer-scale': String(
      setup.scaleHeaderFooterWithDocument === false ? 1 : scale,
    ),
  };
  const bodyClasses = [
    native?.pages.length ? 'fc-print--native' : '',
    setup.centerHorizontally ? 'fc-print--center-h' : '',
    setup.centerVertically ? 'fc-print--center-v' : '',
    setup.blackAndWhite ? 'fc-print--black-white' : '',
    setup.draftQuality ? 'fc-print--draft' : '',
    (setup.fitWidth ?? 0) > 0 || (setup.fitHeight ?? 0) > 0 ? 'fc-print--fit-to-pages' : '',
    setup.alignHeaderFooterWithMargins === false ? 'fc-print--hf-free' : '',
  ]
    .filter(Boolean)
    .join(' ');

  const css = [
    buildPageRule(setup, headerFooterCtx, native),
    'body { margin: 0; font-family: system-ui, -apple-system, "Segoe UI", sans-serif; font-size: 10pt; color: #111; }',
    'table.fc-print__table { border-collapse: collapse; width: 100%; table-layout: auto; }',
    'table.fc-print__table td, table.fc-print__table th { padding: 2px 6px; vertical-align: bottom; box-sizing: border-box; }',
    'table.fc-print__table thead th { background: #f5f5f5; }',
    'table.fc-print__table .fc-print__title-col-cell { background: #f8f8f8; }',
    '.fc-print__rowhead, .fc-print__colhead, .fc-print__corner { background: #f0f0f0; color: #555; font-weight: 500; text-align: center; min-width: 24px; }',
    '.fc-print__header, .fc-print__footer { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 8px; padding: 4px 0; font-size: 9pt; color: #555; }',
    '.fc-print__header, .fc-print__footer { transform: scale(var(--fc-print-header-footer-scale, 1)); transform-origin: top left; }',
    '.fc-print__header { margin-top: calc(-1 * var(--fc-print-header-margin, 0.3in)); }',
    '.fc-print__footer { margin-bottom: calc(-1 * var(--fc-print-footer-margin, 0.3in)); }',
    '.fc-print--hf-free .fc-print__header, .fc-print--hf-free .fc-print__footer { margin-left: calc(-1 * var(--fc-print-margin-left)); margin-right: calc(-1 * var(--fc-print-margin-right)); }',
    '.fc-print__header > :nth-child(2), .fc-print__footer > :nth-child(2) { text-align: center; }',
    '.fc-print__header > :nth-child(3), .fc-print__footer > :nth-child(3) { text-align: right; }',
    '.fc-print--center-h .fc-print__table { margin-left: auto; margin-right: auto; width: auto; }',
    '.fc-print--center-v .fc-print__area { min-height: calc(100vh - var(--fc-print-effective-margin-top) - var(--fc-print-effective-margin-bottom)); display: flex; flex-direction: column; justify-content: center; }',
    '.fc-print--native .fc-print__native-page { min-height: 0 !important; display: block !important; justify-content: initial !important; }',
    `.fc-print--native table.fc-print__table { font-size: ${10 * nativeScale}pt; line-height: 1; }`,
    `.fc-print--native table.fc-print__table td, .fc-print--native table.fc-print__table th { position: relative; padding: ${Math.max(0, 3 * nativeScale - 1)}px ${6 * nativeScale}px; line-height: 1; overflow: hidden; }`,
    '.fc-print--native .fc-print__cell-content { position: absolute; inset: 0; display: flex; flex-direction: column; box-sizing: border-box; overflow: hidden; padding: inherit; line-height: inherit; white-space: inherit; }',
    `.fc-print--native .fc-print__comment-note { margin-top: ${3 * nativeScale}px; padding: ${3 * nativeScale}px ${5 * nativeScale}px; font-size: ${8 * nativeScale}pt; }`,
    '.fc-print--native .fc-print__native-page > .fc-print__table { margin: 0 !important; }',
    '.fc-print--black-white { filter: grayscale(1); }',
    '.fc-print--draft * { box-shadow: none !important; text-shadow: none !important; }',
    '.fc-print--fit-to-pages .fc-print__table { width: 100%; max-width: 100%; }',
    '.fc-print__area--break { break-before: page; page-break-before: always; }',
    '.fc-print__comment-note { margin-top: 3px; padding: 3px 5px; border: 1px solid #d6a100; background: #fff8cc; color: #3b3a00; font-size: 8pt; white-space: normal; }',
    '.fc-print__comments { break-before: page; page-break-before: always; margin-top: 16px; }',
    '.fc-print__comments h2 { font-size: 12pt; margin: 0 0 8px; }',
    '.fc-print__comments table { border-collapse: collapse; width: 100%; }',
    '.fc-print__comments th, .fc-print__comments td { border: 1px solid #c8c8c8; padding: 4px 6px; text-align: left; vertical-align: top; }',
    '.fc-print__comments th { width: 80px; background: #f0f0f0; }',
    typeof setup.firstPageNumber === 'number'
      ? `body { counter-reset: page ${Math.max(0, setup.firstPageNumber - 1)}; }`
      : '',
    '.fc-print__title-col { background-color: rgba(0,0,0,0.0); }',
    '.fc-print__manual-break-row { break-before: page; page-break-before: always; }',
    '.fc-print__manual-break-col { break-before: page; page-break-before: always; }',
    '@media print { thead { display: table-header-group; } tfoot { display: table-footer-group; } tr, td, th { page-break-inside: avoid; } .fc-print__manual-break-row, .fc-print__manual-break-col { break-before: page; page-break-before: always; } }',
  ].join('\n');

  const html = `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>${escapeHtml(title)}</title>
<style>
${css}
</style>
</head>
<body class="${bodyClasses}">
${tablesHtml}
${commentsHtml}
</body>
</html>`;

  return { html, cssVars };
}

/** Mount the print document into a hidden iframe attached to `host`, then
 *  invoke the iframe's print dialog. Returns a disposer that removes the
 *  iframe — callers usually let the auto-cleanup (`afterprint` / focus
 *  return) handle removal. Falls back to `window.print()` when the iframe
 *  contentWindow is unavailable (e.g. SSR / detached host). */
export function printSheet(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
  host: HTMLElement,
  title?: string,
  mode: 'print' | 'pdf' = 'print',
  options: PrintSheetOptions = {},
): () => void {
  const setup = getPageSetup(store.getState(), sheet);
  const printableBounds = options.printerProfiles
    ? (resolvePrinterProfileBounds(setup, options.printerProfiles, options.printerProfileId) ??
      null)
    : undefined;
  const doc = buildPrintDocument(wb, store, sheet, title, {
    printableBounds,
    fontEffects: (host.closest<HTMLElement>('.fc-host') ?? host).dataset.fcPlatform === 'mac',
  });
  const iframe = document.createElement('iframe');
  iframe.setAttribute('aria-hidden', 'true');
  iframe.dataset.fcPrintMode = mode;
  iframe.style.position = 'fixed';
  iframe.style.right = '0';
  iframe.style.bottom = '0';
  iframe.style.width = '0';
  iframe.style.height = '0';
  iframe.style.border = '0';
  iframe.style.visibility = 'hidden';
  // Same-origin srcdoc lets us drive `print()` without a network round-trip.
  iframe.srcdoc = doc.html;
  host.appendChild(iframe);

  let removed = false;
  const remove = (): void => {
    if (removed) return;
    removed = true;
    iframe.remove();
  };

  const triggerPrint = (): void => {
    const cw = iframe.contentWindow;
    if (!cw) {
      // No contentWindow (SSR / detached) — fall back to window.print so the
      // dialog still shows. The host page won't carry our @page rules but
      // the operation completes.
      window.print();
      remove();
      return;
    }
    try {
      cw.focus();
      cw.print();
    } catch {
      // Some browsers throw on synchronous print across iframes — fall back.
      window.print();
    }
    // Schedule removal after the print dialog closes. We listen on both the
    // iframe's afterprint event and a focus return on the host page; whichever
    // fires first triggers cleanup.
    const onAfterPrint = (): void => remove();
    cw.addEventListener?.('afterprint', onAfterPrint, { once: true });
    window.setTimeout(remove, 60_000);
  };

  // srcdoc loads asynchronously — wait for `load` so the document body has
  // rendered before we ask it to print.
  if (iframe.contentDocument?.readyState === 'complete') {
    triggerPrint();
  } else {
    iframe.addEventListener('load', triggerPrint, { once: true });
  }

  return remove;
}

/** Re-export the slice default so callers wiring an initial setup don't need
 *  to import from the store module. */
export { defaultPageSetup };
