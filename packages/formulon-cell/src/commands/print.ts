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
import {
  type CellFormat,
  defaultPageSetup,
  getPageSetup,
  type PageMargins,
  type PageSetup,
  type PrintCellErrorsMode,
  type SpreadsheetStore,
} from '../store/store.js';
import { type PrinterProfile, resolvePrinterProfileBounds } from './printer-profile.js';
import { formatA1FormulaAsR1C1 } from './refs.js';

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
}

export interface PrintSheetOptions {
  printerProfiles?: readonly PrinterProfile[];
  printerProfileId?: string;
}

export interface PrintAreaBounds {
  row0: number;
  col0: number;
  row1: number;
  col1: number;
}

/** Convert a 0-indexed column number to A1 letter form ("A", "B", … "Z",
 *  "AA", …). Used for column-letter headings and print-title parsing. */
export function colLetter(col: number): string {
  let n = col;
  let out = '';
  do {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
}

/** Reverse of `colLetter`. Returns -1 on parse failure. */
function colFromLetters(letters: string): number {
  let col = 0;
  const upper = letters.toUpperCase();
  for (let i = 0; i < upper.length; i += 1) {
    const code = upper.charCodeAt(i);
    if (code < 65 || code > 90) return -1;
    col = col * 26 + (code - 64);
  }
  return col - 1;
}

/** Parse an A1-style row range like "1:3" / "$1:$3" / "2" → `[r0, r1]`
 *  inclusive, 0-indexed. Returns null on bad input. */
export function parsePrintTitleRows(raw?: string): [number, number] | null {
  if (!raw) return null;
  const trimmed = raw.trim().replace(/\$/g, '');
  if (!trimmed) return null;
  const parts = trimmed.split(':');
  const a = Number.parseInt(parts[0] ?? '', 10);
  if (!Number.isFinite(a) || a < 1) return null;
  if (parts.length === 1) return [a - 1, a - 1];
  const b = Number.parseInt(parts[1] ?? '', 10);
  if (!Number.isFinite(b) || b < 1) return null;
  return [Math.min(a, b) - 1, Math.max(a, b) - 1];
}

/** Parse an A1-style col range like "A:B" / "$A:$B" / "C" → `[c0, c1]`
 *  inclusive, 0-indexed. Returns null on bad input. */
export function parsePrintTitleCols(raw?: string): [number, number] | null {
  if (!raw) return null;
  const trimmed = raw.trim().replace(/\$/g, '');
  if (!trimmed) return null;
  const parts = trimmed.split(':');
  const a = colFromLetters(parts[0] ?? '');
  if (a < 0) return null;
  if (parts.length === 1) return [a, a];
  const b = colFromLetters(parts[1] ?? '');
  if (b < 0) return null;
  return [Math.min(a, b), Math.max(a, b)];
}

/** Parse an A1-style rectangular print area like "A1:D20" / "$A$1:$D$20"
 *  / "B2" → zero-indexed bounds. Returns null on bad input. */
export function parsePrintArea(raw?: string): PrintAreaBounds | null {
  if (!raw) return null;
  const trimmed = raw.trim().replace(/\$/g, '');
  if (!trimmed || trimmed.includes(',')) return null;
  const parseCell = (cell: string): { row: number; col: number } | null => {
    const match = /^([A-Za-z]+)([1-9][0-9]*)$/.exec(cell.trim());
    if (!match) return null;
    const col = colFromLetters(match[1] ?? '');
    const row = Number.parseInt(match[2] ?? '', 10) - 1;
    if (col < 0 || row < 0) return null;
    return { row, col };
  };
  const parts = trimmed.split(':');
  if (parts.length > 2) return null;
  const a = parseCell(parts[0] ?? '');
  const b = parseCell(parts[1] ?? parts[0] ?? '');
  if (!a || !b) return null;
  return {
    row0: Math.min(a.row, b.row),
    col0: Math.min(a.col, b.col),
    row1: Math.max(a.row, b.row),
    col1: Math.max(a.col, b.col),
  };
}

export function parsePrintAreas(raw?: string): PrintAreaBounds[] | null {
  if (!raw) return null;
  const parts = raw
    .split(',')
    .map((part) => part.trim())
    .filter(Boolean);
  if (parts.length === 0) return null;
  const areas = parts.map((part) => parsePrintArea(part));
  if (areas.some((area) => area === null)) return null;
  return areas as PrintAreaBounds[];
}

const escapeHtml = (s: string): string =>
  s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');

const cellDisplay = (
  value: CellValue,
  formula: string | null,
  showFormulas: boolean,
  addr?: { row: number; col: number },
  r1c1 = false,
  cellErrorsAs: PrintCellErrorsMode = 'displayed',
): string => {
  if (showFormulas && formula) return r1c1 && addr ? formatA1FormulaAsR1C1(formula, addr) : formula;
  if (value.kind === 'blank') return '';
  if (value.kind === 'error') {
    if (cellErrorsAs === 'blank') return '';
    if (cellErrorsAs === 'dash') return '--';
    if (cellErrorsAs === 'na') return '#N/A';
  }
  return formatCell(value);
};

const fillPatternCss = (pattern: CellFormat['fillPattern'], color = '#000000'): string => {
  switch (pattern) {
    case 'gray125':
      return `radial-gradient(${color} 0.6px, transparent 0.6px)`;
    case 'gray25':
      return `radial-gradient(${color} 1px, transparent 1px)`;
    case 'gray50':
      return `repeating-linear-gradient(45deg, ${color} 0 2px, transparent 2px 4px)`;
    case 'horizontal':
      return `repeating-linear-gradient(0deg, ${color} 0 1px, transparent 1px 4px)`;
    case 'vertical':
      return `repeating-linear-gradient(90deg, ${color} 0 1px, transparent 1px 4px)`;
    case 'diagonalDown':
      return `repeating-linear-gradient(45deg, ${color} 0 1px, transparent 1px 5px)`;
    case 'diagonalUp':
      return `repeating-linear-gradient(135deg, ${color} 0 1px, transparent 1px 5px)`;
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
): string {
  const parts: string[] = [];
  if (showGridlines) {
    // Default hairline grid for unformatted cells. Per-side overrides below.
    parts.push('border:1px solid #c8c8c8');
  }
  if (!fmt) return parts.join(';');
  if (fmt.bold) parts.push('font-weight:600');
  if (fmt.italic) parts.push('font-style:italic');
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
    parts.push(`padding-left:${4 + fmt.indent * 8}px`);
  }
  if (typeof fmt.rotation === 'number' && fmt.rotation !== 0) {
    parts.push(`transform:rotate(${-fmt.rotation}deg);transform-origin:center`);
  }
  if (fmt.textDirection === 'rtl') parts.push('direction:rtl');
  else if (fmt.textDirection === 'ltr') parts.push('direction:ltr');
  const redNegative =
    value.kind === 'number' &&
    value.value < 0 &&
    (fmt.numFmt?.kind === 'fixed' || fmt.numFmt?.kind === 'currency') &&
    (fmt.numFmt.negativeStyle === 'red' || fmt.numFmt.negativeStyle === 'red-parens');
  if (fmt.color && !blackAndWhite) parts.push(`color:${fmt.color}`);
  else if (redNegative && !blackAndWhite) parts.push('color:#c00000');
  if (fmt.fill && !blackAndWhite) parts.push(`background:${fmt.fill}`);
  if (fmt.fillPattern && !blackAndWhite) {
    const image = fillPatternCss(fmt.fillPattern, fmt.fillPatternColor);
    if (image) parts.push(`background-image:${image}`);
    if (fmt.fillPattern === 'gray125' || fmt.fillPattern === 'gray25') {
      parts.push('background-size:4px 4px');
    }
  }
  if (fmt.fontFamily) parts.push(`font-family:${fmt.fontFamily}`);
  if (typeof fmt.fontSize === 'number') parts.push(`font-size:${fmt.fontSize}px`);
  if (fmt.borders) {
    const sideToCss = (side: 'top' | 'right' | 'bottom' | 'left'): void => {
      const b = fmt.borders?.[side];
      if (!b) return;
      if (b === true) {
        parts.push(`border-${side}:1px solid #000`);
        return;
      }
      const style =
        b.style === 'thick'
          ? '2px solid'
          : b.style === 'medium'
            ? '1.5px solid'
            : b.style === 'double'
              ? '3px double'
              : b.style === 'dashed' || b.style === 'mediumDashed' || b.style === 'dashDot'
                ? '1px dashed'
                : b.style === 'dotted'
                  ? '1px dotted'
                  : '1px solid';
      parts.push(`border-${side}:${style} ${blackAndWhite ? '#000' : (b.color ?? '#000')}`);
    };
    sideToCss('top');
    sideToCss('right');
    sideToCss('bottom');
    sideToCss('left');
  }
  return parts.join(';');
}

const PAPER_DIMENSIONS: Record<string, string> = {
  A3: 'A3',
  A4: 'A4',
  A5: 'A5',
  letter: 'letter',
  legal: 'legal',
  tabloid: 'tabloid',
};

export function effectivePrintMargins(setup: PageSetup): PageMargins {
  const printable = setup.printableBounds;
  if (!printable) return { ...setup.margins };
  return {
    top: Math.max(setup.margins.top, printable.top),
    right: Math.max(setup.margins.right, printable.right),
    bottom: Math.max(setup.margins.bottom, printable.bottom),
    left: Math.max(setup.margins.left, printable.left),
  };
}

export interface PrintableMarginAdjustment {
  side: keyof PageMargins;
  margin: number;
  minimum: number;
  effective: number;
}

export function printableMarginAdjustments(setup: PageSetup): PrintableMarginAdjustment[] {
  const printable = setup.printableBounds;
  if (!printable) return [];
  const sides: (keyof PageMargins)[] = ['top', 'right', 'bottom', 'left'];
  return sides
    .map((side) => ({
      side,
      margin: setup.margins[side],
      minimum: printable[side],
      effective: Math.max(setup.margins[side], printable[side]),
    }))
    .filter((item) => item.minimum > item.margin);
}

/** Build the @page CSS string for a given setup. Browsers honour orientation
 *  and (for print preview / PDF) the size keyword. */
function buildPageRule(setup: PageSetup): string {
  const size = `${PAPER_DIMENSIONS[setup.paperSize] ?? 'A4'} ${setup.orientation}`;
  const m = effectivePrintMargins(setup);
  return `@page { size: ${size}; margin: ${m.top}in ${m.right}in ${m.bottom}in ${m.left}in; }`;
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
  const hiddenRows = state.layout.hiddenRows;
  const hiddenCols = state.layout.hiddenCols;

  // Collect populated cells. Track max row/col so the table doesn't run past
  // the last data cell.
  const cellMap = new Map<string, CellSnap>();
  let maxRow = -1;
  let maxCol = -1;
  for (const e of wb.cells(sheet)) {
    const key = `${e.addr.row}:${e.addr.col}`;
    const fmt = state.format.formats.get(`${sheet}:${e.addr.row}:${e.addr.col}`);
    cellMap.set(key, { value: e.value, formula: e.formula, format: fmt });
    if (e.addr.row > maxRow) maxRow = e.addr.row;
    if (e.addr.col > maxCol) maxCol = e.addr.col;
  }
  for (const [formatKey, fmt] of state.format.formats) {
    const [sheetPart, rowPart, colPart] = formatKey.split(':');
    if (Number(sheetPart) !== sheet) continue;
    const row = Number(rowPart);
    const col = Number(colPart);
    if (!Number.isInteger(row) || !Number.isInteger(col)) continue;
    const key = `${row}:${col}`;
    if (!cellMap.has(key)) {
      cellMap.set(key, { value: { kind: 'blank' }, formula: null, format: fmt });
    }
    if (row > maxRow) maxRow = row;
    if (col > maxCol) maxCol = col;
  }

  // Merges shape colspan/rowspan + cell-skip set (cells inside but not at
  // the anchor are absorbed into the anchor's span).
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

  if (maxRow < 0) {
    // Empty sheet — print at least one empty row so the user gets a real page.
    maxRow = 0;
    maxCol = 0;
  }

  const titleRowRange = parsePrintTitleRows(setup.printTitleRows);
  const titleColRange = parsePrintTitleCols(setup.printTitleCols);
  const manualRowBreaks = new Set(setup.manualPageBreakRows ?? []);
  const manualColBreaks = new Set(setup.manualPageBreakCols ?? []);
  const printAreas = parsePrintAreas(setup.printArea);
  const printRegions: PrintAreaBounds[] = (
    printAreas ?? [{ row0: 0, col0: 0, row1: maxRow, col1: maxCol }]
  ).map((area) => {
    const row1 = Math.min(maxRow, area.row1);
    const col1 = Math.min(maxCol, area.col1);
    return {
      row0: area.row0,
      col0: area.col0,
      row1: row1 < area.row0 ? area.row0 : row1,
      col1: col1 < area.col0 ? area.col0 : col1,
    };
  });
  const cellInPrintRegions = (row: number, col: number): boolean =>
    printRegions.some(
      (area) => row >= area.row0 && row <= area.row1 && col >= area.col0 && col <= area.col1,
    );

  const commentEntries: PrintCommentSnap[] =
    setup.comments === 'none'
      ? []
      : [...state.format.formats.entries()]
          .map(([formatKey, fmt]): PrintCommentSnap | null => {
            if (typeof fmt.comment !== 'string' || fmt.comment.length === 0) return null;
            const [sheetPart, rowPart, colPart] = formatKey.split(':');
            if (Number(sheetPart) !== sheet) return null;
            const row = Number(rowPart);
            const col = Number(colPart);
            if (!Number.isInteger(row) || !Number.isInteger(col)) return null;
            if (!cellInPrintRegions(row, col)) return null;
            if (hiddenRows.has(row) || hiddenCols.has(col)) return null;
            return { row, col, text: fmt.comment };
          })
          .filter((entry): entry is PrintCommentSnap => entry !== null)
          .sort((a, b) => a.row - b.row || a.col - b.col);
  const commentByCell = new Map(
    commentEntries.map((entry) => [`${entry.row}:${entry.col}`, entry]),
  );

  // Build column letters used for the optional `<colgroup>` headings strip.
  const renderCell = (
    row: number,
    col: number,
    tag: 'td' | 'th',
    regionMaxCol: number,
    renderedSkipKeys: Set<string>,
  ): string => {
    const key = `${row}:${col}`;
    if (renderedSkipKeys.has(key)) return '';
    const snap = cellMap.get(key);
    const value: CellValue = snap?.value ?? { kind: 'blank' };
    const text = cellDisplay(
      value,
      snap?.formula ?? null,
      showFormulas,
      { row, col },
      state.ui.r1c1 === true,
      setup.cellErrorsAs,
    );
    const span = (() => {
      const mergeSpan = spanByAnchor.get(key);
      if (mergeSpan) return mergeSpan;
      if (snap?.format?.align !== 'centerContinuous' || !text) return undefined;
      let colspan = 1;
      for (let nextCol = col + 1; nextCol <= regionMaxCol; nextCol += 1) {
        if (hiddenCols.has(nextCol)) continue;
        const nextKey = `${row}:${nextCol}`;
        if (renderedSkipKeys.has(nextKey)) break;
        const nextSnap = cellMap.get(nextKey);
        const nextText = cellDisplay(
          nextSnap?.value ?? { kind: 'blank' },
          nextSnap?.formula ?? null,
          showFormulas,
          { row, col: nextCol },
          state.ui.r1c1 === true,
          setup.cellErrorsAs,
        );
        if (nextSnap?.format?.align !== 'centerContinuous' || nextText) break;
        renderedSkipKeys.add(nextKey);
        colspan += 1;
      }
      return colspan > 1 ? { rowspan: 1, colspan } : undefined;
    })();
    const style = inlineCellStyle(
      snap?.format,
      value,
      setup.showGridlines === true,
      setup.blackAndWhite === true,
    );
    const rowspanAttr = span && span.rowspan > 1 ? ` rowspan="${span.rowspan}"` : '';
    const colspanAttr = span && span.colspan > 1 ? ` colspan="${span.colspan}"` : '';
    const styleAttr = style ? ` style="${style}"` : '';
    const comment = setup.comments === 'asDisplayed' ? commentByCell.get(key) : undefined;
    const commentHtml = comment
      ? `<div class="fc-print__comment-note">${escapeHtml(comment.text)}</div>`
      : '';
    const classAttr = manualColBreaks.has(col) ? ' class="fc-print__manual-break-col"' : '';
    return `<${tag}${classAttr}${rowspanAttr}${colspanAttr}${styleAttr}>${escapeHtml(text)}${commentHtml}</${tag}>`;
  };

  const renderRow = (
    row: number,
    isTitle: boolean,
    region: PrintAreaBounds,
    renderedSkipKeys: Set<string>,
  ): string => {
    if (hiddenRows.has(row)) return '';
    const cells: string[] = [];
    if (setup.showHeadings) {
      cells.push(`<th class="fc-print__rowhead">${row + 1}</th>`);
    }
    for (let c = region.col0; c <= region.col1; c += 1) {
      if (hiddenCols.has(c)) continue;
      cells.push(renderCell(row, c, isTitle ? 'th' : 'td', region.col1, renderedSkipKeys));
    }
    const cls = manualRowBreaks.has(row) && !isTitle ? ' class="fc-print__manual-break-row"' : '';
    return `<tr${cls}>${cells.join('')}</tr>`;
  };

  // Optional column-letter heading row. Repeated inside <thead> when print-
  // title rows are configured so it appears at the top of every page.
  const headingRow = (region: PrintAreaBounds): string => {
    if (!setup.showHeadings) return '';
    const cells: string[] = [`<th class="fc-print__corner"></th>`];
    for (let c = region.col0; c <= region.col1; c += 1) {
      if (hiddenCols.has(c)) continue;
      const cls = manualColBreaks.has(c)
        ? 'fc-print__colhead fc-print__manual-break-col'
        : 'fc-print__colhead';
      cells.push(`<th class="${cls}">${colLetter(c)}</th>`);
    }
    return `<tr>${cells.join('')}</tr>`;
  };

  const renderTableSection = (region: PrintAreaBounds, sectionIndex: number): string => {
    const renderedSkipKeys = new Set(skipKeys);
    const theadRows: string[] = [];
    const heading = headingRow(region);
    if (heading) theadRows.push(heading);
    if (titleRowRange) {
      for (let r = titleRowRange[0]; r <= titleRowRange[1]; r += 1) {
        const html = renderRow(r, true, region, renderedSkipKeys);
        if (html) theadRows.push(html);
      }
    }
    const thead = theadRows.length ? `<thead>${theadRows.join('')}</thead>` : '';

    const bodyRows: string[] = [];
    const titleSet = new Set<number>();
    if (titleRowRange) {
      for (let r = titleRowRange[0]; r <= titleRowRange[1]; r += 1) titleSet.add(r);
    }
    for (let r = region.row0; r <= region.row1; r += 1) {
      if (titleSet.has(r)) continue;
      const html = renderRow(r, false, region, renderedSkipKeys);
      if (html) bodyRows.push(html);
    }
    const tbody = `<tbody>${bodyRows.join('')}</tbody>`;

    const colgroup = (() => {
      if (!titleColRange) return '';
      const cols: string[] = [];
      if (setup.showHeadings) cols.push('<col>');
      for (let c = region.col0; c <= region.col1; c += 1) {
        if (hiddenCols.has(c)) continue;
        const isTitle = c >= titleColRange[0] && c <= titleColRange[1];
        cols.push(isTitle ? '<col class="fc-print__title-col">' : '<col>');
      }
      return `<colgroup>${cols.join('')}</colgroup>`;
    })();

    const sectionClass =
      sectionIndex === 0 ? 'fc-print__area' : 'fc-print__area fc-print__area--break';
    return `<section class="${sectionClass}"><table class="fc-print__table"${tableTransform}>${colgroup}${thead}${tbody}</table></section>`;
  };

  // Header / footer strips — spreadsheets paint up to three slots per strip. We
  // emit them as fixed-position rows above/below the table; the @page
  // margins reserve enough whitespace so the print preview doesn't overlap
  // the body content.
  const headerHtml = ((): string => {
    const l = setup.headerLeft ?? '';
    const c = setup.headerCenter ?? '';
    const r = setup.headerRight ?? '';
    if (!l && !c && !r) return '';
    return `<div class="fc-print__header"><span>${escapeHtml(l)}</span><span>${escapeHtml(
      c,
    )}</span><span>${escapeHtml(r)}</span></div>`;
  })();
  const footerHtml = ((): string => {
    const l = setup.footerLeft ?? '';
    const c = setup.footerCenter ?? '';
    const r = setup.footerRight ?? '';
    if (!l && !c && !r) return '';
    return `<div class="fc-print__footer"><span>${escapeHtml(l)}</span><span>${escapeHtml(
      c,
    )}</span><span>${escapeHtml(r)}</span></div>`;
  })();

  const commentsHtml =
    setup.comments === 'endOfSheet' && commentEntries.length
      ? `<section class="fc-print__comments"><h2>Comments and Notes</h2><table><tbody>${commentEntries
          .map(
            (entry) =>
              `<tr><th>${colLetter(entry.col)}${entry.row + 1}</th><td>${escapeHtml(
                entry.text,
              )}</td></tr>`,
          )
          .join('')}</tbody></table></section>`
      : '';

  const hasFitToPages = (setup.fitWidth ?? 0) > 0 || (setup.fitHeight ?? 0) > 0;
  const scale = hasFitToPages ? 1 : setup.scale && setup.scale > 0 ? setup.scale : 1;
  const tableTransform =
    scale === 1 ? '' : ` style="transform:scale(${scale});transform-origin:top left;"`;

  const tablesHtml = printRegions
    .map((region, index) => renderTableSection(region, index))
    .join('');

  const effectiveMargins = effectivePrintMargins(setup);
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
    setup.centerHorizontally ? 'fc-print--center-h' : '',
    setup.centerVertically ? 'fc-print--center-v' : '',
    setup.blackAndWhite ? 'fc-print--black-white' : '',
    setup.draftQuality ? 'fc-print--draft' : '',
    hasFitToPages ? 'fc-print--fit-to-pages' : '',
    setup.alignHeaderFooterWithMargins === false ? 'fc-print--hf-free' : '',
  ]
    .filter(Boolean)
    .join(' ');

  const css = [
    buildPageRule(setup),
    'body { margin: 0; font-family: system-ui, -apple-system, "Segoe UI", sans-serif; font-size: 10pt; color: #111; }',
    'table.fc-print__table { border-collapse: collapse; width: 100%; table-layout: auto; }',
    'table.fc-print__table td, table.fc-print__table th { padding: 2px 6px; vertical-align: bottom; box-sizing: border-box; }',
    'table.fc-print__table thead th { background: #f5f5f5; }',
    '.fc-print__rowhead, .fc-print__colhead, .fc-print__corner { background: #f0f0f0; color: #555; font-weight: 500; text-align: center; min-width: 24px; }',
    '.fc-print__header, .fc-print__footer { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 8px; padding: 4px 0; font-size: 9pt; color: #555; }',
    '.fc-print__header, .fc-print__footer { transform: scale(var(--fc-print-header-footer-scale, 1)); transform-origin: top left; }',
    '.fc-print__header { margin-top: calc(-1 * var(--fc-print-header-margin, 0.3in)); }',
    '.fc-print__footer { margin-bottom: calc(-1 * var(--fc-print-footer-margin, 0.3in)); }',
    '.fc-print--hf-free .fc-print__header, .fc-print--hf-free .fc-print__footer { margin-left: calc(-1 * var(--fc-print-margin-left)); margin-right: calc(-1 * var(--fc-print-margin-right)); }',
    '.fc-print__header > :nth-child(2), .fc-print__footer > :nth-child(2) { text-align: center; }',
    '.fc-print__header > :nth-child(3), .fc-print__footer > :nth-child(3) { text-align: right; }',
    '.fc-print--center-h .fc-print__table { margin-left: auto; margin-right: auto; width: auto; }',
    '.fc-print--center-v { min-height: calc(100vh - var(--fc-print-margin-top) - var(--fc-print-margin-bottom)); display: flex; flex-direction: column; justify-content: center; }',
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
${headerHtml}
${tablesHtml}
${commentsHtml}
${footerHtml}
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
  const doc = buildPrintDocument(wb, store, sheet, title, { printableBounds });
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
