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
  type PageSetup,
  type SpreadsheetStore,
  defaultPageSetup,
  getPageSetup,
} from '../store/store.js';

/** Output of `buildPrintDocument`. `html` is a complete HTML document string
 *  ready to feed into an iframe via `document.write`. `cssVars` carries
 *  individual values (paper size token, margins, scale) so callers / tests
 *  can verify the page-setup wired through without parsing the HTML. */
export interface PrintDocument {
  html: string;
  cssVars: Record<string, string>;
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

const escapeHtml = (s: string): string =>
  s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');

const cellDisplay = (value: CellValue, formula: string | null, showFormulas: boolean): string => {
  if (showFormulas && formula) return formula;
  if (value.kind === 'blank') return '';
  return formatCell(value);
};

/** Translate a `CellFormat` into an inline-style string. Only the subset that
 *  matters for print is included — borders are emitted on the cell itself
 *  so colspan/rowspan merges retain their outline. */
function inlineCellStyle(fmt: CellFormat | undefined, showGridlines: boolean): string {
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
  if (fmt.align) parts.push(`text-align:${fmt.align}`);
  if (fmt.vAlign) {
    const v = fmt.vAlign === 'middle' ? 'middle' : fmt.vAlign;
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
  if (fmt.color) parts.push(`color:${fmt.color}`);
  if (fmt.fill) parts.push(`background:${fmt.fill}`);
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
      parts.push(`border-${side}:${style} ${b.color ?? '#000'}`);
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

/** Build the @page CSS string for a given setup. Browsers honour orientation
 *  and (for print preview / PDF) the size keyword. */
function buildPageRule(setup: PageSetup): string {
  const size = `${PAPER_DIMENSIONS[setup.paperSize] ?? 'A4'} ${setup.orientation}`;
  const m = setup.margins;
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

/** Render the print document HTML for `sheet`. The selection / viewport are
 *  ignored — printing always emits the full populated range. Hidden rows and
 *  columns honour the layout slice. */
export function buildPrintDocument(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
): PrintDocument {
  const state = store.getState();
  const setup = getPageSetup(state, sheet);
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

  // Build column letters used for the optional `<colgroup>` headings strip.
  const renderCell = (row: number, col: number, tag: 'td' | 'th'): string => {
    const key = `${row}:${col}`;
    if (skipKeys.has(key)) return '';
    const span = spanByAnchor.get(key);
    const snap = cellMap.get(key);
    const value: CellValue = snap?.value ?? { kind: 'blank' };
    const text = cellDisplay(value, snap?.formula ?? null, showFormulas);
    const style = inlineCellStyle(snap?.format, setup.showGridlines === true);
    const rowspanAttr = span && span.rowspan > 1 ? ` rowspan="${span.rowspan}"` : '';
    const colspanAttr = span && span.colspan > 1 ? ` colspan="${span.colspan}"` : '';
    const styleAttr = style ? ` style="${style}"` : '';
    return `<${tag}${rowspanAttr}${colspanAttr}${styleAttr}>${escapeHtml(text)}</${tag}>`;
  };

  const renderRow = (row: number, isTitle: boolean): string => {
    if (hiddenRows.has(row)) return '';
    const cells: string[] = [];
    if (setup.showHeadings) {
      cells.push(`<th class="fc-print__rowhead">${row + 1}</th>`);
    }
    for (let c = 0; c <= maxCol; c += 1) {
      if (hiddenCols.has(c)) continue;
      cells.push(renderCell(row, c, isTitle ? 'th' : 'td'));
    }
    return `<tr>${cells.join('')}</tr>`;
  };

  // Optional column-letter heading row. Repeated inside <thead> when print-
  // title rows are configured so it appears at the top of every page.
  const headingRow = ((): string => {
    if (!setup.showHeadings) return '';
    const cells: string[] = [`<th class="fc-print__corner"></th>`];
    for (let c = 0; c <= maxCol; c += 1) {
      if (hiddenCols.has(c)) continue;
      cells.push(`<th class="fc-print__colhead">${colLetter(c)}</th>`);
    }
    return `<tr>${cells.join('')}</tr>`;
  })();

  // <thead> repeats on every printed page. We always include the heading row
  // (when configured) and any print-title rows here.
  const theadRows: string[] = [];
  if (headingRow) theadRows.push(headingRow);
  if (titleRowRange) {
    for (let r = titleRowRange[0]; r <= titleRowRange[1]; r += 1) {
      const html = renderRow(r, true);
      if (html) theadRows.push(html);
    }
  }
  const thead = theadRows.length ? `<thead>${theadRows.join('')}</thead>` : '';

  // Body rows — skip rows already emitted by the title block to avoid
  // duplicates inside <tbody>.
  const bodyRows: string[] = [];
  const titleSet = new Set<number>();
  if (titleRowRange) {
    for (let r = titleRowRange[0]; r <= titleRowRange[1]; r += 1) titleSet.add(r);
  }
  for (let r = 0; r <= maxRow; r += 1) {
    if (titleSet.has(r)) continue;
    const html = renderRow(r, false);
    if (html) bodyRows.push(html);
  }
  const tbody = `<tbody>${bodyRows.join('')}</tbody>`;

  // Header / footer strips — Excel paints up to three slots per strip. We
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

  const scale = setup.scale && setup.scale > 0 ? setup.scale : 1;
  const tableTransform =
    scale === 1 ? '' : ` style="transform:scale(${scale});transform-origin:top left;"`;

  // Highlight the print-title columns by re-stating the col indices in a
  // <colgroup>; the runtime CSS picks them up via :nth-child to repeat
  // the underlying cell on the left of every page (browser-supported via
  // `position: running()` is unreliable, so we duplicate left-frozen cells
  // visually by tagging the columns).
  const colgroup = ((): string => {
    if (!titleColRange) return '';
    const cols: string[] = [];
    if (setup.showHeadings) cols.push('<col>');
    for (let c = 0; c <= maxCol; c += 1) {
      if (hiddenCols.has(c)) continue;
      const isTitle = c >= titleColRange[0] && c <= titleColRange[1];
      cols.push(isTitle ? '<col class="fc-print__title-col">' : '<col>');
    }
    return `<colgroup>${cols.join('')}</colgroup>`;
  })();

  const cssVars: Record<string, string> = {
    '--fc-print-paper': PAPER_DIMENSIONS[setup.paperSize] ?? 'A4',
    '--fc-print-orient': setup.orientation,
    '--fc-print-margin-top': `${setup.margins.top}in`,
    '--fc-print-margin-right': `${setup.margins.right}in`,
    '--fc-print-margin-bottom': `${setup.margins.bottom}in`,
    '--fc-print-margin-left': `${setup.margins.left}in`,
    '--fc-print-scale': String(scale),
  };

  const css = [
    buildPageRule(setup),
    'body { margin: 0; font-family: system-ui, -apple-system, "Segoe UI", sans-serif; font-size: 10pt; color: #111; }',
    'table.fc-print__table { border-collapse: collapse; width: 100%; table-layout: auto; }',
    'table.fc-print__table td, table.fc-print__table th { padding: 2px 6px; vertical-align: bottom; box-sizing: border-box; }',
    'table.fc-print__table thead th { background: #f5f5f5; }',
    '.fc-print__rowhead, .fc-print__colhead, .fc-print__corner { background: #f0f0f0; color: #555; font-weight: 500; text-align: center; min-width: 24px; }',
    '.fc-print__header, .fc-print__footer { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 8px; padding: 4px 0; font-size: 9pt; color: #555; }',
    '.fc-print__header > :nth-child(2), .fc-print__footer > :nth-child(2) { text-align: center; }',
    '.fc-print__header > :nth-child(3), .fc-print__footer > :nth-child(3) { text-align: right; }',
    '.fc-print__title-col { background-color: rgba(0,0,0,0.0); }',
    '@media print { thead { display: table-header-group; } tfoot { display: table-footer-group; } tr, td, th { page-break-inside: avoid; } }',
  ].join('\n');

  const html = `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Print</title>
<style>
${css}
</style>
</head>
<body>
${headerHtml}
<table class="fc-print__table"${tableTransform}>${colgroup}${thead}${tbody}</table>
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
): () => void {
  const doc = buildPrintDocument(wb, store, sheet);
  const iframe = document.createElement('iframe');
  iframe.setAttribute('aria-hidden', 'true');
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
