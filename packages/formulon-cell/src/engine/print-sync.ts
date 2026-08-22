import {
  getPageSetup,
  mutators,
  type PageOrientation,
  type PageSetup,
  type PaperSize,
  type PrintCellErrorsMode,
  type PrintCommentsMode,
  type PrintPageOrder,
  type PrintQuality,
  type SpreadsheetStore,
} from '../store/store.js';
import type { EngineOrientation, EnginePageBreaks, EnginePageMargins } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

/**
 * Worksheet print settings: the bridge between the store's `PageSetup` record
 * and the engine's print elements (`<pageSetup>`, `<pageMargins>`,
 * `<printOptions>`, `<headerFooter>`, `_xlnm.Print_Area` / `_xlnm.Print_Titles`
 * and the manual break table).
 *
 * The engine models only what its paginator needs as typed fields; the rest of
 * `<pageSetup>` lives in the raw fragment it keeps as the writer's source of
 * truth. Both are used here, so every Page Setup control that maps to an OOXML
 * attribute round-trips. `printableBounds` is the one exception, and a real
 * one: it records a physical printer's minimum margins, which OOXML has no
 * concept of, so it stays session state driving the browser print document.
 *
 * Sync is idempotent and whole-record: `syncPageSetupToEngine` states every
 * modelled setting for the sheet rather than diffing, so an undo replay or a
 * repeated flush converges on the same file.
 */

/** OOXML `<pageSetup orientation>` codes. 0 leaves the attribute off. */
const ORIENTATION_PORTRAIT: EngineOrientation = 1;
const ORIENTATION_LANDSCAPE: EngineOrientation = 2;

/** OOXML `paperSize` codes for the sheets the store's `PaperSize` names. */
const PAPER_SIZE_CODES: Record<PaperSize, number> = {
  letter: 1,
  tabloid: 3,
  legal: 5,
  A3: 8,
  A4: 9,
  A5: 11,
};

const PAPER_SIZE_FROM_CODE = new Map<number, PaperSize>(
  (Object.entries(PAPER_SIZE_CODES) as [PaperSize, number][]).map(([name, code]) => [code, name]),
);

/** `<pageSetup cellComments>` tokens for the store's comment modes. */
const CELL_COMMENTS_TOKENS: Record<PrintCommentsMode, string> = {
  none: 'none',
  asDisplayed: 'asDisplayed',
  endOfSheet: 'atEnd',
};
const CELL_COMMENTS_FROM_TOKEN = new Map<string, PrintCommentsMode>(
  (Object.entries(CELL_COMMENTS_TOKENS) as [PrintCommentsMode, string][]).map(([mode, token]) => [
    token,
    mode,
  ]),
);

/** `<pageSetup errors>` tokens for the store's cell-error modes. */
const CELL_ERRORS_TOKENS: Record<PrintCellErrorsMode, string> = {
  displayed: 'displayed',
  blank: 'blank',
  dash: 'dash',
  na: 'NA',
};
const CELL_ERRORS_FROM_TOKEN = new Map<string, PrintCellErrorsMode>(
  (Object.entries(CELL_ERRORS_TOKENS) as [PrintCellErrorsMode, string][]).map(([mode, token]) => [
    token,
    mode,
  ]),
);

/** `<pageSetup pageOrder>` tokens. Both spellings are the store's own. */
const PAGE_ORDER_TOKENS: readonly PrintPageOrder[] = ['downThenOver', 'overThenDown'];

/** The store's print-quality steps as `horizontalDpi` / `verticalDpi` values.
 *  `automatic` states no DPI at all, which is what a spreadsheet writes when
 *  the user leaves the printer to decide. */
const PRINT_QUALITY_DPI: Record<PrintQuality, number | null> = {
  automatic: null,
  '300': 300,
  '600': 600,
  '1200': 1200,
};

const printQualityFromDpi = (dpi: number | null): PrintQuality => {
  for (const [quality, value] of Object.entries(PRINT_QUALITY_DPI) as [
    PrintQuality,
    number | null,
  ][]) {
    if (value !== null && value === dpi) return quality;
  }
  return 'automatic';
};

/** Store scale is a fraction (1 = 100%); the engine states a percentage. */
const scaleToPercent = (scale: number): number => Math.round(scale * 100);
const scaleFromPercent = (percent: number): number => percent / 100;

const isFiniteNumber = (value: unknown): value is number =>
  typeof value === 'number' && Number.isFinite(value);

/** A margin the engine will accept: finite and not negative. */
const isPrintableMargin = (value: number | undefined): value is number =>
  isFiniteNumber(value) && value >= 0;

/* ---------- header / footer sections ---------- */

/** The three section markers a header or footer string is built from. */
type HeaderFooterSections = { left: string; center: string; right: string };

/**
 * Split a decoded header/footer string into its left / center / right
 * sections. `&L`, `&C` and `&R` open a section; every other code — `&P`,
 * `&"font,style"`, `&B` — belongs to the text of whichever section is open,
 * and `&&` is a literal ampersand that must not be mistaken for a marker.
 * Text before the first marker is Excel's implicit centre section.
 */
export function splitHeaderFooter(text: string): HeaderFooterSections {
  const out: HeaderFooterSections = { left: '', center: '', right: '' };
  let section: keyof HeaderFooterSections = 'center';
  let i = 0;
  while (i < text.length) {
    if (text[i] === '&' && i + 1 < text.length) {
      const code = text[i + 1];
      if (code === '&') {
        out[section] += '&&';
        i += 2;
        continue;
      }
      if (code === 'L' || code === 'C' || code === 'R') {
        section = code === 'L' ? 'left' : code === 'C' ? 'center' : 'right';
        i += 2;
        continue;
      }
      out[section] += text.slice(i, i + 2);
      i += 2;
      continue;
    }
    out[section] += text[i];
    i += 1;
  }
  return out;
}

/** Rebuild a header/footer string from its sections. Empty sections are left
 *  out, so an all-empty record produces the empty string that clears the
 *  element rather than a bare `&L&C&R`. */
export function joinHeaderFooter(sections: Partial<HeaderFooterSections>): string {
  let out = '';
  if (sections.left) out += `&L${sections.left}`;
  if (sections.center) out += `&C${sections.center}`;
  if (sections.right) out += `&R${sections.right}`;
  return out;
}

/* ---------- hydrate ---------- */

/**
 * Seed the page-setup slice for `sheet` from the engine. Called after a
 * workbook is loaded so a spreadsheet's print settings drive both the print
 * document and the Page Break Preview overlay.
 *
 * No-op when the engine carries no print-settings surface — the store keeps
 * whatever the host or the defaults put there.
 */
export function hydratePageSetupFromEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
): void {
  if (!wb.capabilities.printSettings) return;
  const patch: Partial<PageSetup> = {};

  const setup = wb.getSheetPageSetup(sheet);
  if (setup) {
    if (setup.orientationStated) {
      patch.orientation = setup.orientation === ORIENTATION_LANDSCAPE ? 'landscape' : 'portrait';
    }
    if (setup.paperSizeStated) {
      const named = PAPER_SIZE_FROM_CODE.get(setup.paperSize);
      if (named) {
        patch.paperSize = named;
        patch.paperSizeCode = undefined;
      } else {
        // Keep the code so a save does not re-paper the sheet as A4.
        patch.paperSizeCode = setup.paperSize;
      }
    }
    if (setup.scaleStated) patch.scale = scaleFromPercent(setup.scale);
    if (setup.fitToPage) {
      patch.fitWidth = setup.fitToWidth > 0 ? setup.fitToWidth : undefined;
      patch.fitHeight = setup.fitToHeight > 0 ? setup.fitToHeight : undefined;
    } else if (setup.fitToPageStated) {
      patch.fitWidth = undefined;
      patch.fitHeight = undefined;
    }
  }

  const extras = wb.getSheetPageSetupExtras(sheet);
  if (extras) {
    patch.blackAndWhite = extras.blackAndWhite;
    patch.draftQuality = extras.draft;
    patch.comments = CELL_COMMENTS_FROM_TOKEN.get(extras.cellComments) ?? 'none';
    patch.cellErrorsAs = CELL_ERRORS_FROM_TOKEN.get(extras.errors) ?? 'displayed';
    patch.pageOrder = PAGE_ORDER_TOKENS.includes(extras.pageOrder as PrintPageOrder)
      ? (extras.pageOrder as PrintPageOrder)
      : 'downThenOver';
    // Excel numbers from the start unless the sheet both states a first page
    // and asks for it, so the flag decides whether the number is in force.
    patch.firstPageNumber =
      extras.useFirstPageNumber && extras.firstPageNumber !== null
        ? extras.firstPageNumber
        : undefined;
    patch.printQuality = printQualityFromDpi(extras.horizontalDpi);
  }

  const margins = wb.getSheetPageMargins(sheet);
  if (margins) {
    patch.margins = {
      top: margins.top,
      right: margins.right,
      bottom: margins.bottom,
      left: margins.left,
    };
    patch.headerMargin = margins.header;
    patch.footerMargin = margins.footer;
  }

  const options = wb.getSheetPrintOptions(sheet);
  if (options) {
    patch.showGridlines = options.gridLines;
    patch.showHeadings = options.headings;
    patch.centerHorizontally = options.horizontalCentered;
    patch.centerVertically = options.verticalCentered;
  }

  const headerFooter = wb.getSheetHeaderFooter(sheet);
  if (headerFooter) {
    const header = splitHeaderFooter(headerFooter.oddHeader);
    const footer = splitHeaderFooter(headerFooter.oddFooter);
    patch.headerLeft = header.left || undefined;
    patch.headerCenter = header.center || undefined;
    patch.headerRight = header.right || undefined;
    patch.footerLeft = footer.left || undefined;
    patch.footerCenter = footer.center || undefined;
    patch.footerRight = footer.right || undefined;
    patch.differentOddEvenPages = headerFooter.differentOddEven;
    patch.differentFirstPage = headerFooter.differentFirst;
    patch.scaleHeaderFooterWithDocument = headerFooter.scaleWithDoc;
    patch.alignHeaderFooterWithMargins = headerFooter.alignWithMargins;
  }

  const printArea = wb.getSheetPrintArea(sheet);
  if (printArea !== null) patch.printArea = printArea || undefined;

  const titles = wb.getSheetPrintTitles(sheet);
  if (titles) {
    patch.printTitleRows = titles.repeatRows || undefined;
    patch.printTitleCols = titles.repeatCols || undefined;
  }

  const breaks = wb.getSheetPageBreaks(sheet);
  if (breaks) {
    patch.manualPageBreakRows = breaks.rows.length ? [...breaks.rows] : undefined;
    patch.manualPageBreakCols = breaks.cols.length ? [...breaks.cols] : undefined;
  }

  // Seeding the slice must not read back as a UI edit, or the subscription
  // would immediately write the same settings straight back to the engine.
  wb.withEngineSyncMuted(() => {
    mutators.setPageSetup(store, sheet, patch);
  });
}

/* ---------- writeback ---------- */

/**
 * Push the store's page setup for `sheet` into the engine so it reaches the
 * saved .xlsx. Safe to call repeatedly; each call states the full modelled
 * record. No-op without the engine's print-settings surface.
 */
export function syncPageSetupToEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
): void {
  if (!wb.capabilities.printSettings) return;
  const setup = getPageSetup(store.getState(), sheet);

  const fitWidth = setup.fitWidth ?? 0;
  const fitHeight = setup.fitHeight ?? 0;
  wb.setSheetPageSetup(sheet, {
    orientation: orientationCode(setup.orientation),
    paperSize: setup.paperSizeCode ?? PAPER_SIZE_CODES[setup.paperSize],
    // The engine rejects a scale outside 10..400 rather than clamping it, so
    // an out-of-range store value would fail the whole partial update.
    ...(isFiniteNumber(setup.scale) && setup.scale >= 0.1 && setup.scale <= 4
      ? { scale: scaleToPercent(setup.scale) }
      : {}),
    fitToWidth: fitWidth,
    fitToHeight: fitHeight,
    fitToPage: fitWidth > 0 || fitHeight > 0,
  });

  // The attributes beside the typed setter's own reach the file only through
  // the raw fragment, so a host without that seam keeps them session-only.
  const dpi = PRINT_QUALITY_DPI[setup.printQuality ?? 'automatic'] ?? null;
  const comments = setup.comments ?? 'none';
  const errors = setup.cellErrorsAs ?? 'displayed';
  const pageOrder = setup.pageOrder ?? 'downThenOver';
  wb.setSheetPageSetupExtras(sheet, {
    blackAndWhite: setup.blackAndWhite === true,
    draft: setup.draftQuality === true,
    // A schema default is left unstated, the way a spreadsheet writes it.
    cellComments: comments === 'none' ? '' : CELL_COMMENTS_TOKENS[comments],
    errors: errors === 'displayed' ? '' : CELL_ERRORS_TOKENS[errors],
    pageOrder: pageOrder === 'downThenOver' ? '' : pageOrder,
    firstPageNumber: setup.firstPageNumber ?? null,
    useFirstPageNumber: setup.firstPageNumber !== undefined,
    horizontalDpi: dpi,
    verticalDpi: dpi,
  });

  // A negative, infinite or NaN margin is rejected outright — the paginator
  // subtracts these from the paper — so a bad side is dropped from the patch
  // and left as it was rather than failing the whole update.
  const margins: { -readonly [K in keyof EnginePageMargins]?: number } = {};
  if (isPrintableMargin(setup.margins.left)) margins.left = setup.margins.left;
  if (isPrintableMargin(setup.margins.right)) margins.right = setup.margins.right;
  if (isPrintableMargin(setup.margins.top)) margins.top = setup.margins.top;
  if (isPrintableMargin(setup.margins.bottom)) margins.bottom = setup.margins.bottom;
  if (isPrintableMargin(setup.headerMargin)) margins.header = setup.headerMargin;
  if (isPrintableMargin(setup.footerMargin)) margins.footer = setup.footerMargin;
  if (Object.keys(margins).length > 0) wb.setSheetPageMargins(sheet, margins);

  wb.setSheetPrintOptions(sheet, {
    gridLines: setup.showGridlines === true,
    headings: setup.showHeadings === true,
    horizontalCentered: setup.centerHorizontally === true,
    verticalCentered: setup.centerVertically === true,
  });

  wb.setSheetHeaderFooter(sheet, {
    oddHeader: joinHeaderFooter({
      left: setup.headerLeft ?? '',
      center: setup.headerCenter ?? '',
      right: setup.headerRight ?? '',
    }),
    oddFooter: joinHeaderFooter({
      left: setup.footerLeft ?? '',
      center: setup.footerCenter ?? '',
      right: setup.footerRight ?? '',
    }),
    differentOddEven: setup.differentOddEvenPages === true,
    differentFirst: setup.differentFirstPage === true,
    scaleWithDoc: setup.scaleHeaderFooterWithDocument !== false,
    alignWithMargins: setup.alignHeaderFooterWithMargins !== false,
  });

  wb.setSheetPrintArea(sheet, setup.printArea ?? '');
  wb.setSheetPrintTitles(sheet, setup.printTitleRows ?? '', setup.printTitleCols ?? '');

  if (wb.capabilities.pageBreaks) {
    const breaks: EnginePageBreaks = {
      rows: setup.manualPageBreakRows ?? [],
      cols: setup.manualPageBreakCols ?? [],
    };
    wb.setSheetPageBreaks(sheet, breaks);
  }
}

const orientationCode = (orientation: PageOrientation): EngineOrientation =>
  orientation === 'landscape' ? ORIENTATION_LANDSCAPE : ORIENTATION_PORTRAIT;
