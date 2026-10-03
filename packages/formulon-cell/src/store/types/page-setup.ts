/** Page orientation for print / PDF export. */
export type PageOrientation = 'portrait' | 'landscape';

/** Paper size — covers the common ISO + ANSI sheets. The print document
 *  emits `@page { size: <paperSize> <orientation> }` which all major browsers
 *  honour for the print preview / PDF rendering. */
export type PaperSize = 'A4' | 'A3' | 'A5' | 'letter' | 'legal' | 'tabloid';

/** Margins in inches — spreadsheet parity. The dialog renders text inputs in inches;
 *  the print-CSS converts to `in` units verbatim. */
export interface PageMargins {
  top: number;
  right: number;
  bottom: number;
  left: number;
}

export type PrintCommentsMode = 'none' | 'asDisplayed' | 'endOfSheet';
export type PrintCellErrorsMode = 'displayed' | 'blank' | 'dash' | 'na';
export type PrintPageOrder = 'downThenOver' | 'overThenDown';
export type PrintQuality = 'automatic' | '300' | '600' | '1200';

/** Per-sheet page-setup configuration. Drives both the Page Setup dialog and
 *  the print document builder. Default values come from `defaultPageSetup()`;
 *  unset fields fall back to that default — `getPageSetup` always returns a
 *  fully-populated record. */
export interface PageSetup {
  orientation: PageOrientation;
  paperSize: PaperSize;
  /** OOXML `paperSize` code for a sheet whose paper is outside `PaperSize`
   *  (B4, Executive, a printer-specific form). Set on hydration and written
   *  back verbatim so loading and saving a workbook does not silently
   *  re-paper it as A4; picking a size from the UI clears it. */
  paperSizeCode?: number;
  margins: PageMargins;
  /** Minimum printable insets from the physical page edge, in inches. This is
   *  distinct from `printArea`: hosts may fill it from a printer profile or
   *  preview preset so content is laid out inside the device's non-printable
   *  border. Browser print cannot discover this automatically. */
  printableBounds?: PageMargins;
  /** Distance from page edge to header/footer text, in inches. */
  headerMargin?: number;
  footerMargin?: number;
  /** Center printed content within the page margins. */
  centerHorizontally?: boolean;
  centerVertically?: boolean;
  /** Header / footer text — desktop spreadsheets splits the strip into three slots
   *  (left / center / right). Empty / missing strings render as nothing. */
  headerLeft?: string;
  headerCenter?: string;
  headerRight?: string;
  footerLeft?: string;
  footerCenter?: string;
  footerRight?: string;
  /** Header/Footer tab options. */
  differentOddEvenPages?: boolean;
  differentFirstPage?: boolean;
  scaleHeaderFooterWithDocument?: boolean;
  alignHeaderFooterWithMargins?: boolean;
  /** A1-style print area, e.g. "A1:D20" or "A1:B2,D4:E5".
   *  Empty means print the used range. */
  printArea?: string;
  /** A1-style row range ("1:3" or "$1:$3") whose rows repeat at the top of
   *  every printed page. Single-row form ("2") is allowed. */
  printTitleRows?: string;
  /** A1-style column range ("A:B"). Repeats those columns on the left of
   *  every printed page. */
  printTitleCols?: string;
  /** Fit-to-N-pages-wide. 0 means no width constraint. */
  fitWidth?: number;
  /** Fit-to-N-pages-tall. 0 means no height constraint. */
  fitHeight?: number;
  /** Manual page breaks before the given zero-based rows / columns. */
  manualPageBreakRows?: number[];
  manualPageBreakCols?: number[];
  /** Print scale, 0.10..4.00 (1 = 100%). When `fitWidth`/`fitHeight` is set
   *  the browser ignores the explicit scale. */
  scale?: number;
  /** Printer quality and first printed page number from Excel's Page tab.
   *  `firstPageNumber` undefined means Auto. */
  printQuality?: PrintQuality;
  firstPageNumber?: number;
  /** Paint inter-cell hairline gridlines on the print document. */
  showGridlines?: boolean;
  /** Paint row-numbers and column-letters on the print document. */
  showHeadings?: boolean;
  /** Excel Sheet tab print options. Some are preserved for parity even when
   *  browser print has no exact equivalent. */
  blackAndWhite?: boolean;
  draftQuality?: boolean;
  comments?: PrintCommentsMode;
  cellErrorsAs?: PrintCellErrorsMode;
  pageOrder?: PrintPageOrder;
}

/** Default page-setup record. Returned by `getPageSetup` when the sheet has
 *  no explicit entry, and used as the baseline for partial-patch merges.
 *
 *  Margins match the "Normal" preset surfaced by the Page Setup dialog
 *  (`commands/page-setup.ts`) so the chrome can faithfully reflect the
 *  active preset in its dropdown. */
export function defaultPageSetup(): PageSetup {
  return {
    orientation: 'portrait',
    paperSize: 'A4',
    margins: { top: 0.75, right: 0.7, bottom: 0.75, left: 0.7 },
    headerMargin: 0.3,
    footerMargin: 0.3,
    centerHorizontally: false,
    centerVertically: false,
    differentOddEvenPages: false,
    differentFirstPage: false,
    scaleHeaderFooterWithDocument: true,
    alignHeaderFooterWithMargins: true,
    scale: 1,
    printQuality: 'automatic',
    showGridlines: false,
    showHeadings: false,
    blackAndWhite: false,
    draftQuality: false,
    comments: 'none',
    cellErrorsAs: 'displayed',
    pageOrder: 'downThenOver',
  };
}
