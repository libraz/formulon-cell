import {
  type CellRect,
  type Color,
  ColorKind,
  PageOrientation,
  type PageSetupPatch,
  type Pagination,
  type PrintTitlesInput,
  type Span,
} from '@libraz/formulon';
import { colLetter, formatA1Range, rangeOfRect } from './address.js';
import {
  type EnginePageSetupExtrasInput,
  mergePageSetupFragment,
  readPageSetupExtras,
} from './page-setup-xml.js';
import { attemptWrite, attemptWriteStatus, orFallback, type WriteResult } from './result.js';
import { isChartsheet, sheetHasCells, sheetHasView } from './sheet-kind.js';
import {
  type EngineSheetPr,
  type EngineSheetPrInput,
  mergeSheetPrFragment,
  readSheetPr,
  type SheetTabColorSpec,
} from './sheet-pr-xml.js';
import type {
  EngineCapabilities,
  EngineHeaderFooter,
  EngineOrientation,
  EnginePageBreaks,
  EnginePageMargins,
  EnginePageSetup,
  EnginePageSetupExtras,
  EnginePrintOptions,
} from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';
import { assertAlive, ws } from './workbook-handle-internals.js';

/** The `<tabColor>` spec of an engine colour; null when the tab has none. */
function tabColorSpec(color: Color | null): SheetTabColorSpec | null {
  if (!color) return null;
  const withTint = (spec: SheetTabColorSpec): SheetTabColorSpec => {
    if (color.tint !== 0) spec.tint = color.tint;
    return spec;
  };
  switch (color.kind) {
    case ColorKind.Rgb:
      return { rgb: (color.argb >>> 0).toString(16).toUpperCase().padStart(8, '0') };
    case ColorKind.Theme:
      return withTint({ theme: color.theme });
    case ColorKind.Indexed:
      return { indexed: color.indexed };
    case ColorKind.Auto:
      return { auto: true };
    default:
      return null;
  }
}

/** The engine colour for a `<tabColor>` spec. */
function colorOfSpec(spec: SheetTabColorSpec): Color {
  const base = { argb: 0, theme: 0, tint: 0, indexed: 0 };
  if (spec.rgb !== undefined) {
    return { ...base, kind: ColorKind.Rgb, argb: Number.parseInt(spec.rgb, 16) >>> 0 };
  }
  if (spec.theme !== undefined) {
    return { ...base, kind: ColorKind.Theme, theme: spec.theme, tint: spec.tint ?? 0 };
  }
  if (spec.indexed !== undefined)
    return { ...base, kind: ColorKind.Indexed, indexed: spec.indexed };
  return { ...base, kind: spec.auto ? ColorKind.Auto : ColorKind.None };
}

const sameTabColor = (a: SheetTabColorSpec | null, b: SheetTabColorSpec | null): boolean =>
  a === b ||
  (a !== null &&
    b !== null &&
    a.rgb === b.rgb &&
    a.theme === b.theme &&
    a.indexed === b.indexed &&
    (a.auto ?? false) === (b.auto ?? false) &&
    (a.tint ?? 0) === (b.tint ?? 0));

type WorkbookHandleCtor = { prototype: WorkbookHandle };

/** Partial `<pageSetup>` update. Mirrors the engine's `PageSetupInput`: only
 *  the keys present are applied. `fitToPage` selects the fit-to-page mode
 *  while `fitToWidth` / `fitToHeight` state the target, so "fit onto one
 *  page" needs all three. */
type EnginePageSetupInput = Partial<{
  orientation: EngineOrientation;
  paperSize: number;
  scale: number;
  fitToWidth: number;
  fitToHeight: number;
  fitToPage: boolean;
}>;

/** Partial `<headerFooter>` update. Every key is optional and the section
 *  strings are tri-state — see `setSheetHeaderFooter`. */
type EngineHeaderFooterInput = Partial<{
  oddHeader: string;
  oddFooter: string;
  evenHeader: string;
  evenFooter: string;
  firstHeader: string;
  firstFooter: string;
  differentOddEven: boolean;
  differentFirst: boolean;
  scaleWithDoc: boolean;
  alignWithMargins: boolean;
}>;

declare module './workbook-handle.js' {
  interface WorkbookHandle extends WorkbookHandlePrintMethods {}
}

/** `<pageSetup orientation>` as the store-facing code: 0 absent, 1 portrait, 2 landscape. */
function orientationCodeOf(orientation: PageOrientation | null): EngineOrientation {
  if (orientation === null) return 0;
  return orientation === PageOrientation.Landscape ? 2 : 1;
}

function engineOrientationOf(code: EngineOrientation): PageOrientation | null {
  if (code === 0) return null;
  return code === 2 ? PageOrientation.Landscape : PageOrientation.Portrait;
}

const spanText = (span: Span | null, label: (index: number) => string): string =>
  span === null ? '' : `${label(span.first)}:${label(span.last)}`;

const rowLabel = (row: number): string => String(row + 1);

/** The engine's span-string form of a title axis; empty means no repeat. */
const titleInput = (text: string): string | null => text || null;

export abstract class WorkbookHandlePrintMethods {
  declare sheetKind: WorkbookHandle['sheetKind'];
  declare readonly capabilities: EngineCapabilities;

  /**
   * Native print geometry for `sheet`, including the physical page placements
   * and the provenance of automatic versus manual breaks.  The result is
   * deliberately nullable: older formulon bundles and the development stub
   * do not expose the paginator, so callers retain their existing layout
   * splitter in that case.
   */
  paginateSheet(sheet: number): Pagination | null {
    assertAlive(this);
    if (this.capabilities.detailedPagination !== true) return null;
    if (!sheetHasCells(this, sheet)) return null;
    return ws(this, sheet).getPagination();
  }

  /** Effective page setup for `sheet`, or null when the engine carries no
   *  print-settings surface. */
  getSheetPageSetup(sheet: number): EnginePageSetup | null {
    assertAlive(this);
    if (!this.capabilities.printSettings) return null;
    const r = ws(this, sheet).getPageSetup();
    return {
      orientation: orientationCodeOf(r.orientation),
      paperSize: r.paperSize,
      scale: r.scalePct,
      fitToWidth: r.fitToWidth,
      fitToHeight: r.fitToHeight,
      fitToPage: r.fitToPage,
      orientationStated: r.orientation !== null,
      paperSizeStated: r.paperSizeStated,
      scaleStated: r.scalePctStated,
      fitToPageStated: r.fitToPageStated,
    };
  }

  /** Apply a partial `<pageSetup>` update. Attributes the engine does not
   *  model — `r:id`, `copies` — are left as they were. */
  setSheetPageSetup(sheet: number, setup: EnginePageSetupInput): boolean {
    assertAlive(this);
    if (!this.capabilities.printSettings) return false;
    const { orientation, scale, ...rest } = setup;
    const patch: PageSetupPatch = {
      ...rest,
      ...(orientation !== undefined ? { orientation: engineOrientationOf(orientation) } : {}),
      ...(scale !== undefined ? { scalePct: scale } : {}),
    };
    return attemptWrite(() => ws(this, sheet).updatePageSetup(patch));
  }

  /** The complete worksheet `<pageSetup>` fragment, or null when the engine
   *  does not expose the raw seam. Empty when the sheet declares no page
   *  setup. This is the escape hatch for the attributes the typed setter does
   *  not model — `setSheetPageSetupExtras` is the typed way in. */
  getSheetPageSetupXml(sheet: number): string | null {
    assertAlive(this);
    if (!this.capabilities.printSettingsXml) return null;
    return ws(this, sheet).ooxml.getPageSetupXml() ?? '';
  }

  /** Replace the worksheet `<pageSetup>` fragment. Empty removes it and
   *  restores the defaults. The engine rejects a fragment that is malformed,
   *  oversized, or carries an `r:id` the sheet has no printerSettings part
   *  for, so a bad fragment fails at the call rather than on save. */
  setSheetPageSetupXml(sheet: number, xml: string): boolean {
    assertAlive(this);
    if (!this.capabilities.printSettingsXml) return false;
    return attemptWrite(() => ws(this, sheet).ooxml.setPageSetupXml(xml || null));
  }

  /** The `<pageSetup>` attributes beside the typed setter's own. Read from the
   *  raw fragment, which is the only surface that carries them. */
  getSheetPageSetupExtras(sheet: number): EnginePageSetupExtras | null {
    assertAlive(this);
    if (!this.capabilities.printSettingsXml) return null;
    return readPageSetupExtras(ws(this, sheet).ooxml.getPageSetupXml() ?? '');
  }

  /**
   * Merge attributes into the raw `<pageSetup>` fragment. An omitted key is
   * left alone; `false`, `''` and `null` remove the attribute, which is what a
   * spreadsheet writes for the default rather than stating it.
   *
   * Everything else the fragment carries — including a `r:id` pointing at the
   * sheet's printerSettings part — is preserved, so this is safe to run over a
   * fragment the engine handed back.
   */
  setSheetPageSetupExtras(sheet: number, extras: EnginePageSetupExtrasInput): boolean {
    assertAlive(this);
    if (!this.capabilities.printSettingsXml) return false;
    const sh = ws(this, sheet);
    const current = sh.ooxml.getPageSetupXml() ?? '';
    const next = mergePageSetupFragment(current, extras);
    if (next === null || next === current) return next !== null;
    return attemptWrite(() => sh.ooxml.setPageSetupXml(next || null));
  }

  /** Effective page margins for `sheet`, in inches. */
  getSheetPageMargins(sheet: number): EnginePageMargins | null {
    assertAlive(this);
    if (!this.capabilities.printSettings) return null;
    const r = ws(this, sheet).getPageMargins();
    return {
      left: r.leftIn,
      right: r.rightIn,
      top: r.topIn,
      bottom: r.bottomIn,
      header: r.headerIn,
      footer: r.footerIn,
    };
  }

  /** Apply a partial `<pageMargins>` update, in inches. The engine rejects a
   *  negative, infinite or NaN margin outright, so callers must not pass one. */
  setSheetPageMargins(sheet: number, margins: Partial<EnginePageMargins>): boolean {
    assertAlive(this);
    if (!this.capabilities.printSettings) return false;
    return attemptWrite(() =>
      ws(this, sheet).updatePageMargins({
        leftIn: margins.left,
        rightIn: margins.right,
        topIn: margins.top,
        bottomIn: margins.bottom,
        headerIn: margins.header,
        footerIn: margins.footer,
      }),
    );
  }

  /** Worksheet `<printOptions>` flags. Every flag defaults to false, which is
   *  also what an absent element means. */
  getSheetPrintOptions(sheet: number): EnginePrintOptions | null {
    assertAlive(this);
    if (!this.capabilities.printSettings) return null;
    const r = ws(this, sheet).getPrintOptions();
    return {
      gridLines: r.gridLines,
      headings: r.headings,
      horizontalCentered: r.horizontalCentered,
      verticalCentered: r.verticalCentered,
    };
  }

  setSheetPrintOptions(sheet: number, options: Partial<EnginePrintOptions>): boolean {
    assertAlive(this);
    if (!this.capabilities.printSettings) return false;
    return attemptWrite(() => ws(this, sheet).updatePrintOptions(options));
  }

  /** Worksheet `<headerFooter>` state with the section strings decoded; a
   *  section the sheet lacks reads empty. `scaleWithDoc` and `alignWithMargins`
   *  default to true per OOXML. */
  getSheetHeaderFooter(sheet: number): EngineHeaderFooter | null {
    assertAlive(this);
    if (!this.capabilities.printSettings) return null;
    const r = ws(this, sheet).getHeaderFooter();
    return {
      oddHeader: r.oddHeader ?? '',
      oddFooter: r.oddFooter ?? '',
      evenHeader: r.evenHeader ?? '',
      evenFooter: r.evenFooter ?? '',
      firstHeader: r.firstHeader ?? '',
      firstFooter: r.firstFooter ?? '',
      differentOddEven: r.differentOddEven,
      differentFirst: r.differentFirst,
      scaleWithDoc: r.scaleWithDoc,
      alignWithMargins: r.alignWithMargins,
    };
  }

  /** Apply a partial `<headerFooter>` update. Each section is tri-state: an
   *  omitted key is left alone, `''` clears it. Section text is decoded, so a
   *  literal ampersand is spelled `&&` the way a spreadsheet's header syntax
   *  does; the engine handles the XML escaping. */
  setSheetHeaderFooter(sheet: number, headerFooter: EngineHeaderFooterInput): boolean {
    assertAlive(this);
    if (!this.capabilities.printSettings) return false;
    return attemptWrite(() => ws(this, sheet).updateHeaderFooter(headerFooter));
  }

  /** Tab colour (typed engine accessor, any sheet kind) and outline summary
   *  position (raw `<sheetPr>` fragment), or null when the engine does not
   *  expose it. */
  getSheetPr(sheet: number): EngineSheetPr | null {
    assertAlive(this);
    if (!this.capabilities.sheetPrXml) return null;
    if (!sheetHasView(this, sheet)) return null;
    const sh = ws(this, sheet);
    const readXml = () => sh.ooxml.getSheetPrXml() ?? '';
    // A chartsheet's part may not carry the fragment; that reads as none.
    const xml = isChartsheet(this, sheet) ? orFallback('', readXml) : readXml();
    return { ...readSheetPr(xml), tabColor: tabColorSpec(sh.getTabColor()) };
  }

  /** Write a tab colour and/or outline summary position. The tab colour goes
   *  through the typed setter; the rest of `<sheetPr>` (code name, filter mode,
   *  fit-to-page, ...) is merged and kept. `status` is null when no engine call
   *  was made. */
  setSheetPr(sheet: number, change: EngineSheetPrInput): WriteResult {
    assertAlive(this);
    if (!this.capabilities.sheetPrXml) return { ok: false, status: null };
    const sh = ws(this, sheet);
    const { tabColor, ...outline } = change;
    const setTab =
      tabColor !== undefined && !sameTabColor(tabColorSpec(sh.getTabColor()), tabColor);
    const hasOutline = outline.summaryBelow !== undefined || outline.summaryRight !== undefined;
    if (hasOutline) {
      const current = sh.ooxml.getSheetPrXml() ?? '';
      const next = mergeSheetPrFragment(current, outline);
      if (next === null) return { ok: false, status: null };
      if (next === current && !setTab) return { ok: true };
    } else if (!setTab) {
      return { ok: true };
    }
    // The typed setter rewrites `<sheetPr>`, so the fragment is merged after it.
    return attemptWriteStatus(() => {
      const previous = sh.getTabColor();
      if (setTab) sh.setTabColor(tabColor ? colorOfSpec(tabColor) : null);
      if (!hasOutline) return;
      const fresh = sh.ooxml.getSheetPrXml() ?? '';
      const next = mergeSheetPrFragment(fresh, outline);
      if (next === null || next === fresh) return;
      try {
        sh.ooxml.setSheetPrXml(next || null);
      } catch (err) {
        // Keep the write atomic: a refused outline change takes the tab colour back.
        if (setTab) sh.setTabColor(previous);
        throw err;
      }
    });
  }

  /** `_xlnm.Print_Area` as comma-separated A1 ranges; empty when unset. */
  getSheetPrintArea(sheet: number): string | null {
    assertAlive(this);
    if (!this.capabilities.printSettings) return null;
    const rects: readonly CellRect[] | null = ws(this, sheet).getPrintArea();
    return (rects ?? []).map((r) => formatA1Range(rangeOfRect(sheet, r))).join(',');
  }

  /** Write `_xlnm.Print_Area`. An empty string removes it. */
  setSheetPrintArea(sheet: number, rangesA1: string): boolean {
    assertAlive(this);
    if (!this.capabilities.printSettings) return false;
    const parts = rangesA1
      .split(',')
      .map((part) => part.trim())
      .filter(Boolean);
    return attemptWrite(() => ws(this, sheet).setPrintArea(parts.length > 0 ? parts : null));
  }

  /** `_xlnm.Print_Titles` as a row span and a column span; either may be
   *  empty when that axis has no repeat setting. */
  getSheetPrintTitles(sheet: number): { repeatRows: string; repeatCols: string } | null {
    assertAlive(this);
    if (!this.capabilities.printSettings) return null;
    const r = ws(this, sheet).getPrintTitles();
    return {
      repeatRows: spanText(r?.rows ?? null, rowLabel),
      repeatCols: spanText(r?.columns ?? null, colLetter),
    };
  }

  /** Write `_xlnm.Print_Titles`. Both spans empty removes it. */
  setSheetPrintTitles(sheet: number, repeatRows: string, repeatCols: string): boolean {
    assertAlive(this);
    if (!this.capabilities.printSettings) return false;
    const titles: PrintTitlesInput = {
      rows: titleInput(repeatRows),
      columns: titleInput(repeatCols),
    };
    const none = titles.rows === null && titles.columns === null;
    return attemptWrite(() => ws(this, sheet).setPrintTitles(none ? null : titles));
  }

  /** Manual page breaks on `sheet`. Breaks the engine computed rather than a
   *  user placed are left out — only manual ones are the store's to own. */
  getSheetPageBreaks(sheet: number): EnginePageBreaks | null {
    assertAlive(this);
    if (!this.capabilities.pageBreaks) return null;
    const sh = ws(this, sheet);
    return {
      rows: sh
        .listRowBreaks()
        .filter((b) => b.manual)
        .map((b) => b.row),
      cols: sh
        .listColumnBreaks()
        .filter((b) => b.manual)
        .map((b) => b.col),
    };
  }

  /** Replace every manual break on `sheet` with the supplied ones. Breaks are
   *  0-based indices of the row / column each break precedes. */
  setSheetPageBreaks(sheet: number, breaks: EnginePageBreaks): boolean {
    assertAlive(this);
    if (!this.capabilities.pageBreaks) return false;
    const sh = ws(this, sheet);
    if (!attemptWrite(() => sh.clearBreaks())) return false;
    let ok = true;
    for (const row of breaks.rows)
      ok = attemptWrite(() => sh.setRowBreak(row, { manual: true })) && ok;
    for (const col of breaks.cols) {
      ok = attemptWrite(() => sh.setColumnBreak(col, { manual: true })) && ok;
    }
    return ok;
  }
}

export function installPrintMethods(target: WorkbookHandleCtor): void {
  for (const key of Object.getOwnPropertyNames(WorkbookHandlePrintMethods.prototype)) {
    if (key === 'constructor') continue;
    const descriptor = Object.getOwnPropertyDescriptor(WorkbookHandlePrintMethods.prototype, key);
    if (!descriptor) continue;
    Object.defineProperty(target.prototype, key, descriptor);
  }
}
