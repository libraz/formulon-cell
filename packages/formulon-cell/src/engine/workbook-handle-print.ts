import {
  type EnginePageSetupExtrasInput,
  mergePageSetupFragment,
  readHeaderFooter,
  readPageSetupExtras,
  readPrintOptions,
} from './page-setup-xml.js';
import type {
  EngineCapabilities,
  EngineHeaderFooter,
  EngineOrientation,
  EnginePageBreaks,
  EnginePageMargins,
  EnginePageSetup,
  EnginePageSetupExtras,
  EnginePrintOptions,
  Workbook,
} from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

type WorkbookHandleCtor = { prototype: WorkbookHandle };
type WorkbookHandleInternals = {
  wb: Workbook;
  assertAlive(): void;
};

/** Partial `<pageSetup>` update. Mirrors the engine's `PageSetupInput`: only
 *  the keys present are applied. `fitToPage` selects the fit-to-page mode
 *  while `fitToWidth` / `fitToHeight` state the target, so "fit onto one
 *  page" needs all three. */
export type EnginePageSetupInput = Partial<{
  orientation: EngineOrientation;
  paperSize: number;
  scale: number;
  fitToWidth: number;
  fitToHeight: number;
  fitToPage: boolean;
}>;

/** Partial `<headerFooter>` update. Every key is optional and the section
 *  strings are tri-state — see `setSheetHeaderFooter`. */
export type EngineHeaderFooterInput = Partial<{
  oddHeader: string;
  oddFooter: string;
  differentOddEven: boolean;
  differentFirst: boolean;
  scaleWithDoc: boolean;
  alignWithMargins: boolean;
}>;

declare module './workbook-handle.js' {
  interface WorkbookHandle extends WorkbookHandlePrintMethods {}
}

function internals(handle: unknown): WorkbookHandleInternals {
  return handle as WorkbookHandleInternals;
}

function assertAlive(handle: unknown): void {
  internals(handle).assertAlive();
}

function wb(handle: unknown): Workbook {
  return internals(handle).wb;
}

export abstract class WorkbookHandlePrintMethods {
  declare readonly capabilities: EngineCapabilities;

  /** Effective page setup for `sheet`, or null when the engine carries no
   *  print-settings surface. */
  getSheetPageSetup(sheet: number): EnginePageSetup | null {
    assertAlive(this);
    if (!this.capabilities.printSettings) return null;
    const r = wb(this).getSheetPageSetup(sheet);
    if (!r.status.ok) return null;
    return {
      orientation: r.orientation,
      paperSize: r.paperSize,
      scale: r.scale,
      fitToWidth: r.fitToWidth,
      fitToHeight: r.fitToHeight,
      fitToPage: r.fitToPage,
      orientationStated: r.orientationStated,
      paperSizeStated: r.paperSizeStated,
      scaleStated: r.scaleStated,
      fitToPageStated: r.fitToPageStated,
    };
  }

  /** Apply a partial `<pageSetup>` update. Attributes the engine does not
   *  model — `r:id`, `copies` — are left as they were. */
  setSheetPageSetup(sheet: number, setup: EnginePageSetupInput): boolean {
    assertAlive(this);
    if (!this.capabilities.printSettings) return false;
    return wb(this).setSheetPageSetup(sheet, setup).ok;
  }

  /** The complete worksheet `<pageSetup>` fragment, or null when the engine
   *  does not expose the raw seam. Empty when the sheet declares no page
   *  setup. This is the escape hatch for the attributes the typed setter does
   *  not model — `setSheetPageSetupExtras` is the typed way in. */
  getSheetPageSetupXml(sheet: number): string | null {
    assertAlive(this);
    if (!this.capabilities.printSettingsXml) return null;
    const r = wb(this).getSheetPageSetupXml(sheet);
    return r.status.ok ? r.xml : null;
  }

  /** Replace the worksheet `<pageSetup>` fragment. Empty removes it and
   *  restores the defaults. The engine rejects a fragment that is malformed,
   *  oversized, or carries an `r:id` the sheet has no printerSettings part
   *  for, so a bad fragment fails at the call rather than on save. */
  setSheetPageSetupXml(sheet: number, xml: string): boolean {
    assertAlive(this);
    if (!this.capabilities.printSettingsXml) return false;
    return wb(this).setSheetPageSetupXml(sheet, xml).ok;
  }

  /** The `<pageSetup>` attributes beside the typed setter's own. Read from the
   *  raw fragment, which is the only surface that carries them. */
  getSheetPageSetupExtras(sheet: number): EnginePageSetupExtras | null {
    assertAlive(this);
    if (!this.capabilities.printSettingsXml) return null;
    const r = wb(this).getSheetPageSetupXml(sheet);
    if (!r.status.ok) return null;
    return readPageSetupExtras(r.xml);
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
    const current = wb(this).getSheetPageSetupXml(sheet);
    if (!current.status.ok) return false;
    const next = mergePageSetupFragment(current.xml, extras);
    if (next === null || next === current.xml) return next !== null;
    return wb(this).setSheetPageSetupXml(sheet, next).ok;
  }

  /** Effective page margins for `sheet`, in inches. */
  getSheetPageMargins(sheet: number): EnginePageMargins | null {
    assertAlive(this);
    if (!this.capabilities.printSettings) return null;
    const r = wb(this).getSheetPageMargins(sheet);
    if (!r.status.ok) return null;
    return {
      left: r.left,
      right: r.right,
      top: r.top,
      bottom: r.bottom,
      header: r.header,
      footer: r.footer,
    };
  }

  /** Apply a partial `<pageMargins>` update, in inches. The engine rejects a
   *  negative, infinite or NaN margin outright, so callers must not pass one. */
  setSheetPageMargins(sheet: number, margins: Partial<EnginePageMargins>): boolean {
    assertAlive(this);
    if (!this.capabilities.printSettings) return false;
    return wb(this).setSheetPageMargins(sheet, margins).ok;
  }

  /** Worksheet `<printOptions>` flags. The engine exposes this element only as
   *  raw XML, so the fragment is parsed here rather than in the sync layer —
   *  the raw engine surface stops at this class. Every flag defaults to false,
   *  which is also what an absent element means. */
  getSheetPrintOptions(sheet: number): EnginePrintOptions | null {
    assertAlive(this);
    if (!this.capabilities.printSettings) return null;
    const r = wb(this).getSheetPrintOptionsXml(sheet);
    if (!r.status.ok) return null;
    return readPrintOptions(r.xml);
  }

  setSheetPrintOptions(sheet: number, options: Partial<EnginePrintOptions>): boolean {
    assertAlive(this);
    if (!this.capabilities.printSettings) return false;
    return wb(this).setSheetPrintOptions(sheet, options).ok;
  }

  /** Worksheet `<headerFooter>` state with the section strings decoded. Parsed
   *  from the raw fragment for the same reason as `getSheetPrintOptions`.
   *  `scaleWithDoc` and `alignWithMargins` default to true per OOXML. */
  getSheetHeaderFooter(sheet: number): EngineHeaderFooter | null {
    assertAlive(this);
    if (!this.capabilities.printSettings) return null;
    const r = wb(this).getSheetHeaderFooterXml(sheet);
    if (!r.status.ok) return null;
    return readHeaderFooter(r.xml);
  }

  /** Apply a partial `<headerFooter>` update. Each section is tri-state: an
   *  omitted key is left alone, `''` clears it. Section text is decoded, so a
   *  literal ampersand is spelled `&&` the way a spreadsheet's header syntax
   *  does; the engine handles the XML escaping. */
  setSheetHeaderFooter(sheet: number, headerFooter: EngineHeaderFooterInput): boolean {
    assertAlive(this);
    if (!this.capabilities.printSettings) return false;
    return wb(this).setSheetHeaderFooter(sheet, headerFooter).ok;
  }

  /** `_xlnm.Print_Area` as comma-separated A1 ranges; empty when unset. */
  getSheetPrintArea(sheet: number): string | null {
    assertAlive(this);
    if (!this.capabilities.printSettings) return null;
    const r = wb(this).getSheetPrintArea(sheet);
    return r.status.ok ? r.ranges : null;
  }

  /** Write `_xlnm.Print_Area`. An empty string removes it. */
  setSheetPrintArea(sheet: number, rangesA1: string): boolean {
    assertAlive(this);
    if (!this.capabilities.printSettings) return false;
    return wb(this).setSheetPrintArea(sheet, rangesA1).ok;
  }

  /** `_xlnm.Print_Titles` as a row span and a column span; either may be
   *  empty when that axis has no repeat setting. */
  getSheetPrintTitles(sheet: number): { repeatRows: string; repeatCols: string } | null {
    assertAlive(this);
    if (!this.capabilities.printSettings) return null;
    const r = wb(this).getSheetPrintTitles(sheet);
    if (!r.status.ok) return null;
    return { repeatRows: r.repeatRows, repeatCols: r.repeatCols };
  }

  /** Write `_xlnm.Print_Titles`. Both spans empty removes it. */
  setSheetPrintTitles(sheet: number, repeatRows: string, repeatCols: string): boolean {
    assertAlive(this);
    if (!this.capabilities.printSettings) return false;
    return wb(this).setSheetPrintTitles(sheet, repeatRows, repeatCols).ok;
  }

  /** Manual page breaks on `sheet`. Breaks the engine computed rather than a
   *  user placed are left out — only manual ones are the store's to own. */
  getSheetPageBreaks(sheet: number): EnginePageBreaks | null {
    assertAlive(this);
    if (!this.capabilities.pageBreaks) return null;
    const rows = wb(this).getSheetRowBreaks(sheet);
    const cols = wb(this).getSheetColBreaks(sheet);
    if (!rows.status.ok || !cols.status.ok) return null;
    return {
      rows: rows.breaks.filter((b) => b.manual).map((b) => b.id),
      cols: cols.breaks.filter((b) => b.manual).map((b) => b.id),
    };
  }

  /** Replace every manual break on `sheet` with the supplied ones. Breaks are
   *  0-based indices of the row / column each break precedes. */
  setSheetPageBreaks(sheet: number, breaks: EnginePageBreaks): boolean {
    assertAlive(this);
    if (!this.capabilities.pageBreaks) return false;
    const w = wb(this);
    if (!w.clearSheetBreaks(sheet).ok) return false;
    let ok = true;
    for (const row of breaks.rows) ok = w.addSheetRowBreak(sheet, row, true).ok && ok;
    for (const col of breaks.cols) ok = w.addSheetColBreak(sheet, col, true).ok && ok;
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
