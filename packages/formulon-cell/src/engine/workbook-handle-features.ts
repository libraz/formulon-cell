import {
  ENGINE_SPREADSHEET_PROFILE_GETTER,
  ENGINE_SPREADSHEET_PROFILE_SETTER,
} from './capabilities.js';
import { computeNamedCellStyles, type NamedCellStyle } from './cell-styles-meta.js';
import { localeTag, mergeFunctionMetadata } from './function-metadata.js';
import { computeEngineSpillRanges } from './spill.js';
import {
  type EngineSpreadsheetProfileId,
  engineProfileToPublic,
  publicProfileToEngine,
} from './spreadsheet-profile.js';
import {
  completeBorderRecord,
  completeDxfRecord,
  completeFillRecord,
  completeFontRecord,
} from './style-records.js';
import type {
  Addr,
  BorderRecord,
  CellValue,
  CellXf,
  ConditionalFormatEntry,
  ConditionalFormatInput,
  DxfRecord,
  EngineCapabilities,
  EngineHeaderFooter,
  EngineOrientation,
  EnginePageBreaks,
  EnginePageMargins,
  EnginePageSetup,
  EnginePageSetupExtras,
  EnginePrintOptions,
  EvalArrayResult,
  EvalResult,
  FillRecord,
  FontRecord,
  FormulonModule,
  FunctionMetadataProvider,
  Range,
  SpreadsheetProfileId,
  TableInput,
  Workbook,
} from './types.js';
import { SheetVisibility } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

/** `setCellStyle`'s sentinel for "this style has no OOXML built-in id". The
 *  binding takes a fixed argument count, so the absence has to be a value. */
const CUSTOM_CELL_STYLE_BUILTIN_ID = 0xffffffff;

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

/** Partial `<pageSetup>` attribute update. An omitted key is left alone;
 *  `false`, `''` and `null` remove the attribute. */
export type EnginePageSetupExtrasInput = Partial<{
  blackAndWhite: boolean;
  draft: boolean;
  cellComments: string;
  errors: string;
  pageOrder: string;
  firstPageNumber: number | null;
  useFirstPageNumber: boolean;
  horizontalDpi: number | null;
  verticalDpi: number | null;
}>;

/** Wrapper element used to parse a bare fragment. It declares the
 *  relationships prefix so a `<pageSetup r:id="...">` — which the engine hands
 *  back without a declaration of its own, because the worksheet root carried
 *  one — parses instead of failing as an undeclared prefix. */
const FRAGMENT_WRAPPER = 'fc-fragment';
const RELATIONSHIPS_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** Parse a `<pageSetup>` fragment inside the namespace-declaring wrapper.
 *  Returns null for the empty string the engine hands back when the sheet
 *  declares no page setup, and for a host with no XML parser. */
function parsePageSetupFragment(xml: string): Element | null {
  if (!xml || typeof DOMParser === 'undefined') return null;
  const doc = new DOMParser().parseFromString(
    `<${FRAGMENT_WRAPPER} xmlns:r="${RELATIONSHIPS_NS}">${xml}</${FRAGMENT_WRAPPER}>`,
    'application/xml',
  );
  const root = doc.documentElement;
  if (!root || root.nodeName !== FRAGMENT_WRAPPER) return null;
  if (root.getElementsByTagName('parsererror').length > 0) return null;
  const el = root.firstElementChild;
  return el?.nodeName === 'pageSetup' ? el : null;
}

/**
 * Apply `extras` to a `<pageSetup>` fragment and return the new fragment.
 *
 * Returns null when the fragment cannot be parsed or the host has no XML
 * serializer, so the caller can tell "nothing to do" from "could not do it".
 * The wrapper is serialized and then unwrapped rather than serializing the
 * element on its own: that keeps the `xmlns:r` declaration on the wrapper,
 * where the engine's own fragment does not carry it either.
 */
export function mergePageSetupFragment(
  xml: string,
  extras: EnginePageSetupExtrasInput,
): string | null {
  if (typeof DOMParser === 'undefined' || typeof XMLSerializer === 'undefined') return null;
  const el = parsePageSetupFragment(xml || '<pageSetup/>');
  if (!el) return null;
  const put = (name: string, value: string | null): void => {
    if (value === null || value === '') el.removeAttribute(name);
    else el.setAttribute(name, value);
  };
  const putFlag = (name: string, value: boolean | undefined): void => {
    if (value === undefined) return;
    put(name, value ? '1' : null);
  };
  const putUint = (name: string, value: number | null | undefined): void => {
    if (value === undefined) return;
    put(name, value === null || !Number.isFinite(value) ? null : String(Math.trunc(value)));
  };
  putFlag('blackAndWhite', extras.blackAndWhite);
  putFlag('draft', extras.draft);
  putFlag('useFirstPageNumber', extras.useFirstPageNumber);
  if (extras.cellComments !== undefined) put('cellComments', extras.cellComments || null);
  if (extras.errors !== undefined) put('errors', extras.errors || null);
  if (extras.pageOrder !== undefined) put('pageOrder', extras.pageOrder || null);
  putUint('firstPageNumber', extras.firstPageNumber);
  putUint('horizontalDpi', extras.horizontalDpi);
  putUint('verticalDpi', extras.verticalDpi);

  const wrapped = new XMLSerializer().serializeToString(el.parentNode as Element);
  const open = wrapped.indexOf('>');
  const close = wrapped.lastIndexOf(`</${FRAGMENT_WRAPPER}>`);
  if (open < 0 || close < open) return null;
  return wrapped.slice(open + 1, close);
}

/** Read an unsigned integer attribute; null when absent or unparseable. */
function xmlUint(el: Element | null, name: string): number | null {
  const raw = el?.getAttribute(name);
  if (raw === null || raw === undefined) return null;
  const n = Number.parseInt(raw, 10);
  return Number.isFinite(n) && n >= 0 ? n : null;
}

/** Parse a single-element OOXML fragment. Returns null for the empty string
 *  the engine hands back when the sheet declares no such element, for a
 *  fragment that fails to parse, and for a host with no XML parser. */
function parseXmlElement(xml: string): Element | null {
  if (!xml || typeof DOMParser === 'undefined') return null;
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const root = doc.documentElement;
  if (!root || root.getElementsByTagName('parsererror').length > 0) return null;
  return root.nodeName === 'parsererror' ? null : root;
}

/** Read an OOXML boolean attribute. Both spellings a schema-valid file may
 *  use are accepted; an absent attribute falls back to the schema default. */
function xmlFlag(el: Element | null, name: string, fallback: boolean): boolean {
  const raw = el?.getAttribute(name);
  if (raw === null || raw === undefined) return fallback;
  return raw === '1' || raw === 'true';
}

/** Text of a header/footer section child. The XML parser has already decoded
 *  the entities, which is the spelling `setSheetHeaderFooter` expects back. */
function xmlSectionText(el: Element | null, name: string): string {
  const child = el?.getElementsByTagName(name)[0];
  return child?.textContent ?? '';
}

/** Resolve a sheet view's tab state. `visibility` is authoritative when the
 *  engine reports it; an engine that carries only the two-state `tabHidden`
 *  cannot distinguish `veryHidden`, so such a sheet reads as `Hidden`. */
function sheetVisibilityOf(view: { tabHidden: number; visibility?: number }): SheetVisibility {
  if (view.visibility === SheetVisibility.VeryHidden) return SheetVisibility.VeryHidden;
  if (view.visibility === SheetVisibility.Hidden || view.tabHidden !== 0) {
    return SheetVisibility.Hidden;
  }
  return SheetVisibility.Visible;
}

type WorkbookHandleCtor = { prototype: WorkbookHandle };
type WorkbookHandleInternals = {
  wb: Workbook;
  module: FormulonModule;
  capabilities: EngineCapabilities;
  functionMetadataProvider: FunctionMetadataProvider | null;
  assertAlive(): void;
};
type EngineCommentEntry = { row: number; col: number; author: string; text: string };
type CommentEnumerableWorkbook = Workbook & {
  getComments?: (sheet: number) => EngineCommentEntry[];
};
type TableAuthoringWorkbook = Workbook & {
  createTable?: (input: TableInput) => { status: { ok: boolean }; index: number };
  updateTable?: (
    index: number,
    input: Pick<TableInput, 'ref' | 'styleName' | 'headerRow' | 'totalsRow'>,
  ) => { ok: boolean };
  removeTable?: (index: number) => { ok: boolean };
};
type AutoFilterWorkbook = Workbook & {
  getSheetAutoFilterXml?: (sheet: number) => { status: { ok: boolean }; xml: string };
  setSheetAutoFilterXml?: (sheet: number, xml: string) => { ok: boolean };
};

declare module './workbook-handle.js' {
  interface WorkbookHandle extends WorkbookHandleFeatureMethods {}
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

function moduleOf(handle: unknown): FormulonModule {
  return internals(handle).module;
}

export abstract class WorkbookHandleFeatureMethods {
  declare readonly capabilities: EngineCapabilities;
  declare readonly sheetCount: number;
  abstract getValue(addr: Addr): CellValue;

  /** Persist a column-width override on `[first, last]` for `sheet`.
   *  No-op (returns false) when the engine doesn't expose `setColumnWidth`
   *  — i.e. the stub fallback. The UI store stays the source of truth in
   *  that case so paint still reflects the drag. */
  setColumnWidth(sheet: number, first: number, last: number, width: number): boolean {
    assertAlive(this);
    if (!this.capabilities.colRowSize) return false;
    const s = wb(this).setColumnWidth(sheet, first, last, width);
    return s.ok;
  }

  /** Persist a row-height override at `row` for `sheet`. See `setColumnWidth`
   *  for the no-op-on-stub rationale. */
  setRowHeight(sheet: number, row: number, height: number): boolean {
    assertAlive(this);
    if (!this.capabilities.colRowSize) return false;
    const s = wb(this).setRowHeight(sheet, row, height);
    return s.ok;
  }

  /** Snapshot of column overrides on `sheet`. Empty array under the stub.
   *  The returned objects own no engine memory — the underlying vector
   *  handle is released before this method returns. */
  getColumnLayouts(
    sheet: number,
  ): { first: number; last: number; width: number; hidden: boolean; outlineLevel: number }[] {
    assertAlive(this);
    if (!this.capabilities.colRowSize) return [];
    const r = wb(this).getSheetColumns(sheet);
    const out: {
      first: number;
      last: number;
      width: number;
      hidden: boolean;
      outlineLevel: number;
    }[] = [];
    if (!r.status.ok) return out;
    for (const e of r.columns) {
      out.push({
        first: e.first,
        last: e.last,
        width: e.width,
        hidden: e.hidden !== 0,
        outlineLevel: e.outlineLevel,
      });
    }
    return out;
  }

  /** Persist frozen-pane counts on `sheet`. No-op (returns false) under stub. */
  setSheetFreeze(sheet: number, freezeRows: number, freezeCols: number): boolean {
    assertAlive(this);
    if (!this.capabilities.freeze) return false;
    const s = wb(this).setSheetFreeze(sheet, freezeRows, freezeCols);
    return s.ok;
  }

  /** Persist sheet zoom percentage (10..400, engine clamps). No-op under stub. */
  setSheetZoom(sheet: number, zoomScale: number): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetZoom) return false;
    const s = wb(this).setSheetZoom(sheet, zoomScale);
    return s.ok;
  }

  /** Toggle the tab-hidden flag on `sheet`. Returns false on engine failure
   *  or when the engine doesn't expose `setSheetTabHidden`.
   *
   *  This is the two-state view: `true` on an already very-hidden sheet leaves
   *  it very-hidden, and `false` reveals it from either hidden state. Use
   *  `setSheetVisibility` to move between the two hidden states. */
  setSheetTabHidden(sheet: number, hidden: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetTabHidden) return false;
    const s = wb(this).setSheetTabHidden(sheet, hidden);
    return s.ok;
  }

  /** Set `sheet`'s tab to one of the three OOXML visibility states. Returns
   *  false when the engine only carries the two-state `setSheetTabHidden`,
   *  which cannot express `veryHidden`. */
  setSheetVisibility(sheet: number, visibility: SheetVisibility): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetVisibility) return false;
    const s = wb(this).setSheetVisibility(sheet, visibility as number);
    return s.ok;
  }

  /** Set the hidden flag on `[first, last]` columns. No-op under stub or
   *  when the engine doesn't expose `setColumnHidden`. */
  setColumnHidden(sheet: number, first: number, last: number, hidden: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.hiddenRowsCols) return false;
    const s = wb(this).setColumnHidden(sheet, first, last, hidden);
    return s.ok;
  }

  /** Set the hidden flag on `row`. */
  setRowHidden(sheet: number, row: number, hidden: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.hiddenRowsCols) return false;
    const s = wb(this).setRowHidden(sheet, row, hidden);
    return s.ok;
  }

  /** Set the outline level on `[first, last]` columns (0..7). */
  setColumnOutline(sheet: number, first: number, last: number, level: number): boolean {
    assertAlive(this);
    if (!this.capabilities.outlines) return false;
    const s = wb(this).setColumnOutline(sheet, first, last, level);
    return s.ok;
  }

  /** Set the outline level on `row` (0..7). */
  setRowOutline(sheet: number, row: number, level: number): boolean {
    assertAlive(this);
    if (!this.capabilities.outlines) return false;
    const s = wb(this).setRowOutline(sheet, row, level);
    return s.ok;
  }

  /** Snapshot of `sheet`'s view, including display flags. Returns null when
   *  the engine doesn't expose `getSheetView` (i.e. the stub or an older bundle). */
  getSheetView(sheet: number): {
    zoomScale: number;
    freezeRows: number;
    freezeCols: number;
    tabHidden: boolean;
    visibility: SheetVisibility;
    showGridLines: boolean;
    showRowColHeaders: boolean;
    showZeros: boolean;
    rightToLeft: boolean;
  } | null {
    assertAlive(this);
    if (!this.capabilities.sheetView) return null;
    const r = wb(this).getSheetView(sheet);
    if (!r.status.ok) return null;
    return {
      zoomScale: r.view.zoomScale,
      freezeRows: r.view.freezeRows,
      freezeCols: r.view.freezeCols,
      tabHidden: r.view.tabHidden !== 0,
      // Engines predating three-state visibility carry only `tabHidden`.
      visibility: sheetVisibilityOf(r.view),
      showGridLines: r.view.showGridLines !== 0,
      showRowColHeaders: r.view.showRowColHeaders !== 0,
      showZeros: r.view.showZeros !== 0,
      rightToLeft: r.view.rightToLeft !== 0,
    };
  }

  setSheetShowGridLines(sheet: number, show: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetViewFlags) return false;
    return wb(this).setSheetShowGridLines(sheet, show).ok;
  }

  setSheetShowRowColHeaders(sheet: number, show: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetViewFlags) return false;
    return wb(this).setSheetShowRowColHeaders(sheet, show).ok;
  }

  setSheetShowZeros(sheet: number, show: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetViewFlags) return false;
    return wb(this).setSheetShowZeros(sheet, show).ok;
  }

  setSheetRightToLeft(sheet: number, rightToLeft: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetViewFlags) return false;
    return wb(this).setSheetRightToLeft(sheet, rightToLeft).ok;
  }

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
    const el = parsePageSetupFragment(r.xml);
    return {
      blackAndWhite: xmlFlag(el, 'blackAndWhite', false),
      draft: xmlFlag(el, 'draft', false),
      cellComments: el?.getAttribute('cellComments') ?? '',
      errors: el?.getAttribute('errors') ?? '',
      pageOrder: el?.getAttribute('pageOrder') ?? '',
      firstPageNumber: xmlUint(el, 'firstPageNumber'),
      useFirstPageNumber: xmlFlag(el, 'useFirstPageNumber', false),
      horizontalDpi: xmlUint(el, 'horizontalDpi'),
      verticalDpi: xmlUint(el, 'verticalDpi'),
    };
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
    const el = parseXmlElement(r.xml);
    return {
      gridLines: xmlFlag(el, 'gridLines', false),
      headings: xmlFlag(el, 'headings', false),
      horizontalCentered: xmlFlag(el, 'horizontalCentered', false),
      verticalCentered: xmlFlag(el, 'verticalCentered', false),
    };
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
    const el = parseXmlElement(r.xml);
    return {
      oddHeader: xmlSectionText(el, 'oddHeader'),
      oddFooter: xmlSectionText(el, 'oddFooter'),
      differentOddEven: xmlFlag(el, 'differentOddEven', false),
      differentFirst: xmlFlag(el, 'differentFirst', false),
      scaleWithDoc: xmlFlag(el, 'scaleWithDoc', true),
      alignWithMargins: xmlFlag(el, 'alignWithMargins', true),
    };
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

  /** Returns the complete worksheet `<autoFilter>` fragment, or null when
   * the engine does not expose the optional raw-definition seam. */
  getSheetAutoFilterXml(sheet: number): string | null {
    assertAlive(this);
    if (!this.capabilities.autoFilter) return null;
    const r = (wb(this) as AutoFilterWorkbook).getSheetAutoFilterXml?.(sheet);
    return r?.status.ok ? r.xml : null;
  }

  /** Replaces the worksheet `<autoFilter>` fragment. Empty XML removes the
   * definition. Returns false for old bundles and failed writes. */
  setSheetAutoFilterXml(sheet: number, xml: string): boolean {
    assertAlive(this);
    if (!this.capabilities.autoFilter) return false;
    return (wb(this) as AutoFilterWorkbook).setSheetAutoFilterXml?.(sheet, xml)?.ok === true;
  }

  /** Insert `count` blank rows at `row` on `sheet`. The engine rewrites
   *  cross-workbook formula refs to follow the shift. Returns false on
   *  engines without `insertDeleteRowsCols`. NOT routed through the
   *  per-cell journal — callers wrap this in their own history entry. */
  engineInsertRows(sheet: number, row: number, count: number): boolean {
    assertAlive(this);
    if (!this.capabilities.insertDeleteRowsCols) return false;
    const s = wb(this).insertRows(sheet, row, count);
    return s.ok;
  }

  /** Delete `count` rows starting at `row` on `sheet`. Refs that fall
   *  inside the deleted interval collapse to `#REF!`. */
  engineDeleteRows(sheet: number, row: number, count: number): boolean {
    assertAlive(this);
    if (!this.capabilities.insertDeleteRowsCols) return false;
    const s = wb(this).deleteRows(sheet, row, count);
    return s.ok;
  }

  /** Insert `count` blank columns at `col` on `sheet`. */
  engineInsertCols(sheet: number, col: number, count: number): boolean {
    assertAlive(this);
    if (!this.capabilities.insertDeleteRowsCols) return false;
    const s = wb(this).insertCols(sheet, col, count);
    return s.ok;
  }

  /** Delete `count` columns starting at `col` on `sheet`. */
  engineDeleteCols(sheet: number, col: number, count: number): boolean {
    assertAlive(this);
    if (!this.capabilities.insertDeleteRowsCols) return false;
    const s = wb(this).deleteCols(sheet, col, count);
    return s.ok;
  }

  /** Read the XF (eXtended Format) table index assigned to `(sheet, row, col)`.
   *  Returns 0 (the workbook's default XF row) on missing cells. Returns null
   *  when the engine doesn't expose `getCellXfIndex` — i.e. the stub or older
   *  bundles. */
  getCellXfIndex(sheet: number, row: number, col: number): number | null {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return null;
    const r = wb(this).getCellXfIndex(sheet, row, col);
    if (!r.status.ok) return null;
    return r.xfIndex;
  }

  /** Pin the XF index of `(sheet, row, col)` to `xfIndex`. The index must
   *  point at an existing row in the workbook's XF table — there is no
   *  upstream API to insert new XF rows yet, so this is mainly useful for
   *  cloning formatting from one cell to another (Format-Painter parity at
   *  the engine layer) or for clearing back to xfIndex 0 (the default).
   *  Returns false on engine failure or when `capabilities.cellFormatting`
   *  is off. */
  setCellXfIndex(sheet: number, row: number, col: number, xfIndex: number): boolean {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return false;
    const s = wb(this).setCellXfIndex(sheet, row, col, xfIndex);
    return s.ok;
  }

  /** Read the cell's OOXML phonetic guide, if the current engine exposes it. */
  getCellPhonetic(sheet: number, row: number, col: number): string | null {
    assertAlive(this);
    if (!this.capabilities.phonetic) return null;
    const r = wb(this).getCellPhonetic(sheet, row, col);
    return r.status.ok && r.value ? r.value : null;
  }

  /** Set (or, with an empty string, clear) the cell's phonetic guide. */
  setCellPhonetic(sheet: number, row: number, col: number, phonetic: string): boolean {
    assertAlive(this);
    if (!this.capabilities.phonetic) return false;
    return wb(this).setCellPhonetic(sheet, row, col, phonetic).ok;
  }

  /** Resolve the XF record at `xfIndex` to its component table indices
   *  (font / fill / border / number-format) plus alignment + wrap flags.
   *  Note that the component indices are themselves opaque without
   *  resolver APIs (`getFont(idx)`, `getFill(idx)`, …) which upstream has
   *  not exposed yet — so this is currently most useful as a metadata
   *  signal (e.g. "do these two cells share the same XF row?"). Returns
   *  null on engine failure or capability off. */
  getCellXf(xfIndex: number): CellXf | null {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return null;
    const r = wb(this).getCellXf(xfIndex);
    if (!r.status.ok) return null;
    const { status: _status, ...record } = r;
    return record;
  }

  /** Resolve a font index to its plain-data record. Returns null on engine
   *  failure or when `capabilities.cellFormatting` is off. */
  getFontRecord(fontIndex: number): FontRecord | null {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return null;
    const r = wb(this).getFont(fontIndex);
    if (!r.status.ok) return null;
    const { status: _status, ...record } = r;
    return record;
  }

  /** Resolve a fill index to its plain-data record. */
  getFillRecord(fillIndex: number): FillRecord | null {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return null;
    const r = wb(this).getFill(fillIndex);
    if (!r.status.ok) return null;
    const { status: _status, ...record } = r;
    return record;
  }

  /** Resolve a border index to its plain-data record. */
  getBorderRecord(borderIndex: number): BorderRecord | null {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return null;
    const r = wb(this).getBorder(borderIndex);
    if (!r.status.ok) return null;
    const { status: _status, ...record } = r;
    return record;
  }

  /** Resolve a number-format id to its format-code string. */
  getNumFmtCode(numFmtId: number): string | null {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return null;
    const r = wb(this).getNumFmt(numFmtId);
    if (!r.status.ok) return null;
    return r.formatCode;
  }

  /** Add or dedup a font record. Returns the resolved font index, or -1 on
   *  engine failure or when capability is off. */
  addFontRecord(record: FontRecord): number {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return -1;
    const r = wb(this).addFont(completeFontRecord(record));
    return r.status.ok ? r.index : -1;
  }

  /** Add or dedup a fill record. */
  addFillRecord(record: FillRecord): number {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return -1;
    const r = wb(this).addFill(completeFillRecord(record));
    return r.status.ok ? r.index : -1;
  }

  /** Add or dedup a border record. */
  addBorderRecord(record: BorderRecord): number {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return -1;
    const r = wb(this).addBorder(completeBorderRecord(record));
    return r.status.ok ? r.index : -1;
  }

  /** Register a number-format code. Built-in matches return the built-in id;
   *  custom codes are appended starting at 164. Returns -1 on failure. */
  addNumFmtCode(formatCode: string): number {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return -1;
    const r = wb(this).addNumFmt(formatCode);
    return r.status.ok ? r.numFmtId : -1;
  }

  /** Add or dedup an XF (eXtended Format) record built from existing
   *  font/fill/border indices and a registered numFmtId. Returns the resolved
   *  xf index, or -1 on failure. */
  addXfRecord(record: CellXf): number {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return -1;
    const r = wb(this).addXf(record);
    return r.status.ok ? r.index : -1;
  }

  /** Creates an OOXML worksheet table and returns its index, or -1 when the
   * loaded engine predates table authoring support. */
  createTable(input: TableInput): number {
    assertAlive(this);
    if (!this.capabilities.tableMutate) return -1;
    const result = (wb(this) as TableAuthoringWorkbook).createTable?.(input);
    return result?.status.ok ? result.index : -1;
  }

  updateTable(
    index: number,
    input: Pick<TableInput, 'ref' | 'styleName' | 'headerRow' | 'totalsRow'>,
  ): boolean {
    assertAlive(this);
    return (
      this.capabilities.tableMutate === true &&
      (wb(this) as TableAuthoringWorkbook).updateTable?.(index, input).ok === true
    );
  }

  removeTable(index: number): boolean {
    assertAlive(this);
    return (
      this.capabilities.tableMutate === true &&
      (wb(this) as TableAuthoringWorkbook).removeTable?.(index).ok === true
    );
  }

  /** Append `range` as a merge on `sheet`. Returns false on engine failure or
   *  when `capabilities.merges` is off. The cell content inside the range is
   *  the caller's responsibility (spreadsheets keep top-left, blanks the rest). */
  engineAddMerge(sheet: number, range: Range): boolean {
    assertAlive(this);
    if (!this.capabilities.merges) return false;
    const s = wb(this).addMerge(sheet, {
      firstRow: range.r0,
      firstCol: range.c0,
      lastRow: range.r1,
      lastCol: range.c1,
    });
    return s.ok;
  }

  /** Remove every merge on `sheet` overlapping `range` (inclusive). No-op when
   *  nothing overlaps. Returns false on engine failure or capability off. */
  engineRemoveMerge(sheet: number, range: Range): boolean {
    assertAlive(this);
    if (!this.capabilities.merges) return false;
    const s = wb(this).removeMerge(sheet, {
      firstRow: range.r0,
      firstCol: range.c0,
      lastRow: range.r1,
      lastCol: range.c1,
    });
    return s.ok;
  }

  /** Drop every merge on `sheet`. Returns false on engine failure or capability
   *  off. */
  engineClearMerges(sheet: number): boolean {
    assertAlive(this);
    if (!this.capabilities.merges) return false;
    const s = wb(this).clearMerges(sheet);
    return s.ok;
  }

  /** Snapshot of every merge on `sheet` as inclusive `Range` records. Empty
   *  array under stub or when `capabilities.merges` is off. */
  getMerges(sheet: number): Range[] {
    assertAlive(this);
    if (!this.capabilities.merges) return [];
    const arr = wb(this).getMerges(sheet);
    return arr.map((m) => ({
      sheet,
      r0: m.firstRow,
      c0: m.firstCol,
      r1: m.lastRow,
      c1: m.lastCol,
    }));
  }

  /** Read the cell comment at `(sheet, row, col)`. Returns null when the
   *  cell has no comment or when the engine doesn't expose `getComment`. */
  getComment(sheet: number, row: number, col: number): { author: string; text: string } | null {
    assertAlive(this);
    if (!this.capabilities.comments) return null;
    const e = wb(this).getComment(sheet, row, col);
    return e ? { author: e.author, text: e.text } : null;
  }

  /** Snapshot every comment on `sheet` when the engine exposes a sheet-wide
   *  enumerator. Empty under stub or older engines. */
  getComments(sheet: number): { row: number; col: number; author: string; text: string }[] {
    assertAlive(this);
    if (!this.capabilities.commentsEnumerable) return [];
    const engineWb = wb(this) as CommentEnumerableWorkbook;
    if (typeof engineWb.getComments !== 'function') return [];
    return engineWb.getComments(sheet).map((e) => ({
      row: e.row,
      col: e.col,
      author: e.author,
      text: e.text,
    }));
  }

  /** Persist a cell comment. Empty `text` removes it. No-op (returns false)
   *  under the stub. */
  setCommentEntry(sheet: number, row: number, col: number, author: string, text: string): boolean {
    assertAlive(this);
    if (!this.capabilities.comments) return false;
    const s = wb(this).setComment(sheet, row, col, author, text);
    return s.ok;
  }

  /** Evaluate a self-contained formula string (references already substituted
   *  with literal values) in a fresh engine workbook and return its
   *  `{ status, value }` result. Callers that need live-cell context must
   *  substitute refs first (see the custom data-validation evaluator and the
   *  F9 / evaluate-formula surfaces). */
  evalFormula(formula: string): EvalResult {
    assertAlive(this);
    return moduleOf(this).evalFormula(formula);
  }

  evaluateFormulaText(addr: Addr, formula: string): EvalResult {
    assertAlive(this);
    if (!this.capabilities.formulaTextEvaluation) return moduleOf(this).evalFormula(formula);
    return wb(this).evaluateFormulaText(addr.sheet, addr.row, addr.col, formula);
  }

  /** Ad-hoc evaluation that returns the *whole* dynamic-array / spilled
   *  result of `formula` anchored at `addr`, without mutating the workbook.
   *  The `cells` grid is row-major (`cells[r][c]`); a scalar is reported as a
   *  1x1 array. Falls back to a 1x1 wrapper around `evaluateFormulaText` when
   *  the engine doesn't expose `evaluateFormulaArray`. */
  evaluateFormulaArray(addr: Addr, formula: string): EvalArrayResult {
    assertAlive(this);
    if (!this.capabilities.arrayFormulaEvaluation) {
      const single = this.evaluateFormulaText(addr, formula);
      return { status: single.status, rows: 1, cols: 1, cells: [[single.value]] };
    }
    return wb(this).evaluateFormulaArray(addr.sheet, addr.row, addr.col, formula);
  }

  evaluateConditionalFormula(
    addr: Addr,
    anchor: Pick<Addr, 'row' | 'col'>,
    formula: string,
  ): EvalResult {
    assertAlive(this);
    if (!this.capabilities.conditionalFormulaEvaluation) return moduleOf(this).evalFormula(formula);
    return wb(this).evaluateConditionalFormula(
      addr.sheet,
      addr.row,
      addr.col,
      anchor.row,
      anchor.col,
      formula,
    );
  }

  /** Evaluate every CF block on `sheet` against the inclusive viewport rect.
   *  Returns a sparse list — only cells with at least one match appear. The
   *  underlying embind vectors are released before this method returns, so
   *  the JS objects own no engine memory. Pass `NaN` for `todaySerial` to
   *  disable `TimePeriod` rules; defaults to `NaN`. Returns `[]` when the
   *  engine doesn't expose `evaluateCfRange`. */
  evaluateCfRange(
    sheet: number,
    firstRow: number,
    firstCol: number,
    lastRow: number,
    lastCol: number,
    todaySerial = Number.NaN,
  ): {
    row: number;
    col: number;
    matches: {
      kind: number;
      priority: number;
      dxfIdEngaged: boolean;
      dxfId: number;
      color: { r: number; g: number; b: number; a: number };
      barLengthPct: number;
      barAxisPositionPct: number;
      barIsNegative: boolean;
      barFill: { r: number; g: number; b: number; a: number };
      barBorderEngaged: boolean;
      barBorder: { r: number; g: number; b: number; a: number };
      barGradient: boolean;
      iconSetName: number;
      iconIndex: number;
    }[];
  }[] {
    assertAlive(this);
    if (!this.capabilities.conditionalFormat) return [];
    const r = wb(this).evaluateCfRange(sheet, firstRow, firstCol, lastRow, lastCol, todaySerial);
    if (!r.status.ok) return [];
    const out: ReturnType<WorkbookHandle['evaluateCfRange']> = [];
    for (const cell of r.cells) {
      const matches: ReturnType<WorkbookHandle['evaluateCfRange']>[number]['matches'] = [];
      for (const m of cell.matches) {
        matches.push({
          kind: m.kind as number,
          priority: m.priority,
          dxfIdEngaged: m.dxfIdEngaged !== 0,
          dxfId: m.dxfId,
          color: { r: m.color.r, g: m.color.g, b: m.color.b, a: m.color.a },
          barLengthPct: m.barLengthPct,
          barAxisPositionPct: m.barAxisPositionPct,
          barIsNegative: m.barIsNegative !== 0,
          barFill: { r: m.barFill.r, g: m.barFill.g, b: m.barFill.b, a: m.barFill.a },
          barBorderEngaged: m.barBorderEngaged !== 0,
          barBorder: {
            r: m.barBorder.r,
            g: m.barBorder.g,
            b: m.barBorder.b,
            a: m.barBorder.a,
          },
          barGradient: m.barGradient !== 0,
          iconSetName: m.iconSetName,
          iconIndex: m.iconIndex,
        });
      }
      out.push({ row: cell.row, col: cell.col, matches });
    }
    return out;
  }

  /** Returns the dynamic-array spill region engaged at `(sheet, row, col)`.
   *  The same struct is returned for the anchor cell and every phantom
   *  cell in the region. Returns `null` when the cell is not part of any
   *  spill or when the engine doesn't expose `spillInfo`. */
  spillInfo(
    sheet: number,
    row: number,
    col: number,
  ): { anchorRow: number; anchorCol: number; rows: number; cols: number } | null {
    assertAlive(this);
    if (!this.capabilities.spillInfo) return null;
    const r = wb(this).spillInfo(sheet, row, col);
    if (!r.engaged) return null;
    return {
      anchorRow: r.anchorRow,
      anchorCol: r.anchorCol,
      rows: r.rows,
      cols: r.cols,
    };
  }

  /** Returns every spill rect on `sheet` at engine precision. Returns
   *  `null` when the engine doesn't expose `spillInfo`; callers should
   *  fall back to the heuristic in `engine/spill.ts` in that case. */
  spillRanges(sheet: number): Range[] | null {
    assertAlive(this);
    if (!this.capabilities.spillInfo) return null;
    return computeEngineSpillRanges(this as unknown as WorkbookHandle, sheet);
  }

  /** Cells that `addr` directly reads (1-step precedents) by default;
   *  pass `depth > 1` for a BFS expansion (engine caps at 32 to avoid
   *  runaway in cyclic graphs). Includes cross-sheet refs — callers that
   *  only want same-sheet relations should filter on `sheet`. Returns
   *  `null` when the engine doesn't expose `precedents`; the regex-based
   *  same-sheet fallback in `engine/refs-graph.ts` covers stub mode. */
  precedents(addr: Addr, depth = 1): Addr[] | null {
    assertAlive(this);
    if (!this.capabilities.traceArrows) return null;
    const arr = wb(this).precedents(addr.sheet, addr.row, addr.col, depth);
    return arr.map((n) => ({ sheet: n.sheet, row: n.row, col: n.col }));
  }

  /** Cells whose formulas read from `addr` (1-step dependents by default).
   *  Same depth + cross-sheet semantics as `precedents`. Returns `null`
   *  when the engine doesn't expose `dependents`. */
  dependents(addr: Addr, depth = 1): Addr[] | null {
    assertAlive(this);
    if (!this.capabilities.traceArrows) return null;
    const arr = wb(this).dependents(addr.sheet, addr.row, addr.col, depth);
    return arr.map((n) => ({ sheet: n.sheet, row: n.row, col: n.col }));
  }

  /** Every registered function's canonical name in ascending sort order.
   *  Returns `null` when the engine doesn't expose `functionNames`; the
   *  static `FUNCTION_NAMES` list in `commands/refs.ts` is the fallback
   *  catalog under stub mode. */
  functionNames(): readonly string[] | null {
    assertAlive(this);
    if (!this.capabilities.functionMetadata) return null;
    return wb(this).functionNames();
  }

  /** Register a host-supplied function-metadata provider — a map of canonical
   *  UPPERCASE function name to localized signature/description/alias overrides
   *  — that `functionMetadata` merges over the engine's structural catalog.
   *  Pass `null` to clear. Display-only: it never affects parsing or
   *  evaluation. See `docs/function-metadata-schema.md`. */
  setFunctionMetadataProvider(provider: FunctionMetadataProvider | null): void {
    internals(this).functionMetadataProvider = provider;
  }

  /** Engine metadata for `name` (case-insensitive), with any host-registered
   *  provider entry merged over it (see `setFunctionMetadataProvider`).
   *  `locale`: 0 = en-US, 1 = ja-JP. The engine guarantees `minArity` /
   *  `maxArity` whenever the function is known; `signatureTemplate` and
   *  `description` come from the per-locale metadata table or a provider
   *  override, and are absent until either is populated. `localizedName`
   *  carries the provider's locale alias when present. Returns `null` when the
   *  engine doesn't expose `functionMetadata` or the function is unknown.
   *  `maxArity` may be `null` to denote an unbounded variadic or a lazy /
   *  special form whose upper arity is unknown. */
  functionMetadata(
    name: string,
    locale = 0,
  ): {
    name: string;
    minArity: number;
    maxArity: number | null;
    signatureTemplate?: string;
    description?: string;
    localizedName?: string;
  } | null {
    assertAlive(this);
    if (!this.capabilities.functionMetadata) return null;
    const m = wb(this).functionMetadata(name, locale);
    if (!m.ok) return null;
    const provider = internals(this).functionMetadataProvider;
    const canonical = m.name ?? name;
    const entry = provider?.[canonical.toUpperCase()];
    const merged = mergeFunctionMetadata(m, entry, localeTag(locale));
    return {
      name: canonical,
      minArity: merged.minArity ?? 0,
      maxArity: merged.maxArity ?? null,
      ...(merged.signatureTemplate ? { signatureTemplate: merged.signatureTemplate } : {}),
      ...(merged.description ? { description: merged.description } : {}),
      ...(merged.localizedName ? { localizedName: merged.localizedName } : {}),
    };
  }

  /** Canonical → localized function-name lookup. `locale`: 0 = en-US,
   *  1 = ja-JP. Returns the canonical name unchanged when no alias is
   *  registered for `locale` (currently the case for every locale except
   *  en-US). Returns `null` when the engine doesn't expose
   *  `localizeFunctionName`. */
  localizeFunctionName(canonicalName: string, locale = 0): string | null {
    assertAlive(this);
    if (!this.capabilities.functionLocale) return null;
    return wb(this).localizeFunctionName(canonicalName, locale);
  }

  /** Localized → canonical function-name lookup. Falls through to a
   *  case-insensitive match on the canonical name when no alias is
   *  registered. Returns the empty string when the engine reports no
   *  matching function. Returns `null` when the engine doesn't expose
   *  `canonicalizeFunctionName`. */
  canonicalizeFunctionName(localizedName: string, locale = 0): string | null {
    assertAlive(this);
    if (!this.capabilities.functionLocale) return null;
    return wb(this).canonicalizeFunctionName(localizedName, locale);
  }

  /** Workbook calc-mode metadata mirroring `<calcPr calcMode>`. The engine
   *  itself does NOT gate evaluation on this value — every `recalc()` call
   *  honours all dirty cells regardless of mode. This wrapper is what enforces
   *  it: `Manual` suppresses the automatic recalc that follows a cell write,
   *  leaving the cells dirty until Calculate Now. Returns `null` when the
   *  engine doesn't expose `calcMode`. Codes: 0 = Auto, 1 = Manual,
   *  2 = AutoNoTable (treated as Auto — data tables have no separate
   *  evaluation path here). */
  calcMode(): 0 | 1 | 2 | null {
    assertAlive(this);
    if (!this.capabilities.calcMode) return null;
    const mode = wb(this).calcMode();
    return (mode as 0 | 1 | 2) ?? null;
  }

  /** Sets the calc-mode metadata. Returns `false` (no-op) under stub or
   *  older engine package builds. Leaving Manual runs the recalc that the
   *  edits made while in Manual were denied, so switching back to Auto
   *  settles the sheet the way a spreadsheet does. */
  setCalcMode(mode: 0 | 1 | 2): boolean {
    assertAlive(this);
    if (!this.capabilities.calcMode) return false;
    const previous = this.calcMode();
    if (!wb(this).setCalcMode(mode).ok) return false;
    // These methods are mixed into WorkbookHandle at runtime, so the base
    // class's own members need the cast to be visible here.
    if (previous === 1 && mode !== 1) (this as unknown as WorkbookHandle).recalc();
    return true;
  }

  /** Formula-behaviour profile selected in the engine. Profiles model host
   *  differences across supported host profiles. Returns
   *  `null` when the engine package does not expose the profile API. */
  spreadsheetProfileId(): SpreadsheetProfileId | null {
    assertAlive(this);
    if (!this.capabilities.spreadsheetProfile) return null;
    const getProfile = (
      wb(this) as unknown as Record<string, ((this: Workbook) => string) | undefined>
    )[ENGINE_SPREADSHEET_PROFILE_GETTER];
    if (!getProfile) return null;
    return engineProfileToPublic(getProfile.call(wb(this)) as EngineSpreadsheetProfileId);
  }

  /** Sets the formula-behaviour profile. Returns `false` when unsupported. */
  setSpreadsheetProfileId(profileId: SpreadsheetProfileId): boolean {
    assertAlive(this);
    if (!this.capabilities.spreadsheetProfile) return false;
    const setProfile = (
      wb(this) as unknown as Record<
        string,
        ((this: Workbook, profile: EngineSpreadsheetProfileId) => { ok: boolean }) | undefined
      >
    )[ENGINE_SPREADSHEET_PROFILE_SETTER];
    if (!setProfile) return false;
    return setProfile.call(wb(this), publicProfileToEngine(profileId)).ok;
  }

  /** Number of `<cellStyle>` entries (named styles) registered on the
   *  workbook. Returns `0` under stub mode and older engine package builds. */
  cellStyleCount(): number {
    assertAlive(this);
    if (!this.capabilities.cellStyles) return 0;
    return wb(this).cellStyleCount();
  }

  /** Number of `<cellStyleXfs>` records — the named-style xf table that
   *  `CellStyleResult.xfId` indexes into. Returns `0` under stub mode and
   *  older engine package builds. */
  cellStyleXfCount(): number {
    assertAlive(this);
    if (!this.capabilities.cellStyles) return 0;
    return wb(this).cellStyleXfCount();
  }

  /** Snapshot of the named cell style at `index`. Returns `null` when the
   *  engine doesn't expose `getCellStyle` or the index is out of range. */
  getCellStyle(index: number): {
    name: string;
    xfId: number;
    builtinId: number;
    iLevel: number;
    hidden: boolean;
    customBuiltin: boolean;
  } | null {
    assertAlive(this);
    if (!this.capabilities.cellStyles) return null;
    const r = wb(this).getCellStyle(index);
    if (!r.status.ok) return null;
    return {
      name: r.name,
      xfId: r.xfId,
      builtinId: r.builtinId,
      iLevel: r.iLevel,
      hidden: r.hidden,
      customBuiltin: r.customBuiltin,
    };
  }

  /** Resolve the `<cellStyleXfs>` row used by a named style. This is
   *  separate from cell XF rows: `getCellStyle(index).xfId` points here. */
  getCellStyleXf(xfId: number): CellXf | null {
    assertAlive(this);
    if (!this.capabilities.cellStyles) return null;
    const r = wb(this).getCellStyleXf(xfId);
    if (!r.status.ok) return null;
    const { status: _status, ...record } = r;
    return record;
  }

  /** Append (deduplicating) a `<cellStyleXfs>` row and return its index, which
   *  is what `setNamedCellStyle` and a cell XF's `xfId` reference. Returns -1
   *  when the engine cannot author named styles. */
  addCellStyleXfRecord(record: CellXf): number {
    assertAlive(this);
    if (!this.capabilities.cellStyleMutate) return -1;
    const r = wb(this).addCellStyleXf(record);
    return r.status.ok ? r.index : -1;
  }

  /** Add or replace the `<cellStyle>` entry named `name`. `builtinId` is an
   *  OOXML ordinal, or null for a style with no built-in counterpart. */
  setNamedCellStyle(name: string, xfId: number, builtinId: number | null): boolean {
    assertAlive(this);
    if (!this.capabilities.cellStyleMutate) return false;
    return wb(this).setCellStyle(name, xfId, builtinId ?? CUSTOM_CELL_STYLE_BUILTIN_ID).ok;
  }

  /** Enumerate every named cell style on the workbook — combines
   *  `cellStyleCount` + `getCellStyle` into one snapshot suitable for
   *  populating a "Cell Styles" UI. Empty under stub mode. Hidden
   *  built-ins are filtered out — the gallery hides those by default. */
  getNamedCellStyles(): readonly NamedCellStyle[] {
    assertAlive(this);
    if (!this.capabilities.cellStyles) return [];
    return computeNamedCellStyles(this as unknown as WorkbookHandle);
  }

  /** Snapshot of every CF rule on `sheet`, in flattened priority order.
   *  Returns `[]` when the engine doesn't expose `getConditionalFormats`
   *  or when there are no rules. The entries borrow rule ids from the
   *  engine's storage; treat them as immutable view objects. */
  getConditionalFormats(sheet: number): readonly ConditionalFormatEntry[] {
    assertAlive(this);
    if (!this.capabilities.conditionalFormatMutate) return [];
    return wb(this).getConditionalFormats(sheet);
  }

  /** Removes the CF rule at `index` (flattened priority order). When the
   *  containing block becomes empty, the engine drops it too. Returns
   *  `false` (no-op) under stub mode and older engine package builds. */
  removeConditionalFormatAt(sheet: number, index: number): boolean {
    assertAlive(this);
    if (!this.capabilities.conditionalFormatMutate) return false;
    return wb(this).removeConditionalFormatAt(sheet, index).ok;
  }

  /** Drops every `<conditionalFormatting>` block on `sheet`. Returns
   *  `false` (no-op) under stub mode and older engine package builds. */
  clearConditionalFormats(sheet: number): boolean {
    assertAlive(this);
    if (!this.capabilities.conditionalFormatMutate) return false;
    return wb(this).clearConditionalFormats(sheet).ok;
  }

  /** Adds one conditional-format rule to `sheet` so it round-trips through
   *  .xlsx. `rule.type` mirrors `formulon::cf::RuleType` (0 expression,
   *  1 cellIs, 2 colorScale, 3 dataBar, 4 iconSet, 5 top10, 6 aboveAverage,
   *  7 containsText, 8 notContainsText, 9 beginsWith, 10 endsWith, …). The
   *  applied differential format is referenced by `dxfId`. Returns the
   *  flattened rule index, or -1 when the engine refuses the write. */
  addConditionalFormat(sheet: number, rule: ConditionalFormatInput): number {
    assertAlive(this);
    if (!this.capabilities.conditionalFormatMutate) return -1;
    const r = wb(this).addConditionalFormat(sheet, rule);
    return r.status.ok ? r.index : -1;
  }

  getDxf(index: number): DxfRecord | null {
    assertAlive(this);
    if (!this.capabilities.conditionalFormatDxf) return null;
    const r = wb(this).getDxf(index);
    if (!r.status.ok) return null;
    const { status: _status, ...record } = r;
    return record;
  }

  addDxf(record: DxfRecord): number {
    assertAlive(this);
    if (!this.capabilities.conditionalFormatDxf) return -1;
    const r = wb(this).addDxf(completeDxfRecord(record));
    return r.status.ok ? r.index : -1;
  }

  dxfCount(): number {
    assertAlive(this);
    if (!this.capabilities.conditionalFormatDxf) return 0;
    return wb(this).dxfCount();
  }

  /** Reads the round-trip `<sheetProtection>` flags. Returns `null` when
   *  the engine doesn't expose `getSheetProtection`. The booleans are
   *  reported as JS booleans (the engine wires them as 0/1 numbers); the
   *  `enabled` flag denotes whether the protection block is emitted on
   *  save. */
  getSheetProtection(sheet: number): {
    enabled: boolean;
    algorithmName: string;
    hashValue: string;
    saltValue: string;
    spinCount: number;
    legacyPassword: string;
    sheet: boolean;
    objects: boolean;
    scenarios: boolean;
    formatCells: boolean;
    formatColumns: boolean;
    formatRows: boolean;
    insertColumns: boolean;
    insertRows: boolean;
    insertHyperlinks: boolean;
    deleteColumns: boolean;
    deleteRows: boolean;
    selectLockedCells: boolean;
    selectUnlockedCells: boolean;
    sort: boolean;
    autoFilter: boolean;
    pivotTables: boolean;
  } | null {
    assertAlive(this);
    if (!this.capabilities.sheetProtectionRoundtrip) return null;
    const r = wb(this).getSheetProtection(sheet);
    if (!r.status.ok) return null;
    const p = r.protection;
    return {
      enabled: p.enabled !== 0,
      algorithmName: p.algorithmName,
      hashValue: p.hashValue,
      saltValue: p.saltValue,
      spinCount: p.spinCount,
      legacyPassword: p.legacyPassword,
      sheet: p.sheet !== 0,
      objects: p.objects !== 0,
      scenarios: p.scenarios !== 0,
      formatCells: p.formatCells !== 0,
      formatColumns: p.formatColumns !== 0,
      formatRows: p.formatRows !== 0,
      insertColumns: p.insertColumns !== 0,
      insertRows: p.insertRows !== 0,
      insertHyperlinks: p.insertHyperlinks !== 0,
      deleteColumns: p.deleteColumns !== 0,
      deleteRows: p.deleteRows !== 0,
      selectLockedCells: p.selectLockedCells !== 0,
      selectUnlockedCells: p.selectUnlockedCells !== 0,
      sort: p.sort !== 0,
      autoFilter: p.autoFilter !== 0,
      pivotTables: p.pivotTables !== 0,
    };
  }

  /** Replaces `<sheetProtection>` flags wholesale. Setting `enabled` to
   *  `false` clears the protection block on save. Returns `false` (no-op)
   *  under stub mode and older engine package builds. */
  setSheetProtection(
    sheet: number,
    protection: {
      enabled: boolean;
      legacyPassword?: string;
      algorithmName?: string;
      hashValue?: string;
      saltValue?: string;
      spinCount?: number;
      sheet?: boolean;
      objects?: boolean;
      scenarios?: boolean;
      formatCells?: boolean;
      formatColumns?: boolean;
      formatRows?: boolean;
      insertColumns?: boolean;
      insertRows?: boolean;
      insertHyperlinks?: boolean;
      deleteColumns?: boolean;
      deleteRows?: boolean;
      selectLockedCells?: boolean;
      selectUnlockedCells?: boolean;
      sort?: boolean;
      autoFilter?: boolean;
      pivotTables?: boolean;
    },
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetProtectionRoundtrip) return false;
    const b = (v: boolean | undefined): number => (v ? 1 : 0);
    const s = wb(this).setSheetProtection(sheet, {
      enabled: b(protection.enabled),
      algorithmName: protection.algorithmName ?? '',
      hashValue: protection.hashValue ?? '',
      saltValue: protection.saltValue ?? '',
      spinCount: protection.spinCount ?? 0,
      legacyPassword: protection.legacyPassword ?? '',
      sheet: b(protection.sheet ?? true),
      objects: b(protection.objects),
      scenarios: b(protection.scenarios),
      formatCells: b(protection.formatCells),
      formatColumns: b(protection.formatColumns),
      formatRows: b(protection.formatRows),
      insertColumns: b(protection.insertColumns),
      insertRows: b(protection.insertRows),
      insertHyperlinks: b(protection.insertHyperlinks),
      deleteColumns: b(protection.deleteColumns),
      deleteRows: b(protection.deleteRows),
      selectLockedCells: b(protection.selectLockedCells),
      selectUnlockedCells: b(protection.selectUnlockedCells),
      sort: b(protection.sort),
      autoFilter: b(protection.autoFilter),
      pivotTables: b(protection.pivotTables),
    });
    return s.ok;
  }
}

export function installWorkbookFeatureMethods(target: WorkbookHandleCtor): void {
  for (const key of Object.getOwnPropertyNames(WorkbookHandleFeatureMethods.prototype)) {
    if (key === 'constructor') continue;
    const descriptor = Object.getOwnPropertyDescriptor(WorkbookHandleFeatureMethods.prototype, key);
    if (!descriptor) continue;
    Object.defineProperty(target.prototype, key, descriptor);
  }
}
