import {
  ENGINE_SPREADSHEET_PROFILE_GETTER,
  ENGINE_SPREADSHEET_PROFILE_SETTER,
} from './capabilities.js';
import { computeNamedCellStyles, type NamedCellStyle } from './cell-styles-meta.js';
import { computeEngineSpillRanges } from './spill.js';
import {
  type EngineSpreadsheetProfileId,
  engineProfileToPublic,
  publicProfileToEngine,
} from './spreadsheet-profile.js';
import type {
  Addr,
  BorderRecord,
  CellValue,
  CellXf,
  EngineCapabilities,
  FillRecord,
  FontRecord,
  Range,
  SpreadsheetProfileId,
  Workbook,
} from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

type WorkbookHandleCtor = { prototype: WorkbookHandle };
type WorkbookHandleInternals = {
  wb: Workbook;
  capabilities: EngineCapabilities;
  assertAlive(): void;
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
    const v = r.columns;
    try {
      const n = v.size();
      for (let i = 0; i < n; i += 1) {
        const e = v.get(i);
        out.push({
          first: e.first,
          last: e.last,
          width: e.width,
          hidden: e.hidden !== 0,
          outlineLevel: e.outlineLevel,
        });
      }
    } finally {
      v.delete();
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
   *  or when the engine doesn't expose `setSheetTabHidden`. */
  setSheetTabHidden(sheet: number, hidden: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetTabHidden) return false;
    const s = wb(this).setSheetTabHidden(sheet, hidden);
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

  /** Snapshot of `sheet`'s view: zoom percentage, frozen-pane counts, and the
   *  tab-hidden flag. Returns null when the engine doesn't expose `getSheetView`
   *  (i.e. the stub or an older bundle). */
  getSheetView(
    sheet: number,
  ): { zoomScale: number; freezeRows: number; freezeCols: number; tabHidden: boolean } | null {
    assertAlive(this);
    if (!this.capabilities.sheetZoom) return null;
    const r = wb(this).getSheetView(sheet);
    if (!r.status.ok) return null;
    return {
      zoomScale: r.view.zoomScale,
      freezeRows: r.view.freezeRows,
      freezeCols: r.view.freezeCols,
      tabHidden: r.view.tabHidden !== 0,
    };
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

  /** Resolve the XF record at `xfIndex` to its component table indices
   *  (font / fill / border / number-format) plus alignment + wrap flags.
   *  Note that the component indices are themselves opaque without
   *  resolver APIs (`getFont(idx)`, `getFill(idx)`, …) which upstream has
   *  not exposed yet — so this is currently most useful as a metadata
   *  signal (e.g. "do these two cells share the same XF row?"). Returns
   *  null on engine failure or capability off. */
  getCellXf(xfIndex: number): {
    fontIndex: number;
    fillIndex: number;
    borderIndex: number;
    numFmtId: number;
    horizontalAlign: number;
    verticalAlign: number;
    wrapText: boolean;
  } | null {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return null;
    const r = wb(this).getCellXf(xfIndex);
    if (!r.status.ok) return null;
    return {
      fontIndex: r.fontIndex,
      fillIndex: r.fillIndex,
      borderIndex: r.borderIndex,
      numFmtId: r.numFmtId,
      horizontalAlign: r.horizontalAlign,
      verticalAlign: r.verticalAlign,
      wrapText: r.wrapText,
    };
  }

  /** Resolve a font index to its plain-data record. Returns null on engine
   *  failure or when `capabilities.cellFormatting` is off. */
  getFontRecord(fontIndex: number): FontRecord | null {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return null;
    const r = wb(this).getFont(fontIndex);
    if (!r.status.ok) return null;
    return {
      name: r.name,
      size: r.size,
      bold: r.bold,
      italic: r.italic,
      strike: r.strike,
      underline: r.underline,
      colorArgb: r.colorArgb,
    };
  }

  /** Resolve a fill index to its plain-data record. */
  getFillRecord(fillIndex: number): FillRecord | null {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return null;
    const r = wb(this).getFill(fillIndex);
    if (!r.status.ok) return null;
    return { pattern: r.pattern, fgArgb: r.fgArgb, bgArgb: r.bgArgb };
  }

  /** Resolve a border index to its plain-data record. */
  getBorderRecord(borderIndex: number): BorderRecord | null {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return null;
    const r = wb(this).getBorder(borderIndex);
    if (!r.status.ok) return null;
    return {
      left: { style: r.left.style, colorArgb: r.left.colorArgb },
      right: { style: r.right.style, colorArgb: r.right.colorArgb },
      top: { style: r.top.style, colorArgb: r.top.colorArgb },
      bottom: { style: r.bottom.style, colorArgb: r.bottom.colorArgb },
      diagonal: { style: r.diagonal.style, colorArgb: r.diagonal.colorArgb },
      diagonalUp: r.diagonalUp,
      diagonalDown: r.diagonalDown,
    };
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
    const r = wb(this).addFont(record);
    return r.status.ok ? r.index : -1;
  }

  /** Add or dedup a fill record. */
  addFillRecord(record: FillRecord): number {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return -1;
    const r = wb(this).addFill(record);
    return r.status.ok ? r.index : -1;
  }

  /** Add or dedup a border record. */
  addBorderRecord(record: BorderRecord): number {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return -1;
    const r = wb(this).addBorder(record);
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

  /** Persist a cell comment. Empty `text` removes it. No-op (returns false)
   *  under the stub. */
  setCommentEntry(sheet: number, row: number, col: number, author: string, text: string): boolean {
    assertAlive(this);
    if (!this.capabilities.comments) return false;
    const s = wb(this).setComment(sheet, row, col, author, text);
    return s.ok;
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
    const cells = r.cells;
    try {
      const n = cells.size();
      for (let i = 0; i < n; i += 1) {
        const cell = cells.get(i);
        const matches: ReturnType<WorkbookHandle['evaluateCfRange']>[number]['matches'] = [];
        const mv = cell.matches;
        try {
          const mn = mv.size();
          for (let j = 0; j < mn; j += 1) {
            const m = mv.get(j);
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
        } finally {
          mv.delete();
        }
        out.push({ row: cell.row, col: cell.col, matches });
      }
    } finally {
      cells.delete();
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

  /** Engine metadata for `name` (case-insensitive). `locale`: 0 = en-US,
   *  1 = ja-JP. The engine guarantees `minArity` / `maxArity` whenever
   *  the function is known; `signatureTemplate` and `description` come
   *  from the per-locale metadata table and are absent until that table
   *  is populated upstream. Returns `null` when the engine doesn't
   *  expose `functionMetadata` or the function is unknown. `maxArity`
   *  may be `0xFFFFFFFF` to denote unbounded variadic. */
  functionMetadata(
    name: string,
    locale = 0,
  ): {
    name: string;
    minArity: number;
    maxArity: number;
    signatureTemplate?: string;
    description?: string;
  } | null {
    assertAlive(this);
    if (!this.capabilities.functionMetadata) return null;
    const m = wb(this).functionMetadata(name, locale);
    if (!m.ok) return null;
    return {
      name: m.name ?? name,
      minArity: m.minArity ?? 0,
      maxArity: m.maxArity ?? 0,
      ...(m.signatureTemplate ? { signatureTemplate: m.signatureTemplate } : {}),
      ...(m.description ? { description: m.description } : {}),
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
   *  honours all dirty cells regardless of mode. The flag is preserved as
   *  round-trip metadata and surfaced here so the UI can mirror the spreadsheet's
   *  user-visible state. Returns `null` when the engine doesn't expose
   *  `calcMode`. Codes: 0 = Auto, 1 = Manual, 2 = AutoNoTable. */
  calcMode(): 0 | 1 | 2 | null {
    assertAlive(this);
    if (!this.capabilities.calcMode) return null;
    const mode = wb(this).calcMode();
    return (mode as 0 | 1 | 2) ?? null;
  }

  /** Sets the calc-mode metadata. Returns `false` (no-op) under stub or
   *  older engine package builds. */
  setCalcMode(mode: 0 | 1 | 2): boolean {
    assertAlive(this);
    if (!this.capabilities.calcMode) return false;
    return wb(this).setCalcMode(mode).ok;
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
  getConditionalFormats(sheet: number): ReadonlyArray<{
    id: string;
    type: number;
    priority: number;
    stopIfTrue: boolean;
    sqref: ReadonlyArray<{ firstRow: number; firstCol: number; lastRow: number; lastCol: number }>;
    dxfId?: number;
    formula1?: string;
    formula2?: string;
    op?: number;
    rank?: number;
    percent?: boolean;
    bottom?: boolean;
    aboveAverage?: boolean;
    equalAverage?: boolean;
    stdDev?: number;
    text?: string;
    timePeriod?: number;
  }> {
    assertAlive(this);
    if (!this.capabilities.conditionalFormatMutate) return [];
    const arr = wb(this).getConditionalFormats(sheet);
    return arr.map((e) => ({
      id: e.id,
      type: e.type,
      priority: e.priority,
      stopIfTrue: e.stopIfTrue,
      sqref: e.sqref.map((r) => ({
        firstRow: r.firstRow,
        firstCol: r.firstCol,
        lastRow: r.lastRow,
        lastCol: r.lastCol,
      })),
      ...(e.dxfId !== undefined ? { dxfId: e.dxfId } : {}),
      ...(e.formula1 !== undefined ? { formula1: e.formula1 } : {}),
      ...(e.formula2 !== undefined ? { formula2: e.formula2 } : {}),
      ...(e.op !== undefined ? { op: e.op } : {}),
      ...(e.rank !== undefined ? { rank: e.rank } : {}),
      ...(e.percent !== undefined ? { percent: e.percent } : {}),
      ...(e.bottom !== undefined ? { bottom: e.bottom } : {}),
      ...(e.aboveAverage !== undefined ? { aboveAverage: e.aboveAverage } : {}),
      ...(e.equalAverage !== undefined ? { equalAverage: e.equalAverage } : {}),
      ...(e.stdDev !== undefined ? { stdDev: e.stdDev } : {}),
      ...(e.text !== undefined ? { text: e.text } : {}),
      ...(e.timePeriod !== undefined ? { timePeriod: e.timePeriod } : {}),
    }));
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
