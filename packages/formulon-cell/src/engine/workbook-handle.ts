import type { History } from '../commands/history.js';
import { detectCapabilities } from './capabilities.js';
import type { LoadOptions } from './loader.js';
import { isUsingStub, loadFormulon } from './loader.js';
import { parseRangeRef as parseTableRef } from './range-resolver.js';
import { computeEngineSpillRanges } from './spill.js';
import type {
  Addr,
  BorderRecord,
  CellValue,
  CellXf,
  DataValidationInput,
  EngineCapabilities,
  FillRecord,
  FontRecord,
  FormulonModule,
  Range,
  Workbook,
} from './types.js';
import { formatCell, fromEngineValue } from './value.js';

export type ChangeListener = (e: ChangeEvent) => void;

export type ChangeEvent =
  | { kind: 'value'; addr: Addr; next: CellValue }
  | { kind: 'recalc'; dirty: ReadonlySet<string> }
  | { kind: 'sheet-add'; index: number; name: string }
  | { kind: 'sheet-rename'; index: number; name: string }
  | { kind: 'sheet-remove'; index: number }
  | { kind: 'sheet-move'; from: number; to: number };

const addrKey = (a: Addr): string => `${a.sheet}:${a.row}:${a.col}`;

const kindLabel = (kind: number): 'unknown' | 'externalBook' | 'ole' | 'dde' => {
  switch (kind) {
    case 1:
      return 'externalBook';
    case 2:
      return 'ole';
    case 3:
      return 'dde';
    default:
      return 'unknown';
  }
};

/** Snapshot of one cell's full state — enough to restore it on undo. */
interface CellSnapshot {
  addr: Addr;
  value: CellValue;
  formula: string | null;
}

const UNDO_LIMIT = 100;

/**
 * Boundary between the WASM Workbook and the rest of the UI. Nothing else
 * in the codebase touches the raw engine. Keeps the dispose contract honest.
 */
export class WorkbookHandle {
  readonly capabilities: EngineCapabilities;

  private readonly module: FormulonModule;

  private readonly wb: Workbook;

  private readonly listeners = new Set<ChangeListener>();

  private disposed = false;

  /** Per-cell inverse history. Each setX pushes one entry; undo replays
   *  it back. Used as a fallback when no `History` is attached. When a
   *  shared History is attached (mount.ts does this), entries route there
   *  instead so format/layout/value undo stay in lockstep. */
  private undoStack: CellSnapshot[] = [];

  private redoStack: CellSnapshot[] = [];

  /** Suppresses journal capture while we're applying an undo/redo. */
  private replaying = false;

  /** Optional shared history. When set, every setX captures a before/after
   *  snapshot pair and pushes a closure entry instead of using the local
   *  stack. */
  private history: History | null = null;

  /** Inclusive rect of cells currently visible to the user, supplied by the
   *  renderer via `setViewportHint`. Drives partial recalc on `setFormula`.
   *  Cleared on sheet switch / setWorkbook so we never apply a stale rect. */
  private viewportHint: {
    sheet: number;
    firstRow: number;
    firstCol: number;
    lastRow: number;
    lastCol: number;
  } | null = null;

  private constructor(module: FormulonModule, wb: Workbook) {
    this.module = module;
    this.wb = wb;
    this.capabilities = detectCapabilities(wb);
  }

  static async createDefault(opts: LoadOptions = {}): Promise<WorkbookHandle> {
    const module = await loadFormulon(opts);
    const wb = module.Workbook.createDefault();
    return new WorkbookHandle(module, wb);
  }

  static async loadBytes(bytes: Uint8Array, opts: LoadOptions = {}): Promise<WorkbookHandle> {
    const module = await loadFormulon(opts);
    const wb = module.Workbook.loadBytes(bytes);
    if (!wb.isValid()) {
      const msg = module.lastErrorMessage();
      wb.delete();
      throw new Error(`formulon loadBytes failed: ${msg}`);
    }
    return new WorkbookHandle(module, wb);
  }

  /** True when the JS fallback stub is providing the engine surface. */
  get isStub(): boolean {
    return isUsingStub();
  }

  get version(): string {
    return this.module.versionString();
  }

  get sheetCount(): number {
    this.assertAlive();
    return this.wb.sheetCount();
  }

  sheetName(idx: number): string {
    this.assertAlive();
    const r = this.wb.sheetName(idx);
    return r.status.ok ? r.value : `Sheet${idx + 1}`;
  }

  /** Append a new empty sheet. Returns the index of the newly added sheet,
   *  or -1 on failure. Emits a `sheet-add` event so the UI can append a tab. */
  addSheet(name?: string): number {
    this.assertAlive();
    const proposed = name ?? this.uniqueSheetName();
    const s = this.wb.addSheet(proposed);
    if (!s.ok) return -1;
    const idx = this.wb.sheetCount() - 1;
    this.emit({ kind: 'sheet-add', index: idx, name: proposed });
    return idx;
  }

  /** Rename the sheet at `idx`. Returns false on failure (e.g. duplicate name)
   *  or when the engine doesn't expose `renameSheet`. Emits `sheet-rename`. */
  renameSheet(idx: number, name: string): boolean {
    this.assertAlive();
    if (!this.capabilities.sheetMutate) return false;
    const s = this.wb.renameSheet(idx, name);
    if (!s.ok) return false;
    this.emit({ kind: 'sheet-rename', index: idx, name });
    return true;
  }

  /** Remove the sheet at `idx`. Returns false on failure (e.g. last sheet) or
   *  when the engine doesn't expose `removeSheet`. Emits `sheet-remove`. */
  removeSheet(idx: number): boolean {
    this.assertAlive();
    if (!this.capabilities.sheetMutate) return false;
    const s = this.wb.removeSheet(idx);
    if (!s.ok) return false;
    this.emit({ kind: 'sheet-remove', index: idx });
    return true;
  }

  /** Move the sheet at `from` to position `to` (post-removal index). Returns
   *  false on failure or when the engine doesn't expose `moveSheet`. Emits
   *  `sheet-move`. */
  moveSheet(from: number, to: number): boolean {
    this.assertAlive();
    if (!this.capabilities.sheetMutate) return false;
    const s = this.wb.moveSheet(from, to);
    if (!s.ok) return false;
    this.emit({ kind: 'sheet-move', from, to });
    return true;
  }

  private uniqueSheetName(): string {
    const existing = new Set<string>();
    const n = this.wb.sheetCount();
    for (let i = 0; i < n; i += 1) {
      const r = this.wb.sheetName(i);
      if (r.status.ok) existing.add(r.value);
    }
    let i = n + 1;
    while (existing.has(`Sheet${i}`)) i += 1;
    return `Sheet${i}`;
  }

  getValue(a: Addr): CellValue {
    this.assertAlive();
    const r = this.wb.getValue(a.sheet, a.row, a.col);
    if (!r.status.ok) return { kind: 'blank' };
    return fromEngineValue(r.value);
  }

  /** Wire a shared history into cell writes. Pass `null` to detach. The local
   *  fallback stack is cleared whenever attachment changes — entries from one
   *  source must not interleave with the other. */
  attachHistory(h: History | null): void {
    this.history = h;
    this.undoStack.length = 0;
    this.redoStack.length = 0;
  }

  setNumber(a: Addr, value: number): void {
    this.assertAlive();
    this.withJournal(a, () => {
      const s = this.wb.setNumber(a.sheet, a.row, a.col, value);
      if (!s.ok) throw new Error(`setNumber: ${s.message}`);
      this.emit({ kind: 'value', addr: a, next: { kind: 'number', value } });
    });
  }

  setText(a: Addr, value: string): void {
    this.assertAlive();
    this.withJournal(a, () => {
      const s = this.wb.setText(a.sheet, a.row, a.col, value);
      if (!s.ok) throw new Error(`setText: ${s.message}`);
      this.emit({ kind: 'value', addr: a, next: { kind: 'text', value } });
    });
  }

  setBool(a: Addr, value: boolean): void {
    this.assertAlive();
    this.withJournal(a, () => {
      const s = this.wb.setBool(a.sheet, a.row, a.col, value);
      if (!s.ok) throw new Error(`setBool: ${s.message}`);
      this.emit({ kind: 'value', addr: a, next: { kind: 'bool', value } });
    });
  }

  setBlank(a: Addr): void {
    this.assertAlive();
    this.withJournal(a, () => {
      const s = this.wb.setBlank(a.sheet, a.row, a.col);
      if (!s.ok) throw new Error(`setBlank: ${s.message}`);
      this.emit({ kind: 'value', addr: a, next: { kind: 'blank' } });
    });
  }

  setFormula(a: Addr, formula: string): void {
    this.assertAlive();
    this.withJournal(a, () => {
      const s = this.wb.setFormula(a.sheet, a.row, a.col, formula);
      if (!s.ok) throw new Error(`setFormula: ${s.message}`);
      this.scheduleRecalc(a);
      this.emit({ kind: 'value', addr: a, next: this.getValue(a) });
    });
  }

  /** Renderer hint — the inclusive rect of cells currently visible. Lets
   *  `partialRecalc` collapse to just the formulas the user can actually see.
   *  Internal: not exposed in the public API surface. */
  setViewportHint(
    sheet: number,
    firstRow: number,
    firstCol: number,
    lastRow: number,
    lastCol: number,
  ): void {
    this.viewportHint = { sheet, firstRow, firstCol, lastRow, lastCol };
  }

  /** Drop the viewport hint — the next setFormula falls back to full recalc.
   *  Called from sheet-switch and setWorkbook so we never apply a stale
   *  rect from a previous sheet. */
  clearViewportHint(): void {
    this.viewportHint = null;
  }

  /** Pick the most economical recalc for a single-cell edit at `a`. When a
   *  viewport hint exists for the same sheet AND the engine supports
   *  partialRecalc, we recompute only formulas whose dependency closure
   *  intersects the viewport rect plus the touched cell. Otherwise we fall
   *  back to a full recalc (still cheap on small workbooks). */
  private scheduleRecalc(a: Addr): void {
    if (!this.capabilities.partialRecalc || !this.viewportHint) {
      this.recalc();
      return;
    }
    const hint = this.viewportHint;
    if (hint.sheet !== a.sheet) {
      this.recalc();
      return;
    }
    const r0 = Math.min(hint.firstRow, a.row);
    const c0 = Math.min(hint.firstCol, a.col);
    const r1 = Math.max(hint.lastRow, a.row);
    const c1 = Math.max(hint.lastCol, a.col);
    const result = this.wb.partialRecalc({
      sheet: hint.sheet,
      firstRow: r0,
      firstCol: c0,
      lastRow: r1,
      lastCol: c1,
    });
    if (!result.status.ok) {
      // Partial failure shouldn't break the edit — fall back to full.
      this.recalc();
    }
  }

  canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  /** Legacy local undo. Returns false when a shared History is attached —
   *  callers should use `History.undo()` directly in that case. */
  undo(): boolean {
    this.assertAlive();
    if (this.history) return false;
    const snap = this.undoStack.pop();
    if (!snap) return false;
    const current = this.captureSnapshot(snap.addr);
    this.replay(snap);
    this.redoStack.push(current);
    return true;
  }

  redo(): boolean {
    this.assertAlive();
    if (this.history) return false;
    const snap = this.redoStack.pop();
    if (!snap) return false;
    const current = this.captureSnapshot(snap.addr);
    this.replay(snap);
    this.undoStack.push(current);
    return true;
  }

  recalc(): void {
    this.assertAlive();
    const s = this.wb.recalc();
    if (!s.ok) throw new Error(`recalc: ${s.message}`);
  }

  /** Recompute only formulas whose dependency closure intersects the viewport
   *  rectangle. Returns the number of cells the engine actually evaluated, or
   *  `null` when the engine doesn't expose `partialRecalc`. Falls back to a
   *  full `recalc()` on engines without the capability so callers can use
   *  this as a drop-in optimization. */
  partialRecalc(
    sheet: number,
    firstRow: number,
    firstCol: number,
    lastRow: number,
    lastCol: number,
  ): number | null {
    this.assertAlive();
    if (!this.capabilities.partialRecalc) {
      this.recalc();
      return null;
    }
    const r = this.wb.partialRecalc({ sheet, firstRow, firstCol, lastRow, lastCol });
    if (!r.status.ok) throw new Error(`partialRecalc: ${r.status.message}`);
    return r.recomputed;
  }

  /** Toggle the iterative-formula solver. `maxIterations` and `maxChange`
   *  cap the Gauss-Seidel loop; matches Excel's File → Options → Formulas
   *  knobs. No-op (returns false) on engines without the iterative surface. */
  setIterative(enabled: boolean, maxIterations: number, maxChange: number): boolean {
    this.assertAlive();
    if (!this.capabilities.iterativeProgress) return false;
    const s = this.wb.setIterative(enabled, maxIterations, maxChange);
    return s.ok;
  }

  /** Install (or clear) a progress callback invoked after each iterative-solve
   *  sweep. Returning `false` from the callback aborts the solve. Pass `null`
   *  to detach. No-op on engines without `setIterativeProgress`. */
  setIterativeProgress(
    callback:
      | ((iteration: number, maxResidual: number, maxIterations: number) => boolean | void)
      | null,
  ): boolean {
    this.assertAlive();
    if (!this.capabilities.iterativeProgress) return false;
    const s = this.wb.setIterativeProgress(callback);
    return s.ok;
  }

  /** Iterate over every populated cell on a sheet. Used for initial paint. */
  *cells(sheet: number): Generator<{ addr: Addr; value: CellValue; formula: string | null }> {
    this.assertAlive();
    const n = this.wb.cellCount(sheet);
    for (let i = 0; i < n; i += 1) {
      const e = this.wb.cellAt(sheet, i);
      if (!e.status.ok) continue;
      yield {
        addr: { sheet, row: e.row, col: e.col },
        value: fromEngineValue(e.value),
        formula: e.formula,
      };
    }
  }

  cellFormula(a: Addr): string | null {
    this.assertAlive();
    const n = this.wb.cellCount(a.sheet);
    for (let i = 0; i < n; i += 1) {
      const e = this.wb.cellAt(a.sheet, i);
      if (e.status.ok && e.row === a.row && e.col === a.col) return e.formula;
    }
    return null;
  }

  /** Iterate over defined names (workbook-scope). Useful for the name box
   *  drop-down. Each entry's `formula` text is what `definedNameAt` returns
   *  from the engine — typically a reference like `Sheet1!$A$1:$B$2`. */
  *definedNames(): Generator<{ name: string; formula: string }> {
    this.assertAlive();
    const n = this.wb.definedNameCount();
    for (let i = 0; i < n; i += 1) {
      const e = this.wb.definedNameAt(i);
      if (!e.status.ok) continue;
      yield { name: e.name, formula: e.formula };
    }
  }

  /** Add or replace a workbook-scoped defined name. Pass an empty `formula`
   *  to remove the name (engine convention). Returns false on engine failure
   *  or when the engine doesn't expose `setDefinedName` (stub or older bundle). */
  setDefinedNameEntry(name: string, formula: string): boolean {
    this.assertAlive();
    if (!this.capabilities.definedNameMutate) return false;
    const s = this.wb.setDefinedName(name, formula);
    return s.ok;
  }

  save(): Uint8Array {
    this.assertAlive();
    const r = this.wb.save();
    if (!r.status.ok || !r.bytes) throw new Error(`save: ${r.status.message}`);
    return r.bytes;
  }

  /** Persist a column-width override on `[first, last]` for `sheet`.
   *  No-op (returns false) when the engine doesn't expose `setColumnWidth`
   *  — i.e. the stub fallback. The UI store stays the source of truth in
   *  that case so paint still reflects the drag. */
  setColumnWidth(sheet: number, first: number, last: number, width: number): boolean {
    this.assertAlive();
    if (!this.capabilities.colRowSize) return false;
    const s = this.wb.setColumnWidth(sheet, first, last, width);
    return s.ok;
  }

  /** Persist a row-height override at `row` for `sheet`. See `setColumnWidth`
   *  for the no-op-on-stub rationale. */
  setRowHeight(sheet: number, row: number, height: number): boolean {
    this.assertAlive();
    if (!this.capabilities.colRowSize) return false;
    const s = this.wb.setRowHeight(sheet, row, height);
    return s.ok;
  }

  /** Snapshot of column overrides on `sheet`. Empty array under the stub.
   *  The returned objects own no engine memory — the underlying vector
   *  handle is released before this method returns. */
  getColumnLayouts(
    sheet: number,
  ): { first: number; last: number; width: number; hidden: boolean; outlineLevel: number }[] {
    this.assertAlive();
    if (!this.capabilities.colRowSize) return [];
    const r = this.wb.getSheetColumns(sheet);
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
    this.assertAlive();
    if (!this.capabilities.freeze) return false;
    const s = this.wb.setSheetFreeze(sheet, freezeRows, freezeCols);
    return s.ok;
  }

  /** Persist sheet zoom percentage (10..400, engine clamps). No-op under stub. */
  setSheetZoom(sheet: number, zoomScale: number): boolean {
    this.assertAlive();
    if (!this.capabilities.sheetZoom) return false;
    const s = this.wb.setSheetZoom(sheet, zoomScale);
    return s.ok;
  }

  /** Toggle the tab-hidden flag on `sheet`. Returns false on engine failure
   *  or when the engine doesn't expose `setSheetTabHidden`. */
  setSheetTabHidden(sheet: number, hidden: boolean): boolean {
    this.assertAlive();
    if (!this.capabilities.sheetTabHidden) return false;
    const s = this.wb.setSheetTabHidden(sheet, hidden);
    return s.ok;
  }

  /** Set the hidden flag on `[first, last]` columns. No-op under stub or
   *  when the engine doesn't expose `setColumnHidden`. */
  setColumnHidden(sheet: number, first: number, last: number, hidden: boolean): boolean {
    this.assertAlive();
    if (!this.capabilities.hiddenRowsCols) return false;
    const s = this.wb.setColumnHidden(sheet, first, last, hidden);
    return s.ok;
  }

  /** Set the hidden flag on `row`. */
  setRowHidden(sheet: number, row: number, hidden: boolean): boolean {
    this.assertAlive();
    if (!this.capabilities.hiddenRowsCols) return false;
    const s = this.wb.setRowHidden(sheet, row, hidden);
    return s.ok;
  }

  /** Set the outline level on `[first, last]` columns (0..7). */
  setColumnOutline(sheet: number, first: number, last: number, level: number): boolean {
    this.assertAlive();
    if (!this.capabilities.outlines) return false;
    const s = this.wb.setColumnOutline(sheet, first, last, level);
    return s.ok;
  }

  /** Set the outline level on `row` (0..7). */
  setRowOutline(sheet: number, row: number, level: number): boolean {
    this.assertAlive();
    if (!this.capabilities.outlines) return false;
    const s = this.wb.setRowOutline(sheet, row, level);
    return s.ok;
  }

  /** Snapshot of `sheet`'s view: zoom percentage, frozen-pane counts, and the
   *  tab-hidden flag. Returns null when the engine doesn't expose `getSheetView`
   *  (i.e. the stub or an older bundle). */
  getSheetView(
    sheet: number,
  ): { zoomScale: number; freezeRows: number; freezeCols: number; tabHidden: boolean } | null {
    this.assertAlive();
    if (!this.capabilities.sheetZoom) return null;
    const r = this.wb.getSheetView(sheet);
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
    this.assertAlive();
    if (!this.capabilities.insertDeleteRowsCols) return false;
    const s = this.wb.insertRows(sheet, row, count);
    return s.ok;
  }

  /** Delete `count` rows starting at `row` on `sheet`. Refs that fall
   *  inside the deleted interval collapse to `#REF!`. */
  engineDeleteRows(sheet: number, row: number, count: number): boolean {
    this.assertAlive();
    if (!this.capabilities.insertDeleteRowsCols) return false;
    const s = this.wb.deleteRows(sheet, row, count);
    return s.ok;
  }

  /** Insert `count` blank columns at `col` on `sheet`. */
  engineInsertCols(sheet: number, col: number, count: number): boolean {
    this.assertAlive();
    if (!this.capabilities.insertDeleteRowsCols) return false;
    const s = this.wb.insertCols(sheet, col, count);
    return s.ok;
  }

  /** Delete `count` columns starting at `col` on `sheet`. */
  engineDeleteCols(sheet: number, col: number, count: number): boolean {
    this.assertAlive();
    if (!this.capabilities.insertDeleteRowsCols) return false;
    const s = this.wb.deleteCols(sheet, col, count);
    return s.ok;
  }

  /** Read the XF (eXtended Format) table index assigned to `(sheet, row, col)`.
   *  Returns 0 (the workbook's default XF row) on missing cells. Returns null
   *  when the engine doesn't expose `getCellXfIndex` — i.e. the stub or older
   *  bundles. */
  getCellXfIndex(sheet: number, row: number, col: number): number | null {
    this.assertAlive();
    if (!this.capabilities.cellFormatting) return null;
    const r = this.wb.getCellXfIndex(sheet, row, col);
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
    this.assertAlive();
    if (!this.capabilities.cellFormatting) return false;
    const s = this.wb.setCellXfIndex(sheet, row, col, xfIndex);
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
    this.assertAlive();
    if (!this.capabilities.cellFormatting) return null;
    const r = this.wb.getCellXf(xfIndex);
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
    this.assertAlive();
    if (!this.capabilities.cellFormatting) return null;
    const r = this.wb.getFont(fontIndex);
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
    this.assertAlive();
    if (!this.capabilities.cellFormatting) return null;
    const r = this.wb.getFill(fillIndex);
    if (!r.status.ok) return null;
    return { pattern: r.pattern, fgArgb: r.fgArgb, bgArgb: r.bgArgb };
  }

  /** Resolve a border index to its plain-data record. */
  getBorderRecord(borderIndex: number): BorderRecord | null {
    this.assertAlive();
    if (!this.capabilities.cellFormatting) return null;
    const r = this.wb.getBorder(borderIndex);
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
    this.assertAlive();
    if (!this.capabilities.cellFormatting) return null;
    const r = this.wb.getNumFmt(numFmtId);
    if (!r.status.ok) return null;
    return r.formatCode;
  }

  /** Add or dedup a font record. Returns the resolved font index, or -1 on
   *  engine failure or when capability is off. */
  addFontRecord(record: FontRecord): number {
    this.assertAlive();
    if (!this.capabilities.cellFormatting) return -1;
    const r = this.wb.addFont(record);
    return r.status.ok ? r.index : -1;
  }

  /** Add or dedup a fill record. */
  addFillRecord(record: FillRecord): number {
    this.assertAlive();
    if (!this.capabilities.cellFormatting) return -1;
    const r = this.wb.addFill(record);
    return r.status.ok ? r.index : -1;
  }

  /** Add or dedup a border record. */
  addBorderRecord(record: BorderRecord): number {
    this.assertAlive();
    if (!this.capabilities.cellFormatting) return -1;
    const r = this.wb.addBorder(record);
    return r.status.ok ? r.index : -1;
  }

  /** Register a number-format code. Built-in matches return the built-in id;
   *  custom codes are appended starting at 164. Returns -1 on failure. */
  addNumFmtCode(formatCode: string): number {
    this.assertAlive();
    if (!this.capabilities.cellFormatting) return -1;
    const r = this.wb.addNumFmt(formatCode);
    return r.status.ok ? r.numFmtId : -1;
  }

  /** Add or dedup an XF (eXtended Format) record built from existing
   *  font/fill/border indices and a registered numFmtId. Returns the resolved
   *  xf index, or -1 on failure. */
  addXfRecord(record: CellXf): number {
    this.assertAlive();
    if (!this.capabilities.cellFormatting) return -1;
    const r = this.wb.addXf(record);
    return r.status.ok ? r.index : -1;
  }

  /** Append `range` as a merge on `sheet`. Returns false on engine failure or
   *  when `capabilities.merges` is off. The cell content inside the range is
   *  the caller's responsibility (Excel keeps top-left, blanks the rest). */
  engineAddMerge(sheet: number, range: Range): boolean {
    this.assertAlive();
    if (!this.capabilities.merges) return false;
    const s = this.wb.addMerge(sheet, {
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
    this.assertAlive();
    if (!this.capabilities.merges) return false;
    const s = this.wb.removeMerge(sheet, {
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
    this.assertAlive();
    if (!this.capabilities.merges) return false;
    const s = this.wb.clearMerges(sheet);
    return s.ok;
  }

  /** Snapshot of every merge on `sheet` as inclusive `Range` records. Empty
   *  array under stub or when `capabilities.merges` is off. */
  getMerges(sheet: number): Range[] {
    this.assertAlive();
    if (!this.capabilities.merges) return [];
    const arr = this.wb.getMerges(sheet);
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
    this.assertAlive();
    if (!this.capabilities.comments) return null;
    const e = this.wb.getComment(sheet, row, col);
    return e ? { author: e.author, text: e.text } : null;
  }

  /** Persist a cell comment. Empty `text` removes it. No-op (returns false)
   *  under the stub. */
  setCommentEntry(sheet: number, row: number, col: number, author: string, text: string): boolean {
    this.assertAlive();
    if (!this.capabilities.comments) return false;
    const s = this.wb.setComment(sheet, row, col, author, text);
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
    this.assertAlive();
    if (!this.capabilities.conditionalFormat) return [];
    const r = this.wb.evaluateCfRange(sheet, firstRow, firstCol, lastRow, lastCol, todaySerial);
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
    this.assertAlive();
    if (!this.capabilities.spillInfo) return null;
    const r = this.wb.spillInfo(sheet, row, col);
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
    this.assertAlive();
    if (!this.capabilities.spillInfo) return null;
    return computeEngineSpillRanges(this, sheet);
  }

  /** Cells that `addr` directly reads (1-step precedents) by default;
   *  pass `depth > 1` for a BFS expansion (engine caps at 32 to avoid
   *  runaway in cyclic graphs). Includes cross-sheet refs — callers that
   *  only want same-sheet relations should filter on `sheet`. Returns
   *  `null` when the engine doesn't expose `precedents`; the regex-based
   *  same-sheet fallback in `engine/refs-graph.ts` covers stub mode. */
  precedents(addr: Addr, depth = 1): Addr[] | null {
    this.assertAlive();
    if (!this.capabilities.traceArrows) return null;
    const arr = this.wb.precedents(addr.sheet, addr.row, addr.col, depth);
    return arr.map((n) => ({ sheet: n.sheet, row: n.row, col: n.col }));
  }

  /** Cells whose formulas read from `addr` (1-step dependents by default).
   *  Same depth + cross-sheet semantics as `precedents`. Returns `null`
   *  when the engine doesn't expose `dependents`. */
  dependents(addr: Addr, depth = 1): Addr[] | null {
    this.assertAlive();
    if (!this.capabilities.traceArrows) return null;
    const arr = this.wb.dependents(addr.sheet, addr.row, addr.col, depth);
    return arr.map((n) => ({ sheet: n.sheet, row: n.row, col: n.col }));
  }

  /** Every registered function's canonical name in ascending sort order.
   *  Returns `null` when the engine doesn't expose `functionNames`; the
   *  static `FUNCTION_NAMES` list in `commands/refs.ts` is the fallback
   *  catalog under stub mode. */
  functionNames(): readonly string[] | null {
    this.assertAlive();
    if (!this.capabilities.functionMetadata) return null;
    return this.wb.functionNames();
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
    this.assertAlive();
    if (!this.capabilities.functionMetadata) return null;
    const m = this.wb.functionMetadata(name, locale);
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
    this.assertAlive();
    if (!this.capabilities.functionLocale) return null;
    return this.wb.localizeFunctionName(canonicalName, locale);
  }

  /** Localized → canonical function-name lookup. Falls through to a
   *  case-insensitive match on the canonical name when no alias is
   *  registered. Returns the empty string when the engine reports no
   *  matching function. Returns `null` when the engine doesn't expose
   *  `canonicalizeFunctionName`. */
  canonicalizeFunctionName(localizedName: string, locale = 0): string | null {
    this.assertAlive();
    if (!this.capabilities.functionLocale) return null;
    return this.wb.canonicalizeFunctionName(localizedName, locale);
  }

  /** Workbook calc-mode metadata mirroring `<calcPr calcMode>`. The engine
   *  itself does NOT gate evaluation on this value — every `recalc()` call
   *  honours all dirty cells regardless of mode. The flag is preserved as
   *  round-trip metadata and surfaced here so the UI can mirror Excel's
   *  user-visible state. Returns `null` when the engine doesn't expose
   *  `calcMode`. Codes: 0 = Auto, 1 = Manual, 2 = AutoNoTable. */
  calcMode(): 0 | 1 | 2 | null {
    this.assertAlive();
    if (!this.capabilities.calcMode) return null;
    const mode = this.wb.calcMode();
    return (mode as 0 | 1 | 2) ?? null;
  }

  /** Sets the calc-mode metadata. Returns `false` (no-op) under stub or
   *  pre-5/5 vendored builds. */
  setCalcMode(mode: 0 | 1 | 2): boolean {
    this.assertAlive();
    if (!this.capabilities.calcMode) return false;
    return this.wb.setCalcMode(mode).ok;
  }

  /** External-link records carried by the workbook in `<externalReferences>`
   *  document order. Empty for fresh workbooks and packages whose source
   *  archive had no `<externalReferences>` block. Returns `[]` when the
   *  engine doesn't expose `getExternalLinks`. */
  getExternalLinks(): ReadonlyArray<{
    index: number;
    relId: string;
    partPath: string;
    target: string;
    targetExternal: boolean;
    kind: 'unknown' | 'externalBook' | 'ole' | 'dde';
  }> {
    this.assertAlive();
    if (!this.capabilities.externalLinks) return [];
    const arr = this.wb.getExternalLinks();
    return arr.map((r) => ({
      index: r.index,
      relId: r.relId,
      partPath: r.partPath,
      target: r.target,
      targetExternal: r.targetExternal,
      kind: kindLabel(r.kind),
    }));
  }

  /** Snapshot of every validation entry on `sheet`. Each entry can apply to
   *  multiple ranges (`ranges`) and carries an Excel-style descriptor: numeric
   *  `type` ordinal (0 none, 1 whole, 2 decimal, 3 list, 4 date, 5 time,
   *  6 textLength, 7 custom), numeric `op` ordinal (0 between … 7 lessThanOrEqual),
   *  formula1/2 strings, and the surrounding error/prompt metadata. Empty
   *  when `capabilities.dataValidation` is off or when the engine returns
   *  no rules. */
  getValidationsForSheet(sheet: number): {
    ranges: Range[];
    type: number;
    op: number;
    errorStyle: number;
    allowBlank: boolean;
    showInputMessage: boolean;
    showErrorMessage: boolean;
    formula1: string;
    formula2: string;
    errorTitle: string;
    errorMessage: string;
    promptTitle: string;
    promptMessage: string;
  }[] {
    this.assertAlive();
    if (!this.capabilities.dataValidation) return [];
    const arr = this.wb.getValidations(sheet);
    return arr.map((v) => ({
      ranges: v.ranges.map((m) => ({
        sheet,
        r0: m.firstRow,
        c0: m.firstCol,
        r1: m.lastRow,
        c1: m.lastCol,
      })),
      type: v.type,
      op: v.op,
      errorStyle: v.errorStyle,
      allowBlank: v.allowBlank,
      showInputMessage: v.showInputMessage,
      showErrorMessage: v.showErrorMessage,
      formula1: v.formula1,
      formula2: v.formula2,
      errorTitle: v.errorTitle,
      errorMessage: v.errorMessage,
      promptTitle: v.promptTitle,
      promptMessage: v.promptMessage,
    }));
  }

  /** Append a data-validation rule to `sheet`. Ranges are inclusive `Range`
   *  records; everything else mirrors the upstream `DataValidationInput` shape.
   *  Returns false on engine failure or when capability is off. */
  addValidationEntry(
    sheet: number,
    input: {
      ranges: Range[];
      type: number;
      op?: number;
      errorStyle?: number;
      allowBlank?: boolean;
      showInputMessage?: boolean;
      showErrorMessage?: boolean;
      formula1?: string;
      formula2?: string;
      errorTitle?: string;
      errorMessage?: string;
      promptTitle?: string;
      promptMessage?: string;
    },
  ): boolean {
    this.assertAlive();
    if (!this.capabilities.dataValidation) return false;
    const dv: DataValidationInput = {
      type: input.type,
      ranges: input.ranges.map((r) => ({
        firstRow: r.r0,
        firstCol: r.c0,
        lastRow: r.r1,
        lastCol: r.c1,
      })),
      ...(input.op !== undefined ? { op: input.op } : {}),
      ...(input.errorStyle !== undefined ? { errorStyle: input.errorStyle } : {}),
      ...(input.allowBlank !== undefined ? { allowBlank: input.allowBlank } : {}),
      ...(input.showInputMessage !== undefined ? { showInputMessage: input.showInputMessage } : {}),
      ...(input.showErrorMessage !== undefined ? { showErrorMessage: input.showErrorMessage } : {}),
      ...(input.formula1 !== undefined ? { formula1: input.formula1 } : {}),
      ...(input.formula2 !== undefined ? { formula2: input.formula2 } : {}),
      ...(input.errorTitle !== undefined ? { errorTitle: input.errorTitle } : {}),
      ...(input.errorMessage !== undefined ? { errorMessage: input.errorMessage } : {}),
      ...(input.promptTitle !== undefined ? { promptTitle: input.promptTitle } : {}),
      ...(input.promptMessage !== undefined ? { promptMessage: input.promptMessage } : {}),
    };
    const s = this.wb.addValidation(sheet, dv);
    return s.ok;
  }

  /** Remove the validation rule at `index` on `sheet`. */
  removeValidationAt(sheet: number, index: number): boolean {
    this.assertAlive();
    if (!this.capabilities.dataValidation) return false;
    const s = this.wb.removeValidationAt(sheet, index);
    return s.ok;
  }

  /** Drop every validation rule on `sheet`. */
  clearValidations(sheet: number): boolean {
    this.assertAlive();
    if (!this.capabilities.dataValidation) return false;
    const s = this.wb.clearValidations(sheet);
    return s.ok;
  }

  /** Snapshot of every hyperlink on `sheet`. Empty array under the stub. */
  getHyperlinks(
    sheet: number,
  ): { row: number; col: number; target: string; display: string; tooltip: string }[] {
    this.assertAlive();
    if (!this.capabilities.hyperlinks) return [];
    const arr = this.wb.getHyperlinks(sheet);
    return arr.map((h) => ({
      row: h.row,
      col: h.col,
      target: h.target,
      display: h.display,
      tooltip: h.tooltip,
    }));
  }

  /** Append a hyperlink at `(sheet, row, col)`. Empty `display` / `tooltip`
   *  mean default. Returns false on engine failure or capability off. */
  addHyperlink(
    sheet: number,
    row: number,
    col: number,
    target: string,
    display = '',
    tooltip = '',
  ): boolean {
    this.assertAlive();
    if (!this.capabilities.hyperlinks) return false;
    const s = this.wb.addHyperlink(sheet, row, col, target, display, tooltip);
    return s.ok;
  }

  /** Remove every hyperlink anchored at `(sheet, row, col)`. No-op when none
   *  match. Returns false on engine failure or capability off. */
  removeHyperlink(sheet: number, row: number, col: number): boolean {
    this.assertAlive();
    if (!this.capabilities.hyperlinks) return false;
    const s = this.wb.removeHyperlink(sheet, row, col);
    return s.ok;
  }

  /** Remove the hyperlink at `index` on `sheet`. */
  removeHyperlinkAt(sheet: number, index: number): boolean {
    this.assertAlive();
    if (!this.capabilities.hyperlinks) return false;
    const s = this.wb.removeHyperlinkAt(sheet, index);
    return s.ok;
  }

  /** Drop every hyperlink on `sheet`. */
  clearHyperlinks(sheet: number): boolean {
    this.assertAlive();
    if (!this.capabilities.hyperlinks) return false;
    const s = this.wb.clearHyperlinks(sheet);
    return s.ok;
  }

  /** Snapshot of row overrides on `sheet`. See `getColumnLayouts`. */
  getRowLayouts(
    sheet: number,
  ): { row: number; height: number; hidden: boolean; outlineLevel: number }[] {
    this.assertAlive();
    if (!this.capabilities.colRowSize) return [];
    const r = this.wb.getSheetRowOverrides(sheet);
    const out: { row: number; height: number; hidden: boolean; outlineLevel: number }[] = [];
    if (!r.status.ok) return out;
    const v = r.rows;
    try {
      const n = v.size();
      for (let i = 0; i < n; i += 1) {
        const e = v.get(i);
        out.push({
          row: e.row,
          height: e.height,
          hidden: e.hidden !== 0,
          outlineLevel: e.outlineLevel,
        });
      }
    } finally {
      v.delete();
    }
    return out;
  }

  /** Snapshot of every Excel Table on the workbook. Read-only in the engine —
   *  we surface it as a badge count + listing for the status bar. Empty array
   *  on the stub. */
  getTables(): {
    name: string;
    displayName: string;
    ref: string;
    sheetIndex: number;
    columns: string[];
  }[] {
    this.assertAlive();
    if (!this.wb.tableCount) return [];
    const n = this.wb.tableCount();
    const out: {
      name: string;
      displayName: string;
      ref: string;
      sheetIndex: number;
      columns: string[];
    }[] = [];
    for (let i = 0; i < n; i += 1) {
      const e = this.wb.tableAt(i);
      if (!e.status.ok) continue;
      out.push({
        name: e.name,
        displayName: e.displayName,
        ref: e.ref,
        sheetIndex: e.sheetIndex,
        columns: this.tableColumnNames(e.sheetIndex, e.ref),
      });
    }
    return out;
  }

  /** Derive column display names from the header row of a table's `ref`.
   *  Returns labels in source order; cells that read blank fall back to the
   *  Excel-style `Column1` / `Column2` placeholder so the structured-ref
   *  autocomplete still has something to insert. */
  private tableColumnNames(sheet: number, ref: string): string[] {
    const parsed = parseTableRef(ref);
    if (!parsed) return [];
    const out: string[] = [];
    for (let col = parsed.c0; col <= parsed.c1; col += 1) {
      const v = this.getValue({ sheet, row: parsed.r0, col });
      const text = formatCell(v);
      out.push(text || `Column${out.length + 1}`);
    }
    return out;
  }

  /** Snapshot of OOXML "passthrough" parts (charts, drawings, pivots, etc.)
   *  preserved verbatim by the engine. Surfaced as a badge so users know
   *  these objects exist even though the UI doesn't render them. */
  getPassthroughs(): { path: string }[] {
    this.assertAlive();
    if (!this.wb.passthroughCount) return [];
    const n = this.wb.passthroughCount();
    const out: { path: string }[] = [];
    for (let i = 0; i < n; i += 1) {
      const e = this.wb.passthroughAt(i);
      if (!e.status.ok) continue;
      out.push({ path: e.path });
    }
    return out;
  }

  subscribe(fn: ChangeListener): () => void {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  }

  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    this.listeners.clear();
    this.wb.delete();
  }

  private emit(e: ChangeEvent): void {
    for (const fn of this.listeners) fn(e);
  }

  private assertAlive(): void {
    if (this.disposed) throw new Error('WorkbookHandle is disposed');
  }

  private captureSnapshot(a: Addr): CellSnapshot {
    return { addr: a, value: this.getValue(a), formula: this.cellFormula(a) };
  }

  /** Capture before/after, run the mutation, and push to the active history.
   *  When a shared History is attached we capture a closure pair; otherwise
   *  fall back to the local snapshot stack. */
  private withJournal(a: Addr, fn: () => void): void {
    if (this.replaying) {
      fn();
      return;
    }
    const before = this.captureSnapshot(a);
    fn();
    if (this.history) {
      const after = this.captureSnapshot(a);
      this.history.push({
        undo: () => this.replay(before),
        redo: () => this.replay(after),
      });
      return;
    }
    this.undoStack.push(before);
    if (this.undoStack.length > UNDO_LIMIT) this.undoStack.shift();
    this.redoStack.length = 0;
  }

  /** Restore the cell to a captured snapshot without journaling. */
  private replay(snap: CellSnapshot): void {
    this.replaying = true;
    try {
      if (snap.formula) {
        this.setFormula(snap.addr, snap.formula);
        return;
      }
      switch (snap.value.kind) {
        case 'number':
          this.setNumber(snap.addr, snap.value.value);
          return;
        case 'text':
          this.setText(snap.addr, snap.value.value);
          return;
        case 'bool':
          this.setBool(snap.addr, snap.value.value);
          return;
        default:
          this.setBlank(snap.addr);
      }
    } finally {
      this.replaying = false;
    }
  }
}

export { addrKey };
