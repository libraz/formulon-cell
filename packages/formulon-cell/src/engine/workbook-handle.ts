import type { History } from '../commands/history.js';
import { addrKey } from './address.js';
import { detectCapabilities } from './capabilities.js';
import { type ExternalLinkKind, externalLinkKindLabel } from './external-links.js';
import type { LoadOptions } from './loader.js';
import { isUsingStub, loadFormulon } from './loader.js';
import { parseRangeRef as parseTableRef } from './range-resolver.js';
import type {
  Addr,
  CellValue,
  DataValidationInput,
  EngineCapabilities,
  FormulonModule,
  Range,
  Workbook,
} from './types.js';
import { formatCell, fromEngineValue } from './value.js';
import { installWorkbookFeatureMethods } from './workbook-handle-features.js';
import { installPivotMethods } from './workbook-handle-pivot.js';

export type ChangeListener = (e: ChangeEvent) => void;

export type ChangeEvent =
  | { kind: 'value'; addr: Addr; next: CellValue }
  | { kind: 'recalc'; dirty: ReadonlySet<string> }
  | { kind: 'sheet-add'; index: number; name: string }
  | { kind: 'sheet-rename'; index: number; name: string }
  | { kind: 'sheet-remove'; index: number }
  | { kind: 'sheet-move'; from: number; to: number };

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
   *  cap the Gauss-Seidel loop; matches "File → Options → Formulas"
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

  /** Iterate over every populated cell on a sheet. Used for initial paint.
   *  Loaded PivotTables are projected after physical cells so the evaluated
   *  layout is what the grid displays when a pivot overlaps cached values. */
  *cells(sheet: number): Generator<{ addr: Addr; value: CellValue; formula: string | null }> {
    yield* this.physicalCells(sheet);
    yield* this.pivotCells(sheet);
  }

  /** Iterate over cells physically stored by the workbook model, excluding
   *  evaluated overlays such as PivotTable projections. */
  *physicalCells(
    sheet: number,
  ): Generator<{ addr: Addr; value: CellValue; formula: string | null }> {
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

  /** Renders the lambda value stored at `addr` as spreadsheet formula text. The
   *  returned string never carries a leading `=` — callers prepending it
   *  for the formula-bar edit seed should add the prefix themselves.
   *  Returns `null` when the engine doesn't expose `getLambdaText` or
   *  when the cell is absent / its cached value is not a lambda. */
  getLambdaText(addr: Addr): string | null {
    this.assertAlive();
    if (!this.capabilities.lambdaText) return null;
    const r = this.wb.getLambdaText(addr.sheet, addr.row, addr.col);
    if (!r.status.ok) return null;
    return r.text || null;
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
    kind: ExternalLinkKind;
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
      kind: externalLinkKindLabel(r.kind),
    }));
  }

  /** Snapshot of every validation entry on `sheet`. Each entry can apply to
   *  multiple ranges (`ranges`) and carries an spreadsheet-style descriptor: numeric
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

  /** Snapshot of every spreadsheet Table on the workbook. Read-only in the engine —
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
   *  Spreadsheet-style `Column1` / `Column2` placeholder so the structured-ref
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

installPivotMethods(WorkbookHandle);
installWorkbookFeatureMethods(WorkbookHandle);

export { addrKey };
