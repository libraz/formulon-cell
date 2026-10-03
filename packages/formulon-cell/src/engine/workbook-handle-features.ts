import {
  ENGINE_SPREADSHEET_PROFILE_GETTER,
  ENGINE_SPREADSHEET_PROFILE_SETTER,
} from './capabilities.js';
import { localeTag, mergeFunctionMetadata } from './function-metadata.js';
import { computeEngineSpillRanges } from './spill.js';
import {
  type EngineSpreadsheetProfileId,
  engineProfileToPublic,
  publicProfileToEngine,
} from './spreadsheet-profile.js';
import { completeDxfRecord } from './style-records.js';
import type {
  Addr,
  CellValue,
  ConditionalFormatEntry,
  ConditionalFormatInput,
  DxfRecord,
  EngineCapabilities,
  EvalArrayResult,
  EvalResult,
  FormulonModule,
  FunctionMetadataProvider,
  PhoneticRun,
  Range,
  SpreadsheetProfileId,
  TableInput,
  Workbook,
} from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

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
  getComments?: (
    sheet: number,
  ) => readonly EngineCommentEntry[] & { readonly status: { ok: boolean } };
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

  /** Read the cell's OOXML phonetic guide, if the current engine exposes it. */
  getCellPhonetic(sheet: number, row: number, col: number): string | null {
    assertAlive(this);
    if (!this.capabilities.phonetic) return null;
    const r = wb(this).getCellPhonetic(sheet, row, col);
    return r.status.ok && r.value ? r.value : null;
  }

  /** Set (or, with an empty string, clear) the cell's phonetic guide. The
   *  engine spans the whole cell text, so this replaces any per-run guide the
   *  cell carried — use `setCellPhoneticRuns` to preserve the spans. */
  setCellPhonetic(sheet: number, row: number, col: number, phonetic: string): boolean {
    assertAlive(this);
    if (!this.capabilities.phonetic) return false;
    return wb(this).setCellPhonetic(sheet, row, col, phonetic).ok;
  }

  /** Read the cell's phonetic guide span by span. Returns null when the engine
   *  has no per-run surface, which is distinct from the empty array an
   *  unannotated cell reports. */
  getCellPhoneticRuns(sheet: number, row: number, col: number): PhoneticRun[] | null {
    assertAlive(this);
    if (!this.capabilities.phoneticRuns) return null;
    const r = wb(this).getCellPhoneticRuns(sheet, row, col);
    if (!r.status.ok) return null;
    return r.runs.map((run) => ({ start: run.sb, end: run.eb, text: run.text }));
  }

  /** Replace the cell's phonetic guide with `runs`, an ordered partition of the
   *  cell text. An empty array clears the guide. */
  setCellPhoneticRuns(
    sheet: number,
    row: number,
    col: number,
    runs: readonly PhoneticRun[],
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.phoneticRuns) return false;
    return wb(this).setCellPhoneticRuns(
      sheet,
      row,
      col,
      runs.map((run) => ({ sb: run.start, eb: run.end, text: run.text })),
    ).ok;
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
    const entries = engineWb.getComments(sheet);
    if (!entries.status.ok) return [];
    return entries.map((e) => ({
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
      barDirection: number;
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
          barDirection: m.barDirection,
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
    if (!r.status.ok || !r.engaged) return null;
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
    if (!arr.status.ok) return null;
    return arr.map((n) => ({ sheet: n.sheet, row: n.row, col: n.col }));
  }

  /** Cells whose formulas read from `addr` (1-step dependents by default).
   *  Same depth + cross-sheet semantics as `precedents`. Returns `null`
   *  when the engine doesn't expose `dependents`. */
  dependents(addr: Addr, depth = 1): Addr[] | null {
    assertAlive(this);
    if (!this.capabilities.traceArrows) return null;
    const arr = wb(this).dependents(addr.sheet, addr.row, addr.col, depth);
    if (!arr.status.ok) return null;
    return arr.map((n) => ({ sheet: n.sheet, row: n.row, col: n.col }));
  }

  /** Every registered function's canonical name in ascending sort order.
   *  Returns `null` when the engine doesn't expose `functionNames`; the
   *  static `FUNCTION_NAMES` list in `commands/refs.ts` is the fallback
   *  catalog under stub mode. */
  functionNames(): readonly string[] | null {
    assertAlive(this);
    if (!this.capabilities.functionMetadata) return null;
    const names = wb(this).functionNames();
    return names.status.ok ? [...names] : null;
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
    /** Native availability class; absent when an older engine omits it. */
    availability?: number;
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
      ...(m.availability !== undefined ? { availability: m.availability } : {}),
      ...(merged.signatureTemplate ? { signatureTemplate: merged.signatureTemplate } : {}),
      ...(merged.description ? { description: merged.description } : {}),
      ...(merged.localizedName ? { localizedName: merged.localizedName } : {}),
    };
  }

  /** Canonical → localized function-name lookup. `locale`: 0 = en-US,
   *  1 = ja-JP. Returns the canonical name unchanged when no alias is
   *  registered for `locale` (currently the case for every locale except
   *  en-US). Returns the empty string for an unsuccessful lookup.
   *  Returns `null` when the engine doesn't expose
   *  `localizeFunctionName`. */
  localizeFunctionName(canonicalName: string, locale = 0): string | null {
    assertAlive(this);
    if (!this.capabilities.functionLocale) return null;
    const result = wb(this).localizeFunctionName(canonicalName, locale);
    return result.status.ok ? result.value : '';
  }

  /** Localized → canonical function-name lookup. Falls through to a
   *  case-insensitive match on the canonical name when no alias is
   *  registered. Returns the empty string when the engine reports no
   *  matching function. Returns `null` when the engine doesn't expose
   *  `canonicalizeFunctionName`. */
  canonicalizeFunctionName(localizedName: string, locale = 0): string | null {
    assertAlive(this);
    if (!this.capabilities.functionLocale) return null;
    const result = wb(this).canonicalizeFunctionName(localizedName, locale);
    return result.status.ok ? result.value : '';
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
    return mode.status.ok ? mode.value : null;
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
      wb(this) as unknown as Record<
        string,
        ((this: Workbook) => { status: { ok: boolean }; value: string }) | undefined
      >
    )[ENGINE_SPREADSHEET_PROFILE_GETTER];
    if (!getProfile) return null;
    const result = getProfile.call(wb(this));
    if (!result.status.ok) return null;
    return engineProfileToPublic(result.value as EngineSpreadsheetProfileId);
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

  /** Snapshot of every CF rule on `sheet`, in flattened priority order.
   *  Returns `[]` when the engine doesn't expose `getConditionalFormats`
   *  or when there are no rules. The entries borrow rule ids from the
   *  engine's storage; treat them as immutable view objects. */
  getConditionalFormats(sheet: number): readonly ConditionalFormatEntry[] {
    assertAlive(this);
    if (!this.capabilities.conditionalFormatMutate) return [];
    const entries = wb(this).getConditionalFormats(sheet);
    return entries.status.ok ? [...entries] : [];
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
    const count = wb(this).dxfCount();
    return count.status.ok ? count.value : 0;
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
