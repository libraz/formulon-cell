import {
  ENGINE_SPREADSHEET_PROFILE_GETTER,
  ENGINE_SPREADSHEET_PROFILE_SETTER,
} from './capabilities.js';
import { localeTag, mergeFunctionMetadata } from './function-metadata.js';
import { numberValue } from './result.js';
import { computeEngineSpillRanges } from './spill.js';
import {
  type EngineSpreadsheetProfileId,
  engineProfileToPublic,
  publicProfileToEngine,
} from './spreadsheet-profile.js';
import type {
  Addr,
  EngineCapabilities,
  EvalArrayResult,
  EvalResult,
  FormulonModule,
  FunctionMetadataProvider,
  Range,
  SpreadsheetProfileId,
  Workbook,
} from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

type WorkbookHandleCtor = { prototype: WorkbookHandle };
type WorkbookHandleInternals = {
  wb: Workbook;
  module: FormulonModule;
  functionMetadataProvider: FunctionMetadataProvider | null;
  assertAlive(): void;
};

declare module './workbook-handle.js' {
  interface WorkbookHandle extends WorkbookHandleFormulasMethods {}
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

export abstract class WorkbookHandleFormulasMethods {
  declare readonly capabilities: EngineCapabilities;

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

  /** Toggle the iterative-formula solver. `maxIterations` and `maxChange`
   *  cap the Gauss-Seidel loop; matches "File → Options → Formulas"
   *  knobs. No-op (returns false) on engines without the iterative surface. */
  setIterative(enabled: boolean, maxIterations: number, maxChange: number): boolean {
    assertAlive(this);
    if (!this.capabilities.iterativeProgress) return false;
    const s = wb(this).setIterative(enabled, maxIterations, maxChange);
    return s.ok;
  }

  /** Read back the workbook's stored iterative-calculation settings. The cap
   *  and threshold are meaningful even while `enabled` is false, so a dialog
   *  can open on what the workbook carries rather than on its own defaults.
   *
   *  Returns null when the engine has no readback, in which case the caller
   *  keeps whatever it last wrote. Note the engine clamps `maxIterations` to
   *  32767 on the way in, so a larger request reads back clamped. */
  getIterative(): { enabled: boolean; maxIterations: number; maxChange: number } | null {
    assertAlive(this);
    if (!this.capabilities.iterativeSettings) return null;
    const r = wb(this).getIterative();
    if (!r.status.ok) return null;
    return { enabled: r.enabled, maxIterations: r.maxIterations, maxChange: r.maxChange };
  }

  /** Install (or clear) a progress callback invoked after each iterative-solve
   *  sweep. Returning `false` from the callback aborts the solve. Pass `null`
   *  to detach. No-op on engines without `setIterativeProgress`. */
  setIterativeProgress(
    callback:
      | ((iteration: number, maxResidual: number, maxIterations: number) => boolean | void)
      | null,
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.iterativeProgress) return false;
    const s = wb(this).setIterativeProgress(callback);
    return s.ok;
  }

  /** Renders the lambda value stored at `addr` as spreadsheet formula text. The
   *  returned string never carries a leading `=` — callers prepending it
   *  for the formula-bar edit seed should add the prefix themselves.
   *  Returns `null` when the engine doesn't expose `getLambdaText` or
   *  when the cell is absent / its cached value is not a lambda. */
  getLambdaText(addr: Addr): string | null {
    assertAlive(this);
    if (!this.capabilities.lambdaText) return null;
    const r = wb(this).getLambdaText(addr.sheet, addr.row, addr.col);
    if (!r.status.ok) return null;
    return r.text || null;
  }

  /** Iterate over defined names. `localSheetId === -1` means workbook scope;
   *  otherwise it is the 0-based sheet index for a sheet-scoped name. */
  *definedNames(): Generator<{ name: string; formula: string; localSheetId: number }> {
    assertAlive(this);
    const n = numberValue(wb(this).definedNameCount(), 'definedNameCount');
    for (let i = 0; i < n; i += 1) {
      const e = wb(this).definedNameAt(i);
      if (!e.status.ok || !e.name || e.formula === undefined || e.localSheetId === undefined)
        continue;
      yield { name: e.name, formula: e.formula, localSheetId: e.localSheetId };
    }
  }

  /** Add or replace a workbook/sheet-scoped defined name. Pass an empty `formula`
   *  to remove the name (engine convention). Returns false on engine failure
   *  or when the engine doesn't expose scoped names. */
  setDefinedNameEntry(name: string, formula: string, localSheetId = -1): boolean {
    assertAlive(this);
    if (!this.capabilities.definedNameMutate) return false;
    const s = wb(this).setDefinedNameScoped(name, formula, localSheetId);
    return s.ok;
  }
}

export function installFormulasMethods(target: WorkbookHandleCtor): void {
  for (const key of Object.getOwnPropertyNames(WorkbookHandleFormulasMethods.prototype)) {
    if (key === 'constructor') continue;
    const descriptor = Object.getOwnPropertyDescriptor(
      WorkbookHandleFormulasMethods.prototype,
      key,
    );
    if (!descriptor) continue;
    Object.defineProperty(target.prototype, key, descriptor);
  }
}
