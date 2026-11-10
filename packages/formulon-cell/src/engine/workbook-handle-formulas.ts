import { type CalcMode, type ProfileId, RefMode } from '@libraz/formulon';
import { CALC_MODE_MANUAL, type CalcModeCode } from './calc-mode.js';
import {
  canonicalizeFormula,
  canonicalizeFormulaForEngine,
  type LocaleOrdinal,
  localeOrdinalForSpreadsheetProfile,
  localizeFormula,
  spreadsheetProfileForLocaleOrdinal,
} from './function-locale.js';
import {
  displayLocaleTag,
  localeOrdinalForTag,
  localeTag,
  mergeFunctionMetadata,
} from './function-metadata.js';
import { isFormulonError } from './result.js';
import { computeEngineSpillRanges } from './spill.js';
import {
  type EngineSpreadsheetProfileId,
  engineProfileToPublic,
  publicProfileToEngine,
} from './spreadsheet-profile.js';
import type {
  Addr,
  EngineCapabilities,
  Formulon,
  FunctionMetadataProvider,
  LocaleFacts,
  Range,
  Sheet,
  SpreadsheetProfileId,
  Value,
  Workbook,
} from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

type WorkbookHandleCtor = { prototype: WorkbookHandle };
type WorkbookHandleInternals = {
  wb: Workbook;
  module: Formulon;
  functionMetadataProvider: FunctionMetadataProvider | null;
  assertAlive(): void;
};

declare module './workbook-handle.js' {
  interface WorkbookHandle extends WorkbookHandleFormulasMethods {}
}

/** Workbook calculation options persisted in `<calcPr>` and `<workbookPr>`. */
export interface WorkbookCalcSettings {
  /** `false` is "Set precision as displayed". */
  fullPrecision: boolean;
  /** Recalculate before saving. */
  calcOnSave: boolean;
  /** R1C1 reference style (`refMode`). */
  r1c1: boolean;
  /** 1904 date system. */
  date1904: boolean;
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

function moduleOf(handle: unknown): Formulon {
  return internals(handle).module;
}

function ws(handle: unknown, sheet: number): Sheet {
  return internals(handle).wb.sheet(sheet);
}

function engineProfile(profileId: SpreadsheetProfileId): ProfileId {
  return publicProfileToEngine(profileId) as ProfileId;
}

/** Run an engine call; null when the engine rejects it. */
function attempt<T>(call: () => T): T | null {
  try {
    return call();
  } catch (e) {
    if (!isFormulonError(e)) throw e;
    return null;
  }
}

/** Run an engine write; false when the engine rejects it. */
function attemptWrite(write: () => void): boolean {
  try {
    write();
    return true;
  } catch (e) {
    if (!isFormulonError(e)) throw e;
    return false;
  }
}

/** Defined-name formula text as `definedNames()` reports it: without the `=`. */
function nameFormulaText(formula: string): string {
  return formula.startsWith('=') ? formula.slice(1) : formula;
}

export abstract class WorkbookHandleFormulasMethods {
  declare readonly capabilities: EngineCapabilities;

  /** Evaluate a self-contained formula string (references already substituted
   *  with literal values) in a fresh engine workbook. Callers that need
   *  live-cell context must substitute refs first (see the custom
   *  data-validation evaluator and the F9 / evaluate-formula surfaces). Formula
   *  function aliases are converted from the active German/French profile
   *  before evaluation. Null when the engine cannot evaluate the formula. */
  evalFormula(formula: string): Value | null {
    assertAlive(this);
    if (!this.capabilities.formulaTextEvaluation) return null;
    const canonical = this.formulaForEvaluation(formula);
    const module = moduleOf(this);
    const profile = this.capabilities.spreadsheetProfile ? wb(this).getProfileId() : null;
    if (!profile) return attempt(() => module.evaluateTypedFormula(canonical));
    let scratch: Workbook | null = null;
    try {
      scratch = module.Workbook.createDefault();
      scratch.setProfileId(profile);
      const fresh = scratch;
      return attempt(() => fresh.sheet(0).evaluateTypedFormula(0, 0, canonical));
    } finally {
      scratch?.dispose();
    }
  }

  /** Evaluate `formula` as if entered at `addr`, without mutating the
   *  workbook. Null when the engine cannot evaluate it. */
  evaluateFormulaText(addr: Addr, formula: string): Value | null {
    assertAlive(this);
    if (!this.capabilities.formulaTextEvaluation) return null;
    const canonical = this.formulaForEvaluation(formula);
    return attempt(() => ws(this, addr.sheet).evaluateTypedFormula(addr.row, addr.col, canonical));
  }

  /** Ad-hoc evaluation that returns the *whole* dynamic-array / spilled
   *  result of `formula` anchored at `addr`, without mutating the workbook.
   *  The grid is row-major (`cells[r][c]`); a scalar is reported as a 1x1
   *  array. Null when the engine cannot evaluate it. */
  evaluateFormulaArray(addr: Addr, formula: string): (readonly Value[])[] | null {
    assertAlive(this);
    if (!this.capabilities.arrayFormulaEvaluation) {
      const single = this.evaluateFormulaText(addr, formula);
      return single ? [[single]] : null;
    }
    const canonical = this.formulaForEvaluation(formula);
    const rows = attempt(() =>
      ws(this, addr.sheet).evaluateTypedFormulaArray(addr.row, addr.col, canonical),
    );
    return rows ? [...rows] : null;
  }

  /** Evaluate a conditional-formatting predicate at `addr` with relative
   *  references written from `anchor`, coerced with the CF rules. Null when
   *  the engine cannot evaluate it. */
  evaluateConditionalFormula(
    addr: Addr,
    anchor: Pick<Addr, 'row' | 'col'>,
    formula: string,
  ): boolean | null {
    assertAlive(this);
    if (!this.capabilities.conditionalFormulaEvaluation) return null;
    const canonical = this.formulaForEvaluation(formula);
    return attempt(() =>
      ws(this, addr.sheet).evaluateCfFormula(addr.row, addr.col, anchor.row, anchor.col, canonical),
    );
  }

  private formulaForEvaluation(formula: string): string {
    return canonicalizeFormulaForEngine(
      this as unknown as WorkbookHandle,
      formula,
      localeOrdinalForSpreadsheetProfile(this.spreadsheetProfileId()),
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
    const r = attempt(() => ws(this, sheet).getSpill(row, col));
    if (!r || r.blocked) return null;
    return {
      anchorRow: r.range.firstRow,
      anchorCol: r.range.firstCol,
      rows: r.range.lastRow - r.range.firstRow + 1,
      cols: r.range.lastCol - r.range.firstCol + 1,
    };
  }

  /** The rectangle the formula at (`row`, `col`) tried to spill into while it is held at
   *  #SPILL!; null when it is not a blocked anchor, when the attempt runs past the sheet edge,
   *  or when the engine doesn't expose `spillInfo`. */
  blockedSpillRange(sheet: number, row: number, col: number): Range | null {
    assertAlive(this);
    if (!this.capabilities.spillInfo) return null;
    const r = attempt(() => ws(this, sheet).getSpill(row, col));
    if (!r?.blocked) return null;
    return {
      sheet,
      r0: r.range.firstRow,
      c0: r.range.firstCol,
      r1: r.range.lastRow,
      c1: r.range.lastCol,
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
    const arr = attempt(() => ws(this, addr.sheet).getPrecedents(addr.row, addr.col, depth));
    if (!arr) return null;
    return arr.map((n) => ({ sheet: n.sheet, row: n.row, col: n.col }));
  }

  /** Cells whose formulas read from `addr` (1-step dependents by default).
   *  Same depth + cross-sheet semantics as `precedents`. Returns `null`
   *  when the engine doesn't expose `dependents`. */
  dependents(addr: Addr, depth = 1): Addr[] | null {
    assertAlive(this);
    if (!this.capabilities.traceArrows) return null;
    const arr = attempt(() => ws(this, addr.sheet).getDependents(addr.row, addr.col, depth));
    if (!arr) return null;
    return arr.map((n) => ({ sheet: n.sheet, row: n.row, col: n.col }));
  }

  /** Every registered function's canonical name in ascending sort order.
   *  Returns `null` when the engine doesn't expose `functionNames`; the
   *  static `FUNCTION_NAMES` list in `commands/refs.ts` is the fallback
   *  catalog under stub mode. */
  functionNames(): readonly string[] | null {
    assertAlive(this);
    if (!this.capabilities.functionMetadata) return null;
    const names = attempt(() => moduleOf(this).listFunctionNames());
    return names ? [...names] : null;
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
   *  `locale`: 0 = en-US, 1 = ja-JP, 2 = de-DE, 3 = fr-FR. The engine guarantees `minArity` /
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
    return this.functionMetadataForLocale(name, localeTag(locale));
  }

  /** Locale-tagged metadata lookup for all seven UI languages. Provider
   * aliases/descriptions are selected by the normalized BCP-47 tag; native
   * aliases are queried only for the legacy ja/de/fr ordinal locales. */
  functionMetadataForLocale(
    name: string,
    locale: string,
  ): {
    name: string;
    minArity: number;
    maxArity: number | null;
    availability?: number;
    signatureTemplate?: string;
    description?: string;
    localizedName?: string;
  } | null {
    assertAlive(this);
    if (!this.capabilities.functionMetadata) return null;
    const tag = displayLocaleTag(locale);
    const m = attempt(() => moduleOf(this).getFunctionMetadata(name));
    if (!m) return null;
    const provider = internals(this).functionMetadataProvider;
    const canonical = m.name ?? name;
    const entry = provider?.[canonical.toUpperCase()];
    const merged = mergeFunctionMetadata(m, entry, tag);
    // Newer formulon builds carry aliases in the engine's locale table for
    // German and French. Provider aliases still win, but the catalog should
    // be able to display the native alias when no host metadata document is
    // installed. Older builds simply return the canonical name here.
    const providerAlias = entry?.aliases?.[tag];
    const nativeOrdinal = localeOrdinalForTag(tag);
    const engineAlias =
      nativeOrdinal !== null && nativeOrdinal !== 0 && this.capabilities.functionLocale
        ? this.localizeFunctionName(canonical, nativeOrdinal)
        : null;
    const displayName =
      providerAlias !== undefined
        ? providerAlias
        : engineAlias && engineAlias !== canonical
          ? engineAlias
          : merged.localizedName;
    return {
      name: canonical,
      minArity: merged.minArity ?? 0,
      maxArity: merged.maxArity ?? null,
      availability: m.availability,
      ...(merged.signature ? { signatureTemplate: merged.signature } : {}),
      ...(merged.description ? { description: merged.description } : {}),
      ...(displayName ? { localizedName: displayName } : {}),
    };
  }

  /** Canonical → localized function-name lookup. `locale`: 0 = en-US,
   *  1 = ja-JP, 2 = de-DE, 3 = fr-FR. Returns the canonical name unchanged
   *  when no alias is registered for `locale`. Returns the empty string for
   *  an unsuccessful lookup.
   *  Returns `null` when the engine doesn't expose
   *  `localizeFunctionName`. */
  localizeFunctionName(canonicalName: string, locale = 0): string | null {
    assertAlive(this);
    if (!this.capabilities.functionLocale) return null;
    const profile = engineProfile(
      spreadsheetProfileForLocaleOrdinal(this.spreadsheetProfileId(), locale as LocaleOrdinal),
    );
    return attempt(() => moduleOf(this).localizeFunctionName(canonicalName, profile)) ?? '';
  }

  /** Localized → canonical function-name lookup. Falls through to a
   *  case-insensitive match on the canonical name when no alias is
   *  registered. Returns the empty string when the engine reports no
   *  matching function. Returns `null` when the engine doesn't expose
   *  `canonicalizeFunctionName`. */
  canonicalizeFunctionName(localizedName: string, locale = 0): string | null {
    assertAlive(this);
    if (!this.capabilities.functionLocale) return null;
    const profile = engineProfile(
      spreadsheetProfileForLocaleOrdinal(this.spreadsheetProfileId(), locale as LocaleOrdinal),
    );
    return attempt(() => moduleOf(this).canonicalizeFunctionName(localizedName, profile)) ?? '';
  }

  /**
   * Localize a canonical formula for display under the active profile, or an
   * explicitly supplied public profile. This is display-only: callers that
   * persist formulas must continue using the canonical write boundary.
   *
   * Older formulon packages expose only function-name aliases, so the
   * adapter falls back to the older alias-only conversion when the full text
   * surface is absent or when the engine reports/fails a call. That fallback
   * deliberately leaves separators, literals, and quoted text untouched.
   */
  localizeFormulaText(formula: string, profileOverride?: SpreadsheetProfileId | null): string {
    assertAlive(this);
    const profile = profileOverride ?? this.spreadsheetProfileId();
    const locale = localeOrdinalForSpreadsheetProfile(profile);
    const aliasOnly = (): string =>
      localizeFormula(this as unknown as WorkbookHandle, formula, locale);
    if (!this.capabilities.formulaLocale) return aliasOnly();
    if (!profile) return aliasOnly();
    return (
      attempt(() => moduleOf(this).localizeFormula(formula, engineProfile(profile))) ?? aliasOnly()
    );
  }

  /**
   * Convert a formula written with a profile's display syntax back to
   * formulon's canonical en-invariant text. The conversion is intentionally
   * opt-in at this public method; setFormula and evaluation retain the
   * existing alias-only boundary so comma-separated legacy formulas remain
   * unambiguous until the host has explicitly canonicalized full syntax.
   */
  canonicalizeLocalizedFormula(
    formula: string,
    profileOverride?: SpreadsheetProfileId | null,
  ): string {
    assertAlive(this);
    const profile = profileOverride ?? this.spreadsheetProfileId();
    const locale = localeOrdinalForSpreadsheetProfile(profile);
    const aliasOnly = (): string =>
      canonicalizeFormula(this as unknown as WorkbookHandle, formula, locale);
    if (!this.capabilities.formulaLocale) return aliasOnly();
    if (!profile) return aliasOnly();
    return (
      attempt(() => moduleOf(this).canonicalizeFormula(formula, engineProfile(profile))) ??
      aliasOnly()
    );
  }

  /** Return scalar locale facts for the active or explicitly selected profile. */
  localeFacts(profileOverride?: SpreadsheetProfileId | null): LocaleFacts | null {
    assertAlive(this);
    if (!this.capabilities.localeFacts) return null;
    const profile = profileOverride ?? this.spreadsheetProfileId();
    if (!profile) return null;
    return attempt(() => moduleOf(this).getLocaleFacts(engineProfile(profile)));
  }

  /** Workbook calc-mode metadata mirroring `<calcPr calcMode>`. The engine
   *  itself does NOT gate evaluation on this value — every `recalc()` call
   *  honours all dirty cells regardless of mode. This wrapper is what enforces
   *  it: `Manual` suppresses the automatic recalc that follows a cell write,
   *  leaving the cells dirty until Calculate Now. Returns `null` when the
   *  engine doesn't expose `calcMode`. Codes: 0 = Auto, 1 = Manual,
   *  2 = AutoNoTable, under which the engine's own pass skips data tables
   *  and only `recalc()` (Calculate Now) computes them. */
  calcMode(): CalcModeCode | null {
    assertAlive(this);
    if (!this.capabilities.calcMode) return null;
    return attempt(() => wb(this).getCalcMode());
  }

  /** Sets the calc-mode metadata. Returns `false` (no-op) under stub or
   *  older engine package builds. Leaving Manual runs the recalc that the
   *  edits made while in Manual were denied, so switching back to Auto
   *  settles the sheet the way a spreadsheet does. */
  setCalcMode(mode: CalcModeCode): boolean {
    assertAlive(this);
    if (!this.capabilities.calcMode) return false;
    const previous = this.calcMode();
    if (!attemptWrite(() => wb(this).setCalcMode(mode as CalcMode))) return false;
    // These methods are mixed into WorkbookHandle at runtime, so the base
    // class's own members need the cast to be visible here.
    if (previous === CALC_MODE_MANUAL && mode !== CALC_MODE_MANUAL)
      (this as unknown as WorkbookHandle).recalc();
    return true;
  }

  /** The `<calcPr>` / `<workbookPr>` settings outside the calc mode, or
   *  `null` when the engine doesn't expose them. */
  calcSettings(): WorkbookCalcSettings | null {
    assertAlive(this);
    if (!this.capabilities.calcSettings) return null;
    const w = wb(this);
    return {
      fullPrecision: attempt(() => w.getFullPrecision()) ?? true,
      calcOnSave: attempt(() => w.getCalcOnSave()) ?? true,
      r1c1: (attempt(() => w.getRefMode()) ?? RefMode.A1) !== RefMode.A1,
      date1904: attempt(() => w.getDate1904()) ?? false,
    };
  }

  /** Writes one calc setting. `date1904` switches the date system without
   *  converting stored serials, the way the desktop option does. Returns
   *  `false` when unsupported. */
  setCalcSetting<K extends keyof WorkbookCalcSettings>(
    key: K,
    value: WorkbookCalcSettings[K],
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.calcSettings) return false;
    const w = wb(this);
    switch (key) {
      case 'fullPrecision':
        return attemptWrite(() => w.setFullPrecision(value));
      case 'calcOnSave':
        return attemptWrite(() => w.setCalcOnSave(value));
      case 'r1c1':
        return attemptWrite(() => w.setRefMode(value ? RefMode.R1C1 : RefMode.A1));
      case 'date1904':
        return attemptWrite(() => w.setDate1904(value));
    }
    return false;
  }

  /** Formula-behaviour profile selected in the engine. Profiles model host
   *  differences across supported host profiles. Returns
   *  `null` when the engine package does not expose the profile API. */
  spreadsheetProfileId(): SpreadsheetProfileId | null {
    assertAlive(this);
    if (!this.capabilities.spreadsheetProfile) return null;
    const profile = attempt(() => wb(this).getProfileId());
    return profile ? engineProfileToPublic(profile as EngineSpreadsheetProfileId) : null;
  }

  /** Sets the formula-behaviour profile. Returns `false` when unsupported. */
  setSpreadsheetProfileId(profileId: SpreadsheetProfileId): boolean {
    assertAlive(this);
    if (!this.capabilities.spreadsheetProfile) return false;
    return attemptWrite(() => wb(this).setProfileId(engineProfile(profileId)));
  }

  /** Toggle the iterative-formula solver. `maxIterations` and `maxChange`
   *  cap the Gauss-Seidel loop; matches "File → Options → Formulas"
   *  knobs. No-op (returns false) on engines without the iterative surface. */
  setIterative(enabled: boolean, maxIterations: number, maxChange: number): boolean {
    assertAlive(this);
    if (!this.capabilities.iterativeProgress) return false;
    return attemptWrite(() => wb(this).updateIterative({ enabled, maxIterations, maxChange }));
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
    const r = attempt(() => wb(this).getIterative());
    if (!r) return null;
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
    return attemptWrite(() =>
      wb(this).setIterativeProgress(
        callback
          ? (iteration, maxResidual, maxIterations) =>
              callback(iteration, maxResidual, maxIterations) !== false
          : null,
      ),
    );
  }

  /** Renders the lambda value stored at `addr` as spreadsheet formula text. The
   *  returned string never carries a leading `=` — callers prepending it
   *  for the formula-bar edit seed should add the prefix themselves.
   *  Returns `null` when the engine doesn't expose `getLambdaText` or
   *  when the cell is absent / its cached value is not a lambda. */
  getLambdaText(addr: Addr): string | null {
    assertAlive(this);
    if (!this.capabilities.lambdaText) return null;
    const text = attempt(() => ws(this, addr.sheet).getLambdaText(addr.row, addr.col));
    return text ? nameFormulaText(text) || null : null;
  }

  /** Iterate over defined names, hidden ones included. `localSheetId === -1`
   *  means workbook scope; otherwise it is the 0-based sheet index for a
   *  sheet-scoped name. `comment` and `hidden` are omitted when unset. */
  *definedNames(): Generator<{
    name: string;
    formula: string;
    localSheetId: number;
    comment?: string;
    hidden?: boolean;
  }> {
    assertAlive(this);
    for (const e of wb(this).listDefinedNames()) {
      yield {
        name: e.name,
        formula: nameFormulaText(e.formula),
        localSheetId: e.scope ?? -1,
        ...(e.comment ? { comment: e.comment } : {}),
        ...(e.hidden ? { hidden: true } : {}),
      };
    }
  }

  /** Add or replace a workbook/sheet-scoped defined name. Pass an empty `formula`
   *  to remove the name (engine convention). Returns false on engine failure
   *  or when the engine doesn't expose scoped names. */
  setDefinedNameEntry(name: string, formula: string, localSheetId = -1): boolean {
    assertAlive(this);
    if (!this.capabilities.definedNameMutate) return false;
    const scope = localSheetId < 0 ? null : localSheetId;
    const w = wb(this);
    if (formula.length === 0) return attemptWrite(() => w.removeDefinedName(name, scope));
    // The engine replaces every attribute; a formula edit keeps the note and flag.
    const existing = attempt(() => w.getDefinedName(name, scope));
    return attemptWrite(() =>
      w.setDefinedName({
        name,
        formula,
        scope,
        comment: existing?.comment ?? null,
        hidden: existing?.hidden ?? false,
      }),
    );
  }

  /** Set the comment and/or hidden flag of an existing defined name; an
   *  omitted key is left unchanged. Returns false when the name is missing or
   *  the engine lacks the API. */
  setDefinedNameAttrs(
    name: string,
    localSheetId: number,
    attrs: { comment?: string; hidden?: boolean },
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.definedNameAttrs) return false;
    const scope = localSheetId < 0 ? null : localSheetId;
    const w = wb(this);
    const existing = attempt(() => w.getDefinedName(name, scope));
    if (!existing) return false;
    return attemptWrite(() =>
      w.setDefinedName({
        name: existing.name,
        formula: existing.formula,
        scope,
        comment: attrs.comment !== undefined ? attrs.comment || null : existing.comment,
        hidden: attrs.hidden ?? existing.hidden,
      }),
    );
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
