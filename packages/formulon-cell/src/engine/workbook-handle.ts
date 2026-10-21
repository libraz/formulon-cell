import type { SaveDiagnosticsResult } from '@libraz/formulon';
import type { History } from '../commands/history.js';
import { builtInLocaleFor } from '../i18n/locales.js';
import type { SpreadsheetStore } from '../store/store.js';
import { addrKey } from './address.js';
import { detectCapabilities } from './capabilities.js';
import { type ExternalLinkKind, externalLinkKindLabel } from './external-links.js';
import {
  emptyLoadDiagnostics,
  emptySaveDiagnostics,
  type LoadDiagnostics,
  loadDiagnosticsFromEngine,
  type SaveDiagnostics,
  saveDiagnosticsFromEngine,
} from './file-diagnostics.js';
import {
  canonicalizeFormulaForEngine,
  localeOrdinalForSpreadsheetProfile,
  localizeFormula,
} from './function-locale.js';
import type { LoadOptions } from './loader.js';
import { isUsingStub, loadFormulon } from './loader.js';
import { numberValue } from './result.js';
import { engineProfileIdFor, type SpreadsheetPlatform } from './spreadsheet-profile.js';
import { StoreMirror } from './store-mirror.js';
import type {
  Addr,
  CellValue,
  EngineCapabilities,
  FontRecord,
  FormulonModule,
  FunctionMetadataProvider,
  Range,
  Workbook,
} from './types.js';
import { fromEngineValue } from './value.js';
import {
  installAnnotationsMethods,
  type WorkbookHandleAnnotationsMethods,
} from './workbook-handle-annotations.js';
import {
  installChartsMethods,
  type WorkbookHandleChartsMethods,
} from './workbook-handle-charts.js';
import {
  installConditionalFormatMethods,
  type WorkbookHandleConditionalFormatMethods,
} from './workbook-handle-conditional-format.js';
import {
  installFormulasMethods,
  type WorkbookHandleFormulasMethods,
} from './workbook-handle-formulas.js';
import {
  installLayoutMethods,
  type WorkbookHandleLayoutMethods,
} from './workbook-handle-layout.js';
import { installPivotMethods } from './workbook-handle-pivot.js';
import { installPrintMethods, type WorkbookHandlePrintMethods } from './workbook-handle-print.js';
import {
  installProtectionMethods,
  type WorkbookHandleProtectionMethods,
} from './workbook-handle-protection.js';
import {
  installStylesMethods,
  type WorkbookHandleStylesMethods,
} from './workbook-handle-styles.js';
import {
  installTablesMethods,
  type WorkbookHandleTablesMethods,
} from './workbook-handle-tables.js';
import { installValidationMethods } from './workbook-handle-validation.js';

export type ChangeListener = (e: ChangeEvent) => void;

interface AtomicValueBatchMeta {
  readonly id: number;
  readonly index: number;
  readonly size: number;
  readonly formula: string | null;
}

export type ChangeEvent =
  | { kind: 'value'; addr: Addr; next: CellValue; atomicBatch?: AtomicValueBatchMeta }
  | { kind: 'recalc'; dirty: ReadonlySet<string> }
  | { kind: 'sheet-add'; index: number; name: string }
  | { kind: 'sheet-rename'; index: number; name: string }
  | { kind: 'sheet-remove'; index: number }
  | { kind: 'sheet-move'; from: number; to: number };

/** Snapshot of one cell's full state — enough to restore it on undo. */
export interface CellSnapshot {
  addr: Addr;
  value: CellValue;
  formula: string | null;
}

/** One literal/formula write accepted by the atomic adapter boundary. A
 * formula of `null` explicitly removes an existing formula. */
export interface CellPatch {
  readonly addr: Addr;
  readonly value: CellValue;
  readonly formula?: string | null;
}

export interface CellPatchAtomicResult {
  readonly before: readonly CellSnapshot[];
  readonly after: readonly CellSnapshot[];
  readonly changed: readonly Addr[];
}

const UNDO_LIMIT = 100;

/** `WorkbookFormat` ordinals accepted by the engine's `saveWithDiagnostics`. */
const SAVE_FORMAT_ORDINAL = { xlsx: 1, xlsb: 2 } as const;

const sameCellValue = (a: CellValue, b: CellValue): boolean => {
  if (a.kind !== b.kind) return false;
  switch (a.kind) {
    case 'blank':
      return true;
    case 'number':
      return b.kind === 'number' && Object.is(a.value, b.value);
    case 'bool':
      return b.kind === 'bool' && a.value === b.value;
    case 'text':
      return b.kind === 'text' && a.value === b.value;
    case 'error':
      return b.kind === 'error' && a.code === b.code && a.text === b.text;
  }
};

const sameCellSnapshot = (a: CellSnapshot, b: CellSnapshot): boolean =>
  a.addr.sheet === b.addr.sheet &&
  a.addr.row === b.addr.row &&
  a.addr.col === b.addr.col &&
  a.formula === b.formula &&
  sameCellValue(a.value, b.value);

/** `<calcPr calcMode>` code for Manual. */
const CALC_MODE_MANUAL = 1;

/** Mac Excel's 12 pt new-workbook family per language; others use Aptos Narrow. */
const MAC_DEFAULT_FONT_FAMILY: Readonly<Record<string, string>> = {
  ja: '游ゴシック',
  zh: '等线',
  ko: '맑은 고딕',
  th: 'Tahoma',
};

/** Excel's Japanese UI creates new workbooks with Yu Gothic while the
 * non-Japanese baseline remains Calibri 11 for compatibility with existing
 * formulon workbooks. On `'mac'` it is Mac Excel's 12 pt family for the
 * language. */
export const defaultFontForLocale = (
  locale: string | undefined,
  platform: SpreadsheetPlatform = 'default',
): Pick<FontRecord, 'name' | 'size'> => {
  if (platform === 'mac') {
    const language = locale === undefined ? null : builtInLocaleFor(locale);
    return { name: (language && MAC_DEFAULT_FONT_FAMILY[language]) || 'Aptos Narrow', size: 12 };
  }
  return {
    name: locale?.toLowerCase().startsWith('ja') ? '游ゴシック' : 'Calibri',
    size: 11,
  };
};

/** Default row height Windows Excel writes for a new workbook's Normal font,
 *  keyed by `name size`. Fonts outside the table leave the engine default. */
const WINDOWS_DEFAULT_ROW_HEIGHT_PT: ReadonlyMap<string, number> = new Map([
  ['Calibri 11', 15],
  ['游ゴシック 11', 18.75],
]);

/**
 * Boundary between the WASM Workbook and the rest of the UI. Nothing else
 * in the codebase touches the raw engine. Keeps the dispose contract honest.
 */
export class WorkbookHandle {
  readonly capabilities: EngineCapabilities;

  /** Platform whose column-width model converts between store px and the
   *  engine's raw widths. Set by `createDefault` and by the mount, whose
   *  `ui.platform` wins over the value a host-created handle carries. */
  geometryPlatform: SpreadsheetPlatform = 'default';

  /** Windows default row height `createDefault` stated, repeated on added sheets. */
  private newSheetRowHeight: number | undefined;

  private readonly module: FormulonModule;

  private readonly wb: Workbook;

  private readonly listeners = new Set<ChangeListener>();

  private disposed = false;
  private recoveryFailure: Error | null = null;
  private loadLosses: LoadDiagnostics = emptyLoadDiagnostics();
  private saveLosses: SaveDiagnostics = emptySaveDiagnostics();
  private nextAtomicBatchId = 1;

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

  private readonly storeMirror = new StoreMirror(this);

  /** Locale default font for a workbook this handle created on an engine that
   *  cannot state font 0. Null otherwise, since font 0 is then the baseline. */
  private localeDefaultFont: Pick<FontRecord, 'name' | 'size'> | null = null;

  /** Host-injected localized function documentation, merged over the
   *  engine's structural `functionMetadata()` result. `null` until a host
   *  calls `setFunctionMetadataProvider`. */
  // biome-ignore lint/correctness/noUnusedPrivateClassMembers: read/written via the internals() cast in workbook-handle-formulas.ts
  private functionMetadataProvider: FunctionMetadataProvider | null = null;

  /** Depth of the enclosing `withBatchedRecalc` scopes. While above zero,
   *  writes only mark the batch dirty; one recalc runs on exit. */
  private recalcBatchDepth = 0;

  private pendingRecalc = false;

  /** Cells written since the last recalc pass — the `recalc` event's payload. */
  private dirtySinceRecalc = new Set<string>();

  private constructor(module: FormulonModule, wb: Workbook) {
    this.module = module;
    this.wb = wb;
    this.capabilities = detectCapabilities(wb);
  }

  /** Create a new workbook with the defaults `platform`'s desktop spreadsheet
   *  gives one for `locale`. A Mac book takes the engine's measured Mac
   *  profile (font, row height, base column width); a default-profile book
   *  states the locale font and the row height Windows writes for it. */
  static async createDefault(
    opts: LoadOptions & { locale?: string; platform?: SpreadsheetPlatform } = {},
  ): Promise<WorkbookHandle> {
    const { locale, platform = 'default', ...loadOptions } = opts;
    const module = await loadFormulon(loadOptions);
    if (platform === 'mac') {
      const profiled = module.Workbook.createForProfile(engineProfileIdFor(locale ?? 'en', 'mac'));
      if (profiled.isValid()) {
        const handle = new WorkbookHandle(module, profiled);
        handle.geometryPlatform = 'mac';
        return handle;
      }
      profiled.delete();
    }
    const wb = module.Workbook.createDefault();
    const handle = new WorkbookHandle(module, wb);
    handle.geometryPlatform = platform;
    const font = defaultFontForLocale(locale, platform);
    // Stating font 0 is what makes an untouched cell save with the locale font;
    // an engine that cannot restate its default falls back to naming the font
    // on every cell that carries a format, leaving unformatted cells behind.
    if (!handle.setWorkbookDefaultFont(font)) handle.localeDefaultFont = font;
    const rowHeight =
      platform === 'default'
        ? WINDOWS_DEFAULT_ROW_HEIGHT_PT.get(`${font.name} ${font.size}`)
        : undefined;
    if (rowHeight !== undefined) {
      handle.newSheetRowHeight = rowHeight;
      handle.stateNewSheetRowHeight(0);
    }
    return handle;
  }

  static async loadBytes(bytes: Uint8Array, opts: LoadOptions = {}): Promise<WorkbookHandle> {
    const module = await loadFormulon(opts);
    const wb = module.Workbook.loadBytes(bytes);
    if (!wb.isValid()) {
      const msg = module.lastErrorMessage();
      wb.delete();
      throw new Error(`formulon loadBytes failed: ${msg}`);
    }
    const handle = new WorkbookHandle(module, wb);
    if (typeof wb.readDiagnostics === 'function') {
      handle.loadLosses = loadDiagnosticsFromEngine(wb.readDiagnostics());
    }
    return handle;
  }

  /** True when the JS fallback stub is providing the engine surface. */
  get isStub(): boolean {
    return isUsingStub(this.module);
  }

  get version(): string {
    return this.module.versionString();
  }

  /**
   * Font a cell falls back to when its format states no family or size — the
   * baseline every authored font record is built from, and the one stripped
   * back out when engine records are hydrated into the store.
   *
   * Font 0 is that record, for a loaded workbook and for a new one alike —
   * `createDefault` states the locale font there. `localeDefaultFont` only
   * holds a value on an engine that cannot restate its default, where the
   * locale font has to be named on each cell instead.
   */
  get workbookDefaultFont(): Pick<FontRecord, 'name' | 'size' | 'scheme'> | null {
    this.assertAlive();
    return this.localeDefaultFont ?? this.getFontRecord(0);
  }

  /**
   * Font used for display when a cell has no explicit family or size.  This is
   * deliberately resolved through the current workbook theme, while
   * `workbookDefaultFont` remains the raw record used for writeback so a
   * theme-linked `scheme` is never lost.
   *
   * The returned size is still in native points. Callers that paint or measure
   * cells must convert it to CSS pixels at that surface boundary.
   */
  get workbookDisplayFont(): Pick<FontRecord, 'name' | 'size'> | null {
    this.assertAlive();
    if (this.localeDefaultFont) return { ...this.localeDefaultFont };
    const resolved = this.getResolvedFontRecord(0);
    if (resolved) return { name: resolved.name, size: resolved.size };
    const fallback = this.workbookDefaultFont;
    return fallback ? { name: fallback.name, size: fallback.size } : null;
  }

  get sheetCount(): number {
    this.assertAlive();
    return numberValue(this.wb.sheetCount(), 'sheetCount');
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
    const idx = numberValue(this.wb.sheetCount(), 'sheetCount') - 1;
    this.stateNewSheetRowHeight(idx);
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

  /** The engine gives sheets added to a Mac-profile book their profile
   *  defaults; the Windows row height `createDefault` states is ours to repeat. */
  private stateNewSheetRowHeight(sheet: number): void {
    if (this.newSheetRowHeight === undefined || !this.canStateSheetFormatDefaults()) return;
    this.wb.setSheetFormatDefaults(sheet, {
      defaultRowHeight: this.newSheetRowHeight,
      hasDefaultRowHeight: true,
    });
  }

  private canStateSheetFormatDefaults(): boolean {
    return (
      this.capabilities.geometry === true && typeof this.wb.setSheetFormatDefaults === 'function'
    );
  }

  private uniqueSheetName(): string {
    const existing = new Set<string>();
    const n = numberValue(this.wb.sheetCount(), 'sheetCount');
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

  /** Attach the owning spreadsheet store so changes to formats, AutoFilter and
   * page-setup state are immediately mirrored to the engine before an `.xlsx`
   * can be saved. */
  attachStore(store: SpreadsheetStore | null): void {
    this.storeMirror.attach(store);
  }

  /** Detach only the expected owner. A stale asynchronous mount must not
   *  unsubscribe a newer spreadsheet using the same external workbook. */
  detachStore(expectedStore: SpreadsheetStore): void {
    this.storeMirror.detach(expectedStore);
  }

  /** Runs hydration without turning imported filter XML or print settings
   * into a UI-authored replacement. The store subscription still records the
   * new baseline. */
  withEngineSyncMuted<T>(fn: () => T): T {
    return this.storeMirror.withMuted(fn);
  }

  setNumber(a: Addr, value: number): void {
    this.assertAlive();
    this.withJournal(a, () => {
      const s = this.wb.setNumber(a.sheet, a.row, a.col, value);
      if (!s.ok) throw new Error(`setNumber: ${s.message}`);
      this.scheduleRecalc(a);
      this.emit({ kind: 'value', addr: a, next: { kind: 'number', value } });
    });
  }

  setText(a: Addr, value: string): void {
    this.assertAlive();
    this.withJournal(a, () => {
      const s = this.wb.setText(a.sheet, a.row, a.col, value);
      if (!s.ok) throw new Error(`setText: ${s.message}`);
      this.scheduleRecalc(a);
      this.emit({ kind: 'value', addr: a, next: { kind: 'text', value } });
    });
  }

  setBool(a: Addr, value: boolean): void {
    this.assertAlive();
    this.withJournal(a, () => {
      const s = this.wb.setBool(a.sheet, a.row, a.col, value);
      if (!s.ok) throw new Error(`setBool: ${s.message}`);
      this.scheduleRecalc(a);
      this.emit({ kind: 'value', addr: a, next: { kind: 'bool', value } });
    });
  }

  setError(a: Addr, errorCode: number): void {
    this.assertAlive();
    this.withJournal(a, () => {
      const s = this.wb.setError(a.sheet, a.row, a.col, errorCode);
      if (!s.ok) throw new Error(`setError: ${s.message}`);
      this.scheduleRecalc(a);
      this.emit({ kind: 'value', addr: a, next: this.getValue(a) });
    });
  }

  setBlank(a: Addr): void {
    this.assertAlive();
    this.withJournal(a, () => {
      const s = this.wb.setBlank(a.sheet, a.row, a.col);
      if (!s.ok) throw new Error(`setBlank: ${s.message}`);
      this.scheduleRecalc(a);
      this.emit({ kind: 'value', addr: a, next: { kind: 'blank' } });
    });
  }

  setFormula(a: Addr, formula: string): void {
    this.assertAlive();
    this.withJournal(a, () => {
      const s = this.wb.setFormula(a.sheet, a.row, a.col, this.formulaForEngine(formula));
      if (!s.ok) throw new Error(`setFormula: ${s.message}`);
      this.scheduleRecalc(a);
      this.emit({ kind: 'value', addr: a, next: this.getValue(a) });
    });
  }

  /**
   * Render canonical stored formula text with the active profile's display
   * grammar for formula-bar and inline-editor display. This does not change
   * `cellFormula`, save output, or any formula used for engine references.
   */
  getFormulaForDisplay(formula: string): string {
    this.assertAlive();
    // formulon 0.13 can localize the complete formula grammar (function
    // names, separators, decimal literals, booleans, and error spellings).
    // Use that surface when present; the alias-only helper below remains the
    // compatibility path for older engine packages.
    if (this.capabilities.formulaLocale) return this.localizeFormulaText(formula);
    return localizeFormula(
      this,
      formula,
      localeOrdinalForSpreadsheetProfile(this.spreadsheetProfileId()),
    );
  }

  /** Apply a set of cell writes as one adapter transaction. The raw engine is
   * written without the ordinary per-cell journal/event path, then recalculated
   * once. A setter or recalc failure restores touched cells and repairs
   * dependent caches before reporting rejection. If recovery itself fails,
   * the handle rejects subsequent reads/writes until it is replaced. */
  applyCellPatchAtomic(patches: readonly CellPatch[]): CellPatchAtomicResult {
    this.assertAlive();
    const unique = new Map<string, { patch: CellPatch }>();
    for (const patch of patches) {
      const addr = patch.addr;
      if (
        !Number.isInteger(addr.sheet) ||
        !Number.isInteger(addr.row) ||
        !Number.isInteger(addr.col) ||
        addr.sheet < 0 ||
        addr.row < 0 ||
        addr.col < 0
      ) {
        throw new Error('applyCellPatchAtomic: invalid address');
      }
      const key = addrKey(addr);
      unique.set(key, {
        patch:
          typeof patch.formula === 'string'
            ? { ...patch, formula: this.formulaForEngine(patch.formula) }
            : patch,
      });
    }

    const requested = [...unique.values()];
    const beforeSnapshots = this.captureCellSnapshots(requested.map((entry) => entry.patch.addr));
    const entries = requested.map((entry, index) => ({
      ...entry,
      before: beforeSnapshots[index] as CellSnapshot,
    }));
    const before = entries.map((entry) => entry.before);
    const changedEntries = entries.filter((entry) => {
      const formula = entry.patch.formula ?? null;
      return !sameCellSnapshot(entry.before, {
        addr: entry.patch.addr,
        value: entry.patch.value,
        formula,
      });
    });
    if (changedEntries.length === 0) {
      return { before, after: before, changed: [] };
    }

    const pendingBefore = this.pendingRecalc;
    const dirtyBefore = new Set(this.dirtySinceRecalc);
    const manual = this.isManualCalcMode();
    let after: CellSnapshot[] = [];
    try {
      for (const entry of changedEntries) {
        this.writeRawCell(entry.patch.addr, entry.patch.value, entry.patch.formula ?? null);
        this.dirtySinceRecalc.add(addrKey(entry.patch.addr));
      }
      if (!manual) {
        const status = this.wb.recalc();
        if (!status.ok) throw new Error(`recalc: ${status.message}`);
        this.pendingRecalc = false;
      } else {
        // Match ordinary setters in manual calculation mode: engine dirty flags
        // remain pending, but no automatic recalc event is emitted.
        this.pendingRecalc = false;
      }
      after = this.captureCellSnapshots(entries.map((entry) => entry.patch.addr));
    } catch (error) {
      // Restore in reverse order. Direct engine calls intentionally bypass the
      // journal and event paths, so rollback cannot create observable events.
      let rollbackFailure: unknown = null;
      for (let i = changedEntries.length - 1; i >= 0; i -= 1) {
        const entry = changedEntries[i];
        if (!entry) continue;
        try {
          this.writeRawCell(entry.patch.addr, entry.before.value, entry.before.formula);
        } catch (restoreError) {
          rollbackFailure ??= restoreError;
        }
      }
      this.pendingRecalc = pendingBefore;
      this.dirtySinceRecalc = dirtyBefore;
      if (!manual) {
        try {
          // A failed C++ recalc may have changed dependent caches before
          // returning an error. Restore the inputs first, then force one
          // repair pass before reporting the original transaction failure.
          const repair = this.wb.recalc();
          if (!repair.ok) throw new Error(`recalc repair: ${repair.message}`);
        } catch (repairError) {
          rollbackFailure ??= repairError;
        }
      }
      if (rollbackFailure) {
        const detail =
          rollbackFailure instanceof Error ? rollbackFailure.message : 'engine rollback failed';
        this.recoveryFailure = new Error(`applyCellPatchAtomic rollback failed: ${detail}`);
        throw this.recoveryFailure;
      }
      throw error;
    }

    const changed = changedEntries.map((entry) => entry.patch.addr);
    const changedSet = new Set(changedEntries);
    const atomicBatchId = this.nextAtomicBatchId;
    this.nextAtomicBatchId += 1;
    const atomicBatchSize = changedEntries.length;
    let atomicBatchIndex = 0;
    for (let i = 0; i < entries.length; i += 1) {
      const entry = entries[i];
      const snapshot = after[i];
      if (!entry || !snapshot || !changedSet.has(entry)) continue;
      this.emit({
        kind: 'value',
        addr: entry.patch.addr,
        next: { ...snapshot.value },
        atomicBatch: {
          id: atomicBatchId,
          index: atomicBatchIndex,
          size: atomicBatchSize,
          formula: snapshot.formula,
        },
      });
      atomicBatchIndex += 1;
    }
    if (!manual) this.emitRecalc();
    return { before, after, changed };
  }

  /** Coalesce the recalcs of a multi-cell write into a single pass. Bulk
   *  writers (paste, fill, sort, structure edits) wrap their loop in this so
   *  an N-cell write costs one recalc instead of N. Scopes nest; the recalc
   *  runs when the outermost one exits, including on a thrown error so a
   *  partially-applied write is never left stale. An explicit `recalc()`
   *  inside the scope still runs immediately and supersedes what is pending. */
  withBatchedRecalc<T>(fn: () => T): T {
    this.recalcBatchDepth += 1;
    try {
      return fn();
    } finally {
      this.recalcBatchDepth -= 1;
      if (this.recalcBatchDepth === 0 && this.pendingRecalc && !this.disposed) this.recalcAuto();
    }
  }

  /** Recompute after a write at `a`, or leave it to the enclosing batch.
   *  Always a full pass: the engine only evaluates cells it has marked dirty,
   *  so this is already incremental, and unlike `partialRecalc` it can't leave
   *  an off-screen dependent holding a stale value. */
  private scheduleRecalc(a: Addr): void {
    this.dirtySinceRecalc.add(addrKey(a));
    if (this.recalcBatchDepth > 0) {
      this.pendingRecalc = true;
      return;
    }
    this.recalcAuto();
  }

  /** True while the workbook sits in Manual calc mode, where edits accumulate
   *  as engine-side dirty cells until someone asks for a recalc. */
  private isManualCalcMode(): boolean {
    return this.calcMode() === CALC_MODE_MANUAL;
  }

  /** Recalc triggered by an edit rather than by the user. Skipped in Manual
   *  calc mode; `recalc()` stays unconditional so Calculate Now (F9) works
   *  from any mode. */
  recalcAuto(): void {
    this.assertAlive();
    if (this.isManualCalcMode()) {
      // The engine keeps the dirty flags, so the next recalc still catches up.
      this.pendingRecalc = false;
      return;
    }
    this.recalc();
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
    // A full pass covers whatever a batch had queued up.
    this.pendingRecalc = false;
    const s = this.wb.recalc();
    if (!s.ok) throw new Error(`recalc: ${s.message}`);
    this.emitRecalc();
  }

  /** Recompute only formulas whose dependency closure intersects the given
   *  rectangle. Returns the number of cells the engine actually evaluated, or
   *  `null` when the engine doesn't expose `partialRecalc`. Falls back to a
   *  full `recalc()` on engines without the capability so callers can use
   *  this as a drop-in optimization.
   *
   *  Opt-in only: dirty cells outside the rectangle stay dirty and keep their
   *  previous value until a later pass reaches them, so this is for hosts that
   *  knowingly trade freshness off screen for a faster first paint. Cell
   *  writes never take this path. */
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
    this.emitRecalc();
    return r.recomputed;
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
    const n = numberValue(this.wb.cellCount(sheet), `cellCount(${sheet})`);
    for (let i = 0; i < n; i += 1) {
      const e = this.wb.cellAt(sheet, i);
      if (!e.status.ok || e.row === undefined || e.col === undefined || e.value === undefined)
        continue;
      yield {
        addr: { sheet, row: e.row, col: e.col },
        value: fromEngineValue(e.value),
        formula: e.formula ?? null,
      };
    }
  }

  /** Read populated cells within a range, including evaluated spill cells and pivot overlays. */
  *cellsInRange(
    range: Range,
    pageSize = 65536,
  ): Generator<{ addr: Addr; value: CellValue; formula: string | null }> {
    this.assertAlive();
    if (!Number.isInteger(pageSize) || pageSize < 1 || pageSize > 65536) {
      throw new RangeError('cellsInRange pageSize must be between 1 and 65536');
    }
    if (!this.capabilities.rangeEnumeration) {
      for (const entry of this.cells(range.sheet)) {
        if (
          entry.addr.row >= range.r0 &&
          entry.addr.row <= range.r1 &&
          entry.addr.col >= range.c0 &&
          entry.addr.col <= range.c1
        )
          yield entry;
      }
      return;
    }
    let cursor: number | null = null;
    const cursors = new Set<number>();
    do {
      const result = this.wb.getCellsInRange(
        range.sheet,
        {
          firstRow: range.r0,
          firstCol: range.c0,
          lastRow: range.r1,
          lastCol: range.c1,
        },
        cursor,
        pageSize,
      );
      if (!result.status.ok) throw new Error(`getCellsInRange: ${result.status.message}`);
      for (const entry of result.cells) {
        yield {
          addr: { sheet: range.sheet, row: entry.row, col: entry.col },
          value: fromEngineValue(entry.value),
          formula: entry.formula
            ? entry.formula.startsWith('=')
              ? entry.formula
              : `=${entry.formula}`
            : null,
        };
      }
      cursor = result.nextCursor;
      if (cursor !== null) {
        if (cursors.has(cursor)) throw new Error('getCellsInRange returned a repeated cursor');
        cursors.add(cursor);
      }
    } while (cursor !== null);
    for (const entry of this.pivotCells(range.sheet)) {
      if (
        entry.addr.row >= range.r0 &&
        entry.addr.row <= range.r1 &&
        entry.addr.col >= range.c0 &&
        entry.addr.col <= range.c1
      )
        yield entry;
    }
  }

  /** Stored formula in relative R1C1 notation, including the leading equals sign. */
  cellFormulaR1C1(a: Addr): string | null {
    this.assertAlive();
    if (!this.capabilities.formulaR1C1) return null;
    const result = this.wb.getFormulaR1C1(a.sheet, a.row, a.col);
    if (!result.status.ok) return null;
    return result.formula
      ? result.formula.startsWith('=')
        ? result.formula
        : `=${result.formula}`
      : null;
  }

  cellFormula(a: Addr): string | null {
    this.assertAlive();
    if (this.capabilities.formulaRead) {
      const result = this.wb.getFormula(a.sheet, a.row, a.col);
      if (!result.status.ok) throw new Error(`getFormula: ${result.status.message}`);
      return result.formula;
    }
    const n = numberValue(this.wb.cellCount(a.sheet), `cellCount(${a.sheet})`);
    for (let i = 0; i < n; i += 1) {
      const e = this.wb.cellAt(a.sheet, i);
      if (e.status.ok && e.row === a.row && e.col === a.col) return e.formula ?? null;
    }
    return null;
  }

  /** Read formulas for a batch of physical cells with one engine enumeration
   * per touched sheet. Missing cells and physical literals are represented by
   * null, while the returned map keeps the first occurrence order of the
   * requested addresses. */
  cellFormulas(addrs: readonly Addr[]): ReadonlyMap<string, string | null> {
    this.assertAlive();
    return this.readFormulaMap(addrs, false);
  }

  /** Strict formula enumeration used only by the atomic snapshot boundary.
   * Unlike the public best-effort reader, a failed physical entry cannot be
   * mistaken for a literal cell while a transaction is in flight. */
  private readAtomicFormulas(addrs: readonly Addr[]): ReadonlyMap<string, string | null> {
    return this.readFormulaMap(addrs, true);
  }

  private readFormulaMap(
    addrs: readonly Addr[],
    strict: boolean,
  ): ReadonlyMap<string, string | null> {
    const requested = new Map<string, Addr>();
    for (const a of addrs) requested.set(addrKey(a), a);
    const formulas = new Map<string, string | null>();
    if (requested.size === 0) return formulas;

    const bySheet = new Map<number, Map<string, string>>();
    for (const [key, a] of requested) {
      formulas.set(key, null);
      let sheet = bySheet.get(a.sheet);
      if (!sheet) {
        sheet = new Map();
        bySheet.set(a.sheet, sheet);
      }
      sheet.set(`${a.row}:${a.col}`, key);
    }

    for (const [sheet, targets] of bySheet) {
      let n: number;
      try {
        n = numberValue(this.wb.cellCount(sheet), `cellCount(${sheet})`);
      } catch (error) {
        if (!strict) throw error;
        const message = error instanceof Error ? error.message : 'cell count read failed';
        throw new Error(`atomic formula read at sheet:${sheet}: ${message}`);
      }
      for (let i = 0; i < n; i += 1) {
        const e = this.wb.cellAt(sheet, i);
        if (!e.status.ok) {
          if (strict) {
            throw new Error(`atomic formula read at ${sheet}:${i}: ${e.status.message}`);
          }
          continue;
        }
        if (e.row === undefined || e.col === undefined || (strict && e.formula === undefined)) {
          if (strict) {
            throw new Error(`atomic formula read at ${sheet}:${i}: malformed cell entry`);
          }
          continue;
        }
        const coordinateKey = `${e.row}:${e.col}`;
        const key = targets.get(coordinateKey);
        if (key === undefined) continue;
        formulas.set(key, e.formula ?? null);
        targets.delete(coordinateKey);
        if (targets.size === 0) break;
      }
    }
    return formulas;
  }

  save(): Uint8Array {
    return this.saveWithDiagnostics().bytes;
  }

  /** Loss counters captured when the file was read. All zero for workbooks
   *  not loaded from bytes and on engines without `readDiagnostics`. */
  loadDiagnostics(): LoadDiagnostics {
    return { ...this.loadLosses };
  }

  /** Loss counters of the most recent save; all zero before the first one. */
  lastSaveDiagnostics(): SaveDiagnostics {
    return { ...this.saveLosses };
  }

  /** Serialize the workbook and report what the writer dropped or deferred.
   *  Engines without `saveWithDiagnostics` save plainly and report zeros. */
  saveWithDiagnostics(format: 'xlsx' | 'xlsb' = 'xlsx'): {
    bytes: Uint8Array;
    diagnostics: SaveDiagnostics;
  } {
    this.assertAlive();
    const raw =
      typeof this.wb.saveWithDiagnostics === 'function'
        ? this.wb.saveWithDiagnostics(SAVE_FORMAT_ORDINAL[format])
        : this.wb.save();
    if (!raw.status.ok || !raw.bytes) throw new Error(`save: ${raw.status.message}`);
    this.saveLosses = saveDiagnosticsFromEngine(raw as Partial<SaveDiagnosticsResult>);
    return { bytes: raw.bytes, diagnostics: { ...this.saveLosses } };
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
    if (!arr.status.ok) return [];
    return arr.map((r) => ({
      index: r.index,
      relId: r.relId,
      partPath: r.partPath,
      target: r.target,
      targetExternal: r.targetExternal,
      kind: externalLinkKindLabel(r.kind),
    }));
  }

  /** Snapshot of OOXML "passthrough" parts (charts, drawings, pivots, etc.)
   *  preserved verbatim by the engine. Surfaced as a badge so users know
   *  these objects exist even though the UI doesn't render them. */
  getPassthroughs(): { path: string }[] {
    this.assertAlive();
    if (!this.wb.passthroughCount) return [];
    const count = this.wb.passthroughCount();
    if (!count.status.ok) return [];
    const n = count.value;
    const out: { path: string }[] = [];
    for (let i = 0; i < n; i += 1) {
      const e = this.wb.passthroughAt(i);
      if (!e.status.ok || !e.path) continue;
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
    this.storeMirror.attach(null);
    this.disposed = true;
    this.listeners.clear();
    this.wb.delete();
  }

  private emit(e: ChangeEvent): void {
    // Events are emitted only after the engine commit. Listener failures must
    // therefore stay isolated from the committed workbook and from siblings.
    for (const fn of [...this.listeners]) {
      try {
        fn(e);
      } catch {
        // Host/render observers are advisory and cannot roll back an engine.
      }
    }
  }

  /** Announce a completed recalc pass and reset the dirty accumulator. */
  private emitRecalc(): void {
    const dirty = this.dirtySinceRecalc;
    this.dirtySinceRecalc = new Set();
    this.emit({ kind: 'recalc', dirty });
  }

  private assertAlive(): void {
    if (this.disposed) throw new Error('WorkbookHandle is disposed');
    if (this.recoveryFailure) throw this.recoveryFailure;
  }

  private captureSnapshot(a: Addr): CellSnapshot {
    return { addr: a, value: this.getValue(a), formula: this.cellFormula(a) };
  }

  private readAtomicValue(a: Addr): CellValue {
    const r = this.wb.getValue(a.sheet, a.row, a.col);
    if (!r.status.ok) throw new Error(`atomic value read at ${addrKey(a)}: ${r.status.message}`);
    if (r.value === undefined) throw new Error(`atomic value read at ${addrKey(a)}: missing value`);
    return fromEngineValue(r.value);
  }

  private captureCellSnapshots(addrs: readonly Addr[]): CellSnapshot[] {
    const formulas = this.readAtomicFormulas(addrs);
    return addrs.map((a) => ({
      addr: a,
      value: this.readAtomicValue(a),
      formula: formulas.get(addrKey(a)) ?? null,
    }));
  }

  private writeRawCell(a: Addr, value: CellValue, formula: string | null): void {
    const status =
      formula !== null
        ? this.wb.setFormula(a.sheet, a.row, a.col, formula)
        : value.kind === 'number'
          ? this.wb.setNumber(a.sheet, a.row, a.col, value.value)
          : value.kind === 'bool'
            ? this.wb.setBool(a.sheet, a.row, a.col, value.value)
            : value.kind === 'text'
              ? this.wb.setText(a.sheet, a.row, a.col, value.value)
              : value.kind === 'error'
                ? this.wb.setError(a.sheet, a.row, a.col, value.code)
                : this.wb.setBlank(a.sheet, a.row, a.col);
    if (!status.ok) throw new Error(`cell patch: ${status.message}`);
  }

  /**
   * Formulon's stored formula grammar is canonical even when the active host
   * profile displays German/French aliases. Keep this normalization at the
   * two engine formula write boundaries (`setFormula` and atomic patches) so
   * typed input, paste, commands, and trusted host changes persist portable
   * formulas.
   */
  private formulaForEngine(formula: string): string {
    return canonicalizeFormulaForEngine(
      this,
      formula,
      localeOrdinalForSpreadsheetProfile(this.spreadsheetProfileId()),
    );
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
        case 'error':
          this.setError(snap.addr, snap.value.code);
          return;
        default:
          this.setBlank(snap.addr);
      }
    } finally {
      this.replaying = false;
    }
  }
}

/** Aggregate of the feature mixins installed on `WorkbookHandle`. */
export interface WorkbookHandleFeatureMethods
  extends WorkbookHandleAnnotationsMethods,
    WorkbookHandleChartsMethods,
    WorkbookHandleConditionalFormatMethods,
    WorkbookHandleFormulasMethods,
    WorkbookHandleLayoutMethods,
    WorkbookHandlePrintMethods,
    WorkbookHandleProtectionMethods,
    WorkbookHandleStylesMethods,
    WorkbookHandleTablesMethods {}

installAnnotationsMethods(WorkbookHandle);
installChartsMethods(WorkbookHandle);
installConditionalFormatMethods(WorkbookHandle);
installFormulasMethods(WorkbookHandle);
installLayoutMethods(WorkbookHandle);
installPivotMethods(WorkbookHandle);
installPrintMethods(WorkbookHandle);
installProtectionMethods(WorkbookHandle);
installStylesMethods(WorkbookHandle);
installTablesMethods(WorkbookHandle);
installValidationMethods(WorkbookHandle);

export { addrKey };
