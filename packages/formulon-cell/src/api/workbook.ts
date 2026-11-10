import {
  type AddSheetOptions,
  type CellAddress,
  type CellRect,
  type CellValueInput,
  type DisplayText,
  type Sheet as EngineSheet,
  type Workbook as EngineWorkbook,
  FormulonError,
  type LocaleFacts,
  StatusCode,
} from '@libraz/formulon';
import type { InteractionOrigin } from '../commands/interaction-policy.js';
import { WORKBOOK_SAVE_FORMATS, type WorkbookSaveFormat } from '../commands/workbook-save.js';
import {
  emptyLoadDiagnostics,
  type FileLossEntry,
  fileLossEntries,
} from '../engine/file-diagnostics.js';
import { isUsingStub, type LoadOptions } from '../engine/loader.js';
import type { SpreadsheetPlatform } from '../engine/spreadsheet-profile.js';
import type { Addr, FunctionMetadataProvider } from '../engine/types.js';
import {
  type CellSnapshot,
  type FacadeChange,
  type FacadeChangeKind,
  type ChangeEvent as HandleChangeEvent,
  WorkbookHandle,
} from '../engine/workbook-handle.js';
import { type Disposer, Emitter } from './events.js';
import type { MemberSpec, SheetContext, WorkbookContext } from './facade-context.js';
import { Lifecycle } from './lifecycle.js';
import { SHEET_CELL_MEMBERS } from './sheet-cells.js';
import { SHEET_STRUCTURE_MEMBERS, WORKBOOK_NAME_MEMBERS } from './sheet-structure.js';

/** File container a workbook saves to. */
export type SaveFormat = WorkbookSaveFormat;

/** File extension and MIME type of each {@link SaveFormat}, for delivering saved bytes. */
export const SAVE_FORMATS = WORKBOOK_SAVE_FORMATS;

/** One kind of content the engine dropped or deferred while reading or writing a file. */
export type FileLoss = Readonly<FileLossEntry>;

/** What `Workbook.save` produced. */
export interface SaveResult {
  readonly bytes: Uint8Array<ArrayBuffer>;
  readonly format: SaveFormat;
  /** What the writer dropped or deferred; empty for a lossless save. */
  readonly losses: readonly FileLoss[];
}

/** Options of `Workbook.save`. */
export interface SaveOptions {
  /** Container to write. Default: the format the workbook was loaded from (`xlsx` for a new one). */
  readonly format?: SaveFormat;
}

/** Options of `Workbook.createDefault` and `Workbook.load`. */
export interface WorkbookOptions {
  /** Engine to run on. `'stub'` is a small in-memory engine for tests and demos. Default `'wasm'`. */
  readonly engine?: 'wasm' | 'stub';
  /** Locale whose new-workbook defaults `createDefault` applies (font, row height). A loaded file
   *  keeps its own. Default `'en'`. */
  readonly locale?: string;
  /** Desktop platform whose new-workbook profile `createDefault` uses; a profile that cannot load
   *  rejects. Default `'default'` (Windows). */
  readonly platform?: SpreadsheetPlatform;
  /** Resolves the URL of the engine's `.wasm` file for hosts that serve it elsewhere. */
  readonly locateFile?: LoadOptions['locateFile'];
  /** Localized function documentation merged over the engine's function metadata. */
  readonly functionMetadata?: FunctionMetadataProvider;
  /** Cancels the call: an aborted signal rejects with its reason. */
  readonly signal?: AbortSignal;
}

/** A sheet given by facade, engine position or name. */
export type SheetRef = Sheet | number | string;

/** Where a change came from. */
export type ChangeOrigin = InteractionOrigin | `host:${string}`;

/** One changed cell. */
export interface ChangedCell {
  readonly sheet: number;
  readonly at: CellAddress;
}

/** What `on('change')` reports: the cells and the other parts of the workbook one operation changed. */
export interface WorkbookChange {
  readonly origin: ChangeOrigin;
  readonly cells: readonly ChangedCell[];
  /** Non-cell parts touched: `'sheets'` for the sheet list, else the kind of content. */
  readonly parts: readonly (FacadeChangeKind | 'sheets')[];
}

/** Events of a {@link Workbook}. */
export interface WorkbookEvents {
  change: WorkbookChange;
}

interface BatchState {
  depth: number;
  /** Before-state of each cell written in the current batch, by `sheet:row:col`. */
  cells: Map<string, CellSnapshot>;
}

interface WorkbookState {
  readonly handle: WorkbookHandle;
  readonly owned: boolean;
  readonly life: Lifecycle;
  readonly events: Emitter<WorkbookEvents>;
  readonly sheets: WeakMap<EngineSheet, Sheet>;
  readonly batch: BatchState;
  readonly pending: { cells: Map<string, ChangedCell>; parts: Set<FacadeChangeKind | 'sheets'> };
  sourceFormat: SaveFormat;
  origin: ChangeOrigin;
  flushQueued: boolean;
}

const workbookStates = new WeakMap<Workbook, WorkbookState>();
const sheetEngines = new WeakMap<
  Sheet,
  { readonly workbook: Workbook; readonly engine: EngineSheet }
>();
const workbooksByHandle = new WeakMap<WorkbookHandle, Workbook>();
const CONSTRUCT = Symbol('construct');

const stateOf = (wb: Workbook): WorkbookState => {
  const state = workbookStates.get(wb);
  if (!state) throw new TypeError('not a Workbook');
  state.life.assertLive();
  return state;
};

const cellKey = (a: Addr): string => `${a.sheet}:${a.row}:${a.col}`;

/** The `.xlsb` container names its workbook part `xl/workbook.bin`; `.xlsx` uses `xl/workbook.xml`. */
const XLSB_WORKBOOK_PART = new TextEncoder().encode('xl/workbook.bin');

const detectSourceFormat = (bytes: Uint8Array): SaveFormat => {
  const needle = XLSB_WORKBOOK_PART;
  outer: for (let i = 0; i + needle.length <= bytes.length; i += 1) {
    for (let j = 0; j < needle.length; j += 1) if (bytes[i + j] !== needle[j]) continue outer;
    return 'xlsb';
  }
  return 'xlsx';
};

const initFailure = (error: unknown): never => {
  if (error instanceof FormulonError) throw error;
  const message = error instanceof Error ? error.message : String(error);
  const wrapped = new FormulonError(StatusCode.InternalError, message);
  Object.defineProperty(wrapped, 'cause', { value: error, writable: true, configurable: true });
  throw wrapped;
};

const loadOptionsOf = (options: WorkbookOptions): LoadOptions => ({
  preferStub: options.engine === 'stub',
  ...(options.locateFile ? { locateFile: options.locateFile } : {}),
});

/**
 * A workbook: the sheets, names, styles and pivot caches of one file, with the engine's
 * operations. Create one with {@link Workbook.createDefault} or {@link Workbook.load}, and release
 * it with `dispose()` (or a `using` declaration).
 *
 * Members follow the engine's conventions: a failure throws `FormulonError` with a `StatusCode`,
 * a lookup by natural key returns `null` when nothing matches, and a handle lookup throws
 * `NotFound`. Writes made inside {@link Workbook.batch} recalculate once and report one `change`.
 */
export class Workbook {
  /** @internal Use {@link Workbook.createDefault} or {@link Workbook.load}. */
  constructor(token: typeof CONSTRUCT, handle: WorkbookHandle, owned: boolean) {
    if (token !== CONSTRUCT) throw new TypeError('use Workbook.createDefault() or Workbook.load()');
    const state: WorkbookState = {
      handle,
      owned,
      life: new Lifecycle('Workbook'),
      events: new Emitter<WorkbookEvents>(),
      sheets: new WeakMap(),
      batch: { depth: 0, cells: new Map() },
      pending: { cells: new Map(), parts: new Set() },
      sourceFormat: 'xlsx',
      origin: 'instanceApi',
      flushQueued: false,
    };
    workbookStates.set(this, state);
    workbooksByHandle.set(handle, this);
    const offCells = handle.subscribe((e) => this.#record(state, e));
    const offParts = handle.subscribeFacadeChanges((c) => this.#recordPart(state, c));
    state.life.onDispose(() => {
      offCells();
      offParts();
      state.events.clear();
      workbooksByHandle.delete(handle);
      if (owned) handle.dispose();
    });
  }

  /** Creates a new workbook with one sheet and the defaults of `options.locale` and `options.platform`. */
  static async createDefault(options: WorkbookOptions = {}): Promise<Workbook> {
    options.signal?.throwIfAborted();
    const handle = await WorkbookHandle.createDefault({
      ...loadOptionsOf(options),
      locale: options.locale ?? 'en',
      platform: options.platform ?? 'default',
      strictProfile: true,
    }).then(undefined, initFailure);
    return Workbook.#adopt(handle, options, 'xlsx');
  }

  /** Reads a workbook from `.xlsx` or `.xlsb` bytes. */
  static async load(bytes: Uint8Array, options: WorkbookOptions = {}): Promise<Workbook> {
    options.signal?.throwIfAborted();
    const handle = await WorkbookHandle.loadBytes(bytes, loadOptionsOf(options)).then(
      undefined,
      initFailure,
    );
    if (options.platform) handle.geometryPlatform = options.platform;
    return Workbook.#adopt(handle, options, detectSourceFormat(bytes));
  }

  static #adopt(handle: WorkbookHandle, options: WorkbookOptions, format: SaveFormat): Workbook {
    if (options.signal?.aborted) {
      handle.dispose();
      throw options.signal.reason;
    }
    if (options.functionMetadata) handle.setFunctionMetadataProvider(options.functionMetadata);
    const wb = new Workbook(CONSTRUCT, handle, true);
    (workbookStates.get(wb) as WorkbookState).sourceFormat = format;
    return wb;
  }

  /** The engine this workbook runs on. */
  get engine(): 'wasm' | 'stub' {
    return isUsingStub(stateOf(this).handle.engineModule()) ? 'stub' : 'wasm';
  }

  /** True once `dispose()` has run. */
  get disposed(): boolean {
    return (workbookStates.get(this) as WorkbookState).life.disposed;
  }

  /** What the reader dropped or deferred when the file was loaded; empty for a new workbook. */
  get loadLosses(): readonly FileLoss[] {
    return fileLossEntries(stateOf(this).handle.loadDiagnostics());
  }

  /** Releases the workbook. Idempotent; never throws. */
  dispose(): void {
    workbookStates.get(this)?.life.dispose();
  }

  [Symbol.dispose](): void {
    this.dispose();
  }

  /** Subscribes to workbook events; the returned function unsubscribes. */
  on<K extends keyof WorkbookEvents>(
    name: K,
    listener: (event: WorkbookEvents[K]) => void,
  ): Disposer {
    return stateOf(this).events.on(name, listener);
  }

  /**
   * Runs `fn` as one operation: one recalc at the end (none under manual calculation), one `change`
   * event, and one undo step when the workbook is mounted. Cell content writes are atomic: if `fn`
   * throws, every cell it wrote is restored. Other writes made before the throw stay applied and
   * their events still fire; the error is rethrown. Nested batches join the outermost.
   */
  batch<T>(fn: () => T): T {
    return this.#run(stateOf(this), fn);
  }

  /** Recalculates every formula now (Calculate Now). */
  recalc(): void {
    const state = stateOf(this);
    this.#run(state, () => state.handle.recalc());
  }

  /** Writes the workbook. */
  save(options: SaveOptions = {}): SaveResult {
    const state = stateOf(this);
    const format = options.format ?? state.sourceFormat;
    if (format === 'xlsb' && !state.handle.capabilities.xlsbSave) {
      throw new FormulonError(
        StatusCode.NotImplemented,
        'xlsb save unavailable in the stub engine',
      );
    }
    const { bytes, diagnostics } = state.handle.saveWithDiagnostics(format);
    return {
      bytes: bytes as Uint8Array<ArrayBuffer>,
      format,
      losses: fileLossEntries(emptyLoadDiagnostics(), diagnostics),
    };
  }

  /** The sheet `ref` names. An unknown name or a removed sheet throws `NotFound`; a position past
   *  the end, `SheetIndexOutOfRange`. */
  sheet(ref: SheetRef): Sheet {
    const state = stateOf(this);
    return this.#sheetFacade(state, state.handle.engineWorkbook().sheet(engineSheetRef(ref)));
  }

  /** Every sheet, in tab order. */
  listSheets(): Sheet[] {
    const state = stateOf(this);
    return state.handle
      .engineWorkbook()
      .listSheets()
      .map((s) => this.#sheetFacade(state, s));
  }

  /** Appends a sheet; with no name the engine picks the next free `SheetN`. */
  addSheet(name?: string | null, options?: AddSheetOptions): Sheet {
    const state = stateOf(this);
    return this.#run(state, () => {
      const added = state.handle.engineWorkbook().addSheet(name ?? null, options);
      state.handle.noteSheetChange({ kind: 'sheet-add', index: added.index, name: added.name });
      return this.#sheetFacade(state, added);
    });
  }

  /** Removes a sheet. Its facade becomes stale: every member then throws `NotFound`. */
  removeSheet(ref: SheetRef): void {
    const state = stateOf(this);
    this.#run(state, () => {
      const engineWb = state.handle.engineWorkbook();
      const target = engineWb.sheet(engineSheetRef(ref));
      const index = target.index;
      engineWb.removeSheet(target);
      state.sheets.delete(target);
      state.handle.noteSheetChange({ kind: 'sheet-remove', index });
    });
  }

  /** Formats `value` with an Excel number format under the workbook's profile. */
  formatValue(value: CellValueInput, numberFormat: string): DisplayText {
    return stateOf(this).handle.engineWorkbook().formatValue(value, numberFormat);
  }

  /** `formula` (canonical, as stored) in the workbook profile's display language. */
  localizeFormula(formula: string): string {
    const { module, profile } = profileOf(this);
    return module.localizeFormula(formula, profile);
  }

  /** `formula` in the workbook profile's display language, as stored (canonical). */
  canonicalizeFormula(formula: string): string {
    const { module, profile } = profileOf(this);
    return module.canonicalizeFormula(formula, profile);
  }

  /** The workbook profile's localized name of a built-in function. */
  localizeFunctionName(canonicalName: string): string {
    const { module, profile } = profileOf(this);
    return module.localizeFunctionName(canonicalName, profile);
  }

  /** The canonical name of a function given in the workbook profile's language. */
  canonicalizeFunctionName(localizedName: string): string {
    const { module, profile } = profileOf(this);
    return module.canonicalizeFunctionName(localizedName, profile);
  }

  /** Separators, date order and other locale facts of the workbook profile. */
  getLocaleFacts(): LocaleFacts {
    const { module, profile } = profileOf(this);
    return module.getLocaleFacts(profile);
  }

  #sheetFacade(state: WorkbookState, engine: EngineSheet): Sheet {
    let facade = state.sheets.get(engine);
    if (!facade) {
      facade = new Sheet(CONSTRUCT, this, engine);
      state.sheets.set(engine, facade);
    }
    return facade;
  }

  #run<T>(state: WorkbookState, fn: () => T): T {
    const batch = state.batch;
    if (batch.depth > 0) return fn();
    batch.depth = 1;
    batch.cells = new Map();
    const history = state.handle.attachedHistory();
    try {
      const body = (): T => {
        let ok = false;
        try {
          const result = fn();
          ok = true;
          return result;
        } finally {
          if (ok) this.#commitCells(state);
          else state.handle.restoreCells([...batch.cells.values()]);
        }
      };
      return state.handle.withBatchedRecalc(() => (history ? history.transact(body) : body()));
    } finally {
      batch.depth = 0;
      batch.cells = new Map();
      // A failed batch still reports what stayed applied.
      this.#flush(state);
    }
  }

  /** Records the batch's cell writes as one undo entry. */
  #commitCells(state: WorkbookState): void {
    const history = state.handle.attachedHistory();
    const before = [...state.batch.cells.values()];
    if (!history || before.length === 0) return;
    const after = before.map((snap) => state.handle.captureCell(snap.addr));
    history.push({
      undo: () => state.handle.restoreCells(before),
      redo: () => state.handle.restoreCells(after),
    });
  }

  #record(state: WorkbookState, e: HandleChangeEvent): void {
    if (e.kind === 'recalc') return;
    if (e.kind === 'value') {
      state.pending.cells.set(cellKey(e.addr), {
        sheet: e.addr.sheet,
        at: { row: e.addr.row, col: e.addr.col },
      });
    } else {
      state.pending.parts.add('sheets');
    }
    this.#queueFlush(state);
  }

  #recordPart(state: WorkbookState, change: FacadeChange): void {
    state.pending.parts.add(change.kind);
    this.#queueFlush(state);
  }

  /** A change outside any facade operation (a UI edit) is reported once its task settles. */
  #queueFlush(state: WorkbookState): void {
    if (state.batch.depth > 0 || state.flushQueued) return;
    state.flushQueued = true;
    queueMicrotask(() => {
      state.flushQueued = false;
      if (!state.life.disposed && state.batch.depth === 0) this.#flush(state);
    });
  }

  #flush(state: WorkbookState): void {
    const { cells, parts } = state.pending;
    if (cells.size === 0 && parts.size === 0) return;
    const event: WorkbookChange = {
      origin: state.origin,
      cells: [...cells.values()],
      parts: [...parts],
    };
    cells.clear();
    parts.clear();
    state.events.emit('change', event);
  }
}

const profileOf = (wb: Workbook) => {
  const handle = stateOf(wb).handle;
  return { module: handle.engineModule(), profile: handle.engineWorkbook().getProfileId() };
};

const engineSheetRef = (ref: SheetRef): EngineSheet | number | string => {
  if (ref instanceof Sheet) {
    const entry = sheetEngines.get(ref);
    if (!entry) throw new TypeError('not a Sheet');
    return entry.engine;
  }
  return ref;
};

const workbookContext = (wb: Workbook): WorkbookContext => {
  const state = stateOf(wb);
  const run = <T>(fn: () => T): T => wb.batch(fn);
  return {
    workbook: wb,
    handle: state.handle,
    get engine() {
      return stateOf(wb).handle.engineWorkbook();
    },
    run,
    change: (kind, sheet, fn) =>
      run(() => {
        const result = fn();
        state.handle.noteFacadeChange({ kind, sheet });
        return result;
      }),
    sheet: (engine) => wb.sheet(engine.index),
  };
};

/**
 * One worksheet of a {@link Workbook}. Get it from `Workbook.sheet()`, `listSheets()` or
 * `addSheet()`; the workbook keeps one facade per sheet. Cells are addressed as `(row, col)`
 * (0-based) or as one A1 address. Once the sheet is removed every member throws `NotFound`.
 */
export class Sheet {
  /** @internal */
  constructor(token: typeof CONSTRUCT, workbook: Workbook, engine: EngineSheet) {
    if (token !== CONSTRUCT) throw new TypeError('get a Sheet from Workbook.sheet()');
    sheetEngines.set(this, { workbook, engine });
  }

  /** The workbook this sheet belongs to. */
  get workbook(): Workbook {
    return sheetContext(this).workbook;
  }

  /** The sheet's current position in tab order. */
  get index(): number {
    return sheetContext(this).sheetEngine.index;
  }

  /** The sheet's current name. */
  get name(): string {
    return sheetContext(this).sheetEngine.name;
  }

  /** Renames the sheet; formulas that reference it follow. */
  rename(newName: string): void {
    const ctx = sheetContext(this);
    ctx.run(() => {
      ctx.sheetEngine.rename(newName);
      ctx.handle.noteSheetChange({
        kind: 'sheet-rename',
        index: ctx.sheetEngine.index,
        name: ctx.sheetEngine.name,
      });
    });
  }

  /** Moves the sheet in front of `insertBefore`, or to the end for `null`. */
  move(insertBefore: SheetRef | null): void {
    const ctx = sheetContext(this);
    ctx.run(() => {
      const from = ctx.sheetEngine.index;
      ctx.sheetEngine.move(insertBefore === null ? null : engineSheetRef(insertBefore));
      ctx.handle.noteSheetChange({ kind: 'sheet-move', from, to: ctx.sheetEngine.index });
    });
  }
}

const sheetContext = (sheet: Sheet): SheetContext => {
  const entry = sheetEngines.get(sheet);
  if (!entry) throw new TypeError('not a Sheet');
  const base = workbookContext(entry.workbook);
  const state = stateOf(entry.workbook);
  const engine = entry.engine;
  return {
    ...base,
    get engine() {
      return base.engine;
    },
    sheetEngine: engine,
    cellWrite: (args, write) => {
      const [first, second] = args;
      const at =
        typeof first === 'string'
          ? first
          : ({ row: first as number, col: second as number } satisfies CellAddress);
      const { row, col } = cellOf(base.engine, at);
      const cell = { firstRow: row, firstCol: col, lastRow: row, lastCol: col };
      base.run(() => writeCells(state, engine.index, cell, write));
    },
    rangeWrite: (range, write) => base.run(() => writeCells(state, engine.index, range, write)),
    cellAt: (at) => cellOf(base.engine, at),
    rectOf: (range) =>
      typeof range === 'string' ? base.engine.parseReference(range).range : range,
    sheetChange: (kind, fn) => base.change(kind, engine.index, fn),
  };
};

/** Journals the cells of `range` for the running batch, then writes them through `write`. */
const writeCells = <T>(state: WorkbookState, sheet: number, range: CellRect, write: () => T): T => {
  const addrs: Addr[] = [];
  for (let row = range.firstRow; row <= range.lastRow; row += 1) {
    for (let col = range.firstCol; col <= range.lastCol; col += 1) {
      const addr = { sheet, row, col };
      addrs.push(addr);
      const key = cellKey(addr);
      if (!state.batch.cells.has(key)) state.batch.cells.set(key, state.handle.captureCell(addr));
    }
  }
  let result: T | undefined;
  state.handle.writeCellsUnjournaled(addrs, () => {
    result = write();
  });
  return result as T;
};

/** An A1 address or a `{ row, col }` pair as a cell address. */
const cellOf = (wb: EngineWorkbook, at: string | CellAddress): CellAddress => {
  if (typeof at !== 'string') return { row: at.row, col: at.col };
  const { range } = wb.parseReference(at);
  return { row: range.firstRow, col: range.firstCol };
};

type Method = (this: unknown, ...args: unknown[]) => unknown;

const call = (target: object, name: string, args: readonly unknown[]): unknown =>
  (target as Record<string, Method>)[name]?.apply(target, args as unknown[]);

/** How members reach their engine object from a facade instance. */
interface MemberHost<C> {
  readonly proto: object;
  contextOf(self: unknown): C;
  target(ctx: C): object;
  change<T>(ctx: C, kind: FacadeChangeKind, fn: () => T): T;
}

/** Installs a member file's pass-through and hand-written members on a facade prototype. */
const installMembers = <C>(host: MemberHost<C>, spec: MemberSpec<C>): void => {
  const proto = host.proto as Record<string, Method>;
  for (const name of spec.read ?? []) {
    proto[name] = function (...args) {
      return call(host.target(host.contextOf(this)), name, args);
    };
  }
  for (const name of spec.cellWrite ?? []) {
    proto[name] = function (...args) {
      const ctx = host.contextOf(this) as SheetContext;
      ctx.cellWrite(args, () => call(ctx.sheetEngine, name, args));
    };
  }
  for (const [kind, names] of Object.entries(spec.change ?? {})) {
    for (const name of names ?? []) {
      proto[name] = function (...args) {
        const ctx = host.contextOf(this);
        return host.change(ctx, kind as FacadeChangeKind, () => call(host.target(ctx), name, args));
      };
    }
  }
  for (const [name, method] of Object.entries(spec.methods ?? {})) {
    proto[name] = function (...args) {
      return (method as (ctx: C, ...rest: unknown[]) => unknown)(host.contextOf(this), ...args);
    };
  }
};

const SHEET_HOST: MemberHost<SheetContext> = {
  proto: Sheet.prototype,
  contextOf: (self) => sheetContext(self as Sheet),
  target: (ctx) => ctx.sheetEngine,
  change: (ctx, kind, fn) => ctx.sheetChange(kind, fn),
};

const WORKBOOK_HOST: MemberHost<WorkbookContext> = {
  proto: Workbook.prototype,
  contextOf: (self) => workbookContext(self as Workbook),
  target: (ctx) => ctx.engine,
  change: (ctx, kind, fn) => ctx.change(kind, null, fn),
};

installMembers(SHEET_HOST, SHEET_CELL_MEMBERS);
installMembers(SHEET_HOST, SHEET_STRUCTURE_MEMBERS);
installMembers(WORKBOOK_HOST, WORKBOOK_NAME_MEMBERS);

/** @internal The engine binding behind a facade workbook. */
export const workbookHandleOf = (wb: Workbook): WorkbookHandle => stateOf(wb).handle;

/** @internal The facade of an engine binding the caller already holds; `owned` disposes the binding
 *  with the facade. Returns the existing facade when there is one. */
export const workbookOf = (handle: WorkbookHandle, owned = false): Workbook =>
  workbooksByHandle.get(handle) ?? new Workbook(CONSTRUCT, handle, owned);

/** @internal Runs `fn` with `origin` as the origin of the `change` it causes. */
export const withChangeOrigin = <T>(wb: Workbook, origin: ChangeOrigin, fn: () => T): T => {
  const state = stateOf(wb);
  const previous = state.origin;
  state.origin = origin;
  try {
    return fn();
  } finally {
    state.origin = previous;
  }
};
