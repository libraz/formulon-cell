import {
  type AddSheetOptions,
  type CellAddress,
  type CellRect,
  type CellValueInput,
  type DisplayText,
  type Chart as EngineChart,
  type Drawing as EngineDrawing,
  type Picture as EnginePicture,
  type Pivot as EnginePivot,
  type PivotCache as EnginePivotCache,
  type Shape as EngineShape,
  type Sheet as EngineSheet,
  type Slicer as EngineSlicer,
  type Table as EngineTable,
  type Workbook as EngineWorkbook,
  FormulonError,
  type LocaleFacts,
  StatusCode,
} from '@libraz/formulon';
import { type HistoryEntry, recordingHistory, transactIn } from '../commands/history.js';
import type { InteractionOrigin } from '../commands/interaction-policy.js';
import { resetHistoryAfterSheetRemoval } from '../commands/sheet-mutate.js';
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
import { CHART_MEMBERS, Chart, SHEET_CHART_MEMBERS } from './chart.js';
import {
  DRAWING_MEMBERS,
  Drawing,
  PICTURE_MEMBERS,
  Picture,
  SHAPE_MEMBERS,
  SHEET_DRAWING_MEMBERS,
  Shape,
  SLICER_MEMBERS,
  Slicer,
} from './drawing.js';
import { type Disposer, Emitter } from './events.js';
import type {
  FacadeClass,
  HandleContext,
  MemberSpec,
  SheetContext,
  WorkbookContext,
} from './facade-context.js';
import {
  type JournalCodec,
  type JournalScope,
  type JournalSpec,
  type Restore,
  resolveJournal,
} from './journal/index.js';
import {
  captureSheetName,
  isIndexShifting,
  sheetAddedEntry,
  sheetMovedEntry,
} from './journal/structure.js';
import { Lifecycle } from './lifecycle.js';
import { SHEET_OOXML_MEMBERS, SheetOoxml, WORKBOOK_OOXML_MEMBERS, WorkbookOoxml } from './ooxml.js';
import { PIVOT_MEMBERS, Pivot, SHEET_PIVOT_MEMBERS } from './pivot.js';
import { PIVOT_CACHE_MEMBERS, PivotCache, WORKBOOK_PIVOT_CACHE_MEMBERS } from './pivot-cache.js';
import { SHEET_CELL_MEMBERS } from './sheet-cells.js';
import { SHEET_FIND_MEMBERS } from './sheet-find.js';
import { SHEET_FORMAT_MEMBERS, WORKBOOK_FORMAT_MEMBERS } from './sheet-format.js';
import { SHEET_STRUCTURE_MEMBERS, WORKBOOK_NAME_MEMBERS } from './sheet-structure.js';
import { SHEET_TABLE_MEMBERS, TABLE_MEMBERS, Table, WORKBOOK_TABLE_MEMBERS } from './table.js';
import { WORKBOOK_CALC_MEMBERS } from './workbook-calc.js';
import { WORKBOOK_PEOPLE_MEMBERS } from './workbook-people.js';

// Each module augments Sheet/Workbook; a type re-export keeps its declarations in the published
// types (the declaration build drops bare imports). This module is not an entry point.
export type * from './chart.js';
export type * from './drawing.js';
export type * from './pivot.js';
export type * from './pivot-cache.js';
export type * from './sheet-cells.js';
export type * from './sheet-find.js';
export type * from './sheet-format.js';
export type * from './sheet-structure.js';
export type * from './table.js';
export type * from './workbook-calc.js';
export type * from './workbook-people.js';

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
   *  keeps its own defaults; `load` opens it under this locale's profile when given, so a
   *  manually calculated book keeps exactly the stale values the file carries. Default `'en'`. */
  readonly locale?: string;
  /** Desktop platform whose new-workbook profile `createDefault` uses; a profile that cannot load
   *  rejects. `load` opens the file under this platform's profile when given. Pass the same
   *  locale and platform the spreadsheet mounts with. Default `'default'` (Windows). */
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

/** One non-cell write's slice of engine state, captured on first touch in a segment. */
interface JournalRecord {
  /** The non-cell change undo and redo announce; null for a record whose cells announce it. */
  readonly change: FacadeChange | null;
  readonly before: Restore;
  /** Captures the slice again at seal time. */
  readonly capture: () => Restore | null;
}

/** The writes between two seals; sealed, it becomes one history entry. */
interface Segment {
  /** Before-state of each cell written, by `sheet:row:col`. */
  readonly cells: Map<string, CellSnapshot>;
  /** Before-state of each journaled slice, by codec id and scope. */
  readonly parts: Map<string, JournalRecord>;
}

interface BatchState {
  depth: number;
  /** The segment being written. */
  segment: Segment;
  /** Entries of the segments already sealed in this batch. */
  sealed: HistoryEntry[];
  /** A write that cannot be journaled faithfully ran: History is cleared when the batch ends. */
  reset: boolean;
}

const newSegment = (): Segment => ({ cells: new Map(), parts: new Map() });

interface WorkbookState {
  readonly handle: WorkbookHandle;
  readonly owned: boolean;
  readonly life: Lifecycle;
  readonly events: Emitter<WorkbookEvents>;
  readonly sheets: WeakMap<EngineSheet, Sheet>;
  readonly batch: BatchState;
  readonly pending: {
    cells: Map<string, ChangedCell>;
    parts: Set<FacadeChangeKind | 'sheets'>;
    /** Origin in force when the first queued change was recorded. */
    origin: ChangeOrigin;
  };
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
/** The workbook or sheet behind each `ooxml` facade, and each owner's facade. */
const ooxmlOwners = new WeakMap<object, Workbook | Sheet>();
const ooxmlFacades = new WeakMap<Workbook | Sheet, object>();

const ooxmlOf = <T extends object>(owner: Workbook | Sheet, proto: T): T => {
  let facade = ooxmlFacades.get(owner);
  if (!facade) {
    facade = Object.create(proto) as object;
    ooxmlOwners.set(facade, owner);
    ooxmlFacades.set(owner, facade);
  }
  return facade as T;
};

const ooxmlOwner = (facade: unknown): Workbook | Sheet => {
  const owner = ooxmlOwners.get(facade as object);
  if (!owner) throw new TypeError('not an ooxml facade');
  return owner;
};

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
 *
 * InteractionPolicy and sheet protection govern end-user interaction; the Workbook/Sheet API is
 * host code and is not gated by them.
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
      batch: { depth: 0, segment: newSegment(), sealed: [], reset: false },
      pending: { cells: new Map(), parts: new Set(), origin: 'instanceApi' },
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
    }).then(undefined, initFailure);
    return Workbook.#adopt(handle, options, 'xlsx');
  }

  /** Reads a workbook from `.xlsx` or `.xlsb` bytes. */
  static async load(bytes: Uint8Array, options: WorkbookOptions = {}): Promise<Workbook> {
    options.signal?.throwIfAborted();
    const handle = await WorkbookHandle.loadBytes(bytes, {
      ...loadOptionsOf(options),
      ...(options.locale !== undefined ? { locale: options.locale } : {}),
      ...(options.platform !== undefined ? { platform: options.platform } : {}),
    }).then(undefined, initFailure);
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

  /** The version of the calc engine the workbook runs on, as the engine reports it; the stub engine reports `'stub'`. */
  get engineVersion(): string {
    return stateOf(this).handle.version;
  }

  /** True once `dispose()` has run. */
  get disposed(): boolean {
    return (workbookStates.get(this) as WorkbookState).life.disposed;
  }

  /** The OOXML-level layer: the style record tables and the parts kept verbatim. */
  get ooxml(): WorkbookOoxml {
    stateOf(this);
    return ooxmlOf(this, WorkbookOoxml.prototype);
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

  /** Subscribes to workbook events; the returned function unsubscribes. A `Workbook` or `Sheet`
   *  setter and `batch()` report `change` synchronously when they return. A write made outside
   *  them (`instance.applyChanges`, an edit in the grid) is reported one microtask later, after the
   *  write's own values and the instance's synchronous `cell` / `change` events; each event
   *  carries the origin of its writes. */
  on<K extends keyof WorkbookEvents>(
    name: K,
    listener: (event: WorkbookEvents[K]) => void,
  ): Disposer {
    return stateOf(this).events.on(name, listener);
  }

  /**
   * Runs `fn` as one operation: one recalc at the end (none under manual calculation), one `change`
   * event, and one undo step when the workbook is mounted. Cell content writes are atomic: if `fn`
   * throws, every cell it wrote is restored, except cells written before a write that shifts
   * addresses (`insertRows`, `deleteRows`, `insertColumns`, `deleteColumns`, `insertCells`,
   * `deleteCells`, `addSheet`, `Sheet.move`, `Sheet.copy`, `Pivot.showDetails`): that write ends the
   * span the rollback covers. Other writes made before the throw stay applied and their events
   * still fire; the error is rethrown. Nested batches join the outermost.
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
    const ctx = workbookContext(this);
    return this.#run(state, () => {
      const added = ctx.sheetStep(
        () => ctx.engine.addSheet(name ?? null, options),
        (sheet) => ({ kind: 'sheet-add', index: sheet.index, name: sheet.name }),
        (at) =>
          ctx.engine.addSheet(
            at.name,
            at.insertBefore !== undefined ? { insertBefore: at.insertBefore } : undefined,
          ),
      );
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
      // Entries recorded before the removal hold indexes of the old sheet list.
      state.batch.reset = true;
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
    batch.segment = newSegment();
    batch.sealed = [];
    batch.reset = false;
    const history = recordingHistory(state.handle.attachedHistory());
    try {
      const body = (): T => {
        let ok = false;
        try {
          const result = fn();
          ok = true;
          return result;
        } finally {
          if (ok) this.#commitJournal(state, history);
          else state.handle.restoreCells([...batch.segment.cells.values()]);
        }
      };
      return state.handle.withBatchedRecalc(() => transactIn(history, body));
    } finally {
      const reset = batch.reset;
      batch.depth = 0;
      batch.segment = newSegment();
      batch.sealed = [];
      batch.reset = false;
      if (reset) resetHistoryAfterSheetRemoval(history);
      // A failed batch still reports what stayed applied.
      this.#flush(state);
    }
  }

  /** Records the batch's writes as one undo step: every sealed segment's entry joins the open frame. */
  #commitJournal(state: WorkbookState, history: ReturnType<typeof recordingHistory>): void {
    if (!history) return;
    seal(state);
    if (state.batch.reset) return;
    for (const entry of state.batch.sealed) history.push(entry);
  }

  /** Pins the queued change to the origin in force now; a different origin flushes the earlier one first. */
  #captureOrigin(state: WorkbookState): void {
    const { cells, parts } = state.pending;
    if (cells.size === 0 && parts.size === 0) {
      state.pending.origin = state.origin;
    } else if (state.pending.origin !== state.origin && state.batch.depth === 0) {
      this.#flush(state);
      state.pending.origin = state.origin;
    }
  }

  #record(state: WorkbookState, e: HandleChangeEvent): void {
    if (e.kind === 'recalc') return;
    this.#captureOrigin(state);
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
    this.#captureOrigin(state);
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
      origin: state.pending.origin,
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

/** The history entry of a sealed segment: undo and redo restore its slices and cells, announcing
 *  each part; `null` when there is nothing to undo or a slice cannot be captured. */
const sealedEntry = (state: WorkbookState, segment: Segment): HistoryEntry | null => {
  const { handle } = state;
  const cellsBefore = [...segment.cells.values()];
  const parts = [...segment.parts.values()];
  if (cellsBefore.length === 0 && parts.length === 0) return null;
  const cellsAfter = cellsBefore.map((snap) => handle.captureCell(snap.addr));
  const afters: Restore[] = [];
  for (const part of parts) {
    const after = part.capture();
    if (!after) {
      state.batch.reset = true;
      return null;
    }
    afters.push(after);
  }
  return {
    undo: () =>
      handle.withBatchedRecalc(() => {
        for (let i = parts.length - 1; i >= 0; i -= 1) {
          const part = parts[i] as JournalRecord;
          part.before();
          if (part.change) handle.noteFacadeChange(part.change);
        }
        handle.restoreCells(cellsBefore);
      }),
    redo: () =>
      handle.withBatchedRecalc(() => {
        handle.restoreCells(cellsAfter);
        parts.forEach((part, i) => {
          (afters[i] as Restore)();
          if (part.change) handle.noteFacadeChange(part.change);
        });
      }),
  };
};

/** Seals the segment being written into a history entry; a segment with nothing to undo adds none. */
const seal = (state: WorkbookState): void => {
  const batch = state.batch;
  const segment = batch.segment;
  batch.segment = newSegment();
  if (!recordingHistory(state.handle.attachedHistory()) || batch.reset) return;
  const entry = sealedEntry(state, segment);
  if (entry) batch.sealed.push(entry);
};

/** Whether the running batch journals into a History. */
const journaling = (state: WorkbookState): boolean =>
  !state.batch.reset && recordingHistory(state.handle.attachedHistory()) !== null;

/** Captures the before-state of the slice a non-cell write touches, when the workbook records into
 *  a History. Returns what to call once the write succeeded, or `null` when nothing is journaled. */
const openRecord = (
  state: WorkbookState,
  target: { readonly sheet: number | null; readonly range?: CellRect },
  change: FacadeChange | null,
  spec: JournalSpec | undefined,
): (() => void) | null => {
  const { handle, batch } = state;
  if (batch.reset || !recordingHistory(handle.attachedHistory())) return null;
  const choice = resolveJournal(change?.kind ?? null, target.sheet, target.range, spec);
  if (choice === 'none') return null;
  if (choice === 'reset') {
    return () => {
      batch.reset = true;
    };
  }
  const member = typeof spec === 'object' ? spec : undefined;
  const scope: JournalScope = {
    handle,
    engine: handle.engineWorkbook(),
    sheet: target.sheet,
    ...(target.range ? { range: target.range } : {}),
    hostId: member?.hostId ?? '',
    name: member?.name ?? '',
    args: member?.args ?? [],
  };
  const key = `${choice.id}:${choice.scopeKey(scope)}`;
  const { parts } = batch.segment;
  if (parts.has(key)) return null;
  const before = choice.capture(scope);
  if (!before) {
    return () => {
      batch.reset = true;
    };
  }
  return () => {
    parts.set(key, { change, before, capture: () => (choice as JournalCodec).capture(scope) });
  };
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
    change: (kind, sheet, fn, range, journal) =>
      run(() => {
        const change: FacadeChange = range ? { kind, sheet, range } : { kind, sheet };
        const shifting = isIndexShifting(journal);
        if (shifting) seal(state);
        const record = openRecord(state, change, change, journal);
        const result = fn();
        record?.();
        if (shifting) seal(state);
        state.handle.noteFacadeChange(change);
        return result;
      }),
    sheetStep: (write, announce, repeat) =>
      run(() => {
        const countBefore = state.handle.engineWorkbook().listSheets().length;
        const record = journaling(state);
        if (record) seal(state);
        const added = write();
        state.handle.noteSheetChange(announce(added));
        if (record) {
          const { index, name } = added;
          const insertBefore = index < countBefore ? index : undefined;
          state.batch.sealed.push(
            sheetAddedEntry(state.handle, index, announce, () => repeat({ name, insertBefore })),
          );
        }
        return added;
      }),
    sheet: (engine) => wb.sheet(engine.index),
    facade: (engine, cls) => facadeOf(wb, engine, cls),
    engineOf: (value) =>
      typeof value === 'object' && value !== null
        ? (handleEngines.get(value)?.engine ?? sheetEngines.get(value as Sheet)?.engine ?? value)
        : value,
  };
};

/** The engine handle behind each object facade, and each handle's facade. */
const handleEngines = new WeakMap<
  object,
  { readonly workbook: Workbook; readonly engine: object }
>();
const handleFacades = new WeakMap<object, object>();

const facadeOf = <F extends object>(wb: Workbook, engine: object, cls: FacadeClass<F>): F => {
  let facade = handleFacades.get(engine);
  if (!facade || !(facade instanceof cls)) {
    facade = Object.create(cls.prototype) as object;
    handleEngines.set(facade, { workbook: wb, engine });
    handleFacades.set(engine, facade);
  }
  return facade as F;
};

const handleContext = <H>(facade: unknown): HandleContext<H> => {
  const entry = handleEngines.get(facade as object);
  if (!entry) throw new TypeError('not a facade object');
  const base = workbookContext(entry.workbook);
  const target = entry.engine as H & { readonly sheet?: EngineSheet };
  return {
    ...base,
    get engine() {
      return base.engine;
    },
    target,
    get sheetIndex() {
      return target.sheet?.index ?? null;
    },
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

  /** The OOXML-level layer: XF indices, the range-XF remap and the sheet's XML fragments. */
  get ooxml(): SheetOoxml {
    sheetContext(this);
    return ooxmlOf(this, SheetOoxml.prototype);
  }

  /** Renames the sheet; formulas that reference it follow. */
  rename(newName: string): void {
    const ctx = sheetContext(this);
    const state = stateOf(ctx.workbook);
    ctx.run(() => {
      const index = ctx.sheetEngine.index;
      const before = journaling(state) ? captureSheetName(ctx.handle, index) : null;
      ctx.sheetEngine.rename(newName);
      ctx.handle.noteSheetChange({ kind: 'sheet-rename', index, name: ctx.sheetEngine.name });
      const { parts } = state.batch.segment;
      const key = `sheet-name:${index}`;
      if (before && !parts.has(key)) {
        parts.set(key, {
          change: null,
          before,
          capture: () => captureSheetName(ctx.handle, index),
        });
      }
    });
  }

  /** Moves the sheet in front of `insertBefore`, or to the end for `null`. */
  move(insertBefore: SheetRef | null): void {
    const ctx = sheetContext(this);
    const state = stateOf(ctx.workbook);
    ctx.run(() => {
      const recording = journaling(state);
      if (recording) seal(state);
      const from = ctx.sheetEngine.index;
      ctx.sheetEngine.move(insertBefore === null ? null : engineSheetRef(insertBefore));
      const to = ctx.sheetEngine.index;
      ctx.handle.noteSheetChange({ kind: 'sheet-move', from, to });
      if (recording) state.batch.sealed.push(sheetMovedEntry(ctx.handle, from, to));
    });
  }

  /** Copies the sheet with its cells, formats and objects, in front of `insertBefore` (the end
   *  when omitted or `null`), named `newName` or `Name (2)`. Formulas in the copy keep pointing
   *  where they pointed. */
  copy(options: SheetCopyOptions = {}): Sheet {
    const ctx = sheetContext(this);
    return ctx.run(() => {
      const source = ctx.sheetEngine.index;
      const { insertBefore, newName } = options;
      const copied = ctx.sheetStep(
        () =>
          ctx.engine.copySheet(ctx.sheetEngine, {
            ...(insertBefore != null ? { insertBefore: engineSheetRef(insertBefore) } : {}),
            ...(newName !== undefined ? { newName } : {}),
          }),
        (sheet) => ({ kind: 'sheet-copy', source, index: sheet.index, name: sheet.name }),
        (at) =>
          ctx.engine.copySheet(ctx.engine.sheet(source), {
            ...(at.insertBefore !== undefined ? { insertBefore: at.insertBefore } : {}),
            newName: at.name,
          }),
      );
      return ctx.sheet(copied);
    });
  }
}

/** Where `Sheet.copy` puts the copy and what it calls it. */
export interface SheetCopyOptions {
  /** The sheet the copy goes in front of; omitted or `null` appends it. */
  readonly insertBefore?: SheetRef | null;
  /** The copy's name; omitted, `Name (2)` and onward. */
  readonly newName?: string;
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
    rangeAnnounce: (range, write, journal) =>
      base.run(() => {
        const record = openRecord(state, { sheet: engine.index, range }, null, journal);
        const result = announceCells(state, rangeAddrs(engine.index, range), write);
        record?.();
        return result;
      }),
    cellAt: (at) => cellOf(base.engine, at),
    rectOf: (range) =>
      typeof range === 'string' ? base.engine.parseReference(range).range : range,
    sheetChange: (kind, fn, range, journal) => base.change(kind, engine.index, fn, range, journal),
  };
};

/** Journals the cells of `range` for the running batch, then writes them through `write`. */
const rangeAddrs = (sheet: number, range: CellRect): Addr[] => {
  const addrs: Addr[] = [];
  for (let row = range.firstRow; row <= range.lastRow; row += 1) {
    for (let col = range.firstCol; col <= range.lastCol; col += 1) addrs.push({ sheet, row, col });
  }
  return addrs;
};

/** Runs one engine write over `addrs`, announcing each cell's value afterwards. */
const announceCells = <T>(state: WorkbookState, addrs: readonly Addr[], write: () => T): T => {
  let result: T | undefined;
  state.handle.writeCellsUnjournaled(addrs, () => {
    result = write();
  });
  return result as T;
};

const writeCells = <T>(state: WorkbookState, sheet: number, range: CellRect, write: () => T): T => {
  const addrs = rangeAddrs(sheet, range);
  const { cells } = state.batch.segment;
  for (const addr of addrs) {
    const key = cellKey(addr);
    if (!cells.has(key)) cells.set(key, state.handle.captureCell(addr));
  }
  return announceCells(state, addrs, write);
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
  /** Names the host in the journal registry. */
  readonly id: string;
  readonly proto: object;
  contextOf(self: unknown): C;
  target(ctx: C): object;
  change<T>(ctx: C, kind: FacadeChangeKind, fn: () => T, journal?: JournalSpec): T;
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
        return host.change(
          ctx,
          kind as FacadeChangeKind,
          () => call(host.target(ctx), name, args),
          {
            hostId: host.id,
            name,
            args,
          },
        );
      };
    }
  }
  for (const [name, method] of Object.entries(spec.methods ?? {})) {
    proto[name] = function (...args) {
      return (method as (ctx: C, ...rest: unknown[]) => unknown)(host.contextOf(this), ...args);
    };
  }
  for (const name of spec.props ?? []) {
    Object.defineProperty(proto, name, {
      configurable: true,
      get(this: unknown) {
        return (host.target(host.contextOf(this)) as Record<string, unknown>)[name];
      },
    });
  }
  for (const [name, read] of Object.entries(spec.getters ?? {})) {
    Object.defineProperty(proto, name, {
      configurable: true,
      get(this: unknown) {
        return (read as (ctx: C) => unknown)(host.contextOf(this));
      },
    });
  }
};

/** The host of a facade over an engine object handle. */
const handleHost = <H>(id: string, proto: object): MemberHost<HandleContext<H>> => ({
  id,
  proto,
  contextOf: (self) => handleContext<H>(self),
  target: (ctx) => ctx.target as object,
  change: (ctx, kind, fn, journal) => ctx.change(kind, ctx.sheetIndex, fn, undefined, journal),
});

const SHEET_HOST: MemberHost<SheetContext> = {
  id: 'sheet',
  proto: Sheet.prototype,
  contextOf: (self) => sheetContext(self as Sheet),
  target: (ctx) => ctx.sheetEngine,
  change: (ctx, kind, fn, journal) => ctx.sheetChange(kind, fn, undefined, journal),
};

const WORKBOOK_HOST: MemberHost<WorkbookContext> = {
  id: 'workbook',
  proto: Workbook.prototype,
  contextOf: (self) => workbookContext(self as Workbook),
  target: (ctx) => ctx.engine,
  change: (ctx, kind, fn, journal) => ctx.change(kind, null, fn, undefined, journal),
};

const WORKBOOK_OOXML_HOST: MemberHost<WorkbookContext> = {
  ...WORKBOOK_HOST,
  id: 'workbook.ooxml',
  proto: WorkbookOoxml.prototype,
  contextOf: (self) => workbookContext(ooxmlOwner(self) as Workbook),
  target: (ctx) => ctx.engine.ooxml,
};

const SHEET_OOXML_HOST: MemberHost<SheetContext> = {
  ...SHEET_HOST,
  id: 'sheet.ooxml',
  proto: SheetOoxml.prototype,
  contextOf: (self) => sheetContext(ooxmlOwner(self) as Sheet),
  target: (ctx) => ctx.sheetEngine.ooxml,
};

installMembers(SHEET_HOST, SHEET_CELL_MEMBERS);
installMembers(SHEET_HOST, SHEET_STRUCTURE_MEMBERS);
installMembers(SHEET_HOST, SHEET_FORMAT_MEMBERS);
installMembers(SHEET_HOST, SHEET_FIND_MEMBERS);
installMembers(WORKBOOK_HOST, WORKBOOK_NAME_MEMBERS);
installMembers(WORKBOOK_HOST, WORKBOOK_FORMAT_MEMBERS);
installMembers(WORKBOOK_HOST, WORKBOOK_CALC_MEMBERS);
installMembers(WORKBOOK_HOST, WORKBOOK_PEOPLE_MEMBERS);
installMembers(WORKBOOK_OOXML_HOST, WORKBOOK_OOXML_MEMBERS);
installMembers(SHEET_OOXML_HOST, SHEET_OOXML_MEMBERS);
installMembers(SHEET_HOST, SHEET_TABLE_MEMBERS);
installMembers(WORKBOOK_HOST, WORKBOOK_TABLE_MEMBERS);
installMembers(handleHost<EngineTable>('table', Table.prototype), TABLE_MEMBERS);
installMembers(SHEET_HOST, SHEET_DRAWING_MEMBERS);
installMembers(SHEET_HOST, SHEET_CHART_MEMBERS);
installMembers(handleHost<EngineDrawing>('drawing', Drawing.prototype), DRAWING_MEMBERS);
installMembers(handleHost<EnginePicture>('picture', Picture.prototype), PICTURE_MEMBERS);
installMembers(handleHost<EngineShape>('shape', Shape.prototype), SHAPE_MEMBERS);
installMembers(handleHost<EngineSlicer>('slicer', Slicer.prototype), SLICER_MEMBERS);
installMembers(handleHost<EngineChart>('chart', Chart.prototype), CHART_MEMBERS);
installMembers(SHEET_HOST, SHEET_PIVOT_MEMBERS);
installMembers(WORKBOOK_HOST, WORKBOOK_PIVOT_CACHE_MEMBERS);
installMembers(handleHost<EnginePivot>('pivot', Pivot.prototype), PIVOT_MEMBERS);
installMembers(
  handleHost<EnginePivotCache>('pivotCache', PivotCache.prototype),
  PIVOT_CACHE_MEMBERS,
);

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
