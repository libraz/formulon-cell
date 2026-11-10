import {
  type CellAddress,
  type CellRect,
  FormulonError,
  StatusCode,
  type Value,
} from '@libraz/formulon';
import { autofitColWidth, autofitRowHeight } from '../../commands/autofit-measurement.js';
import {
  BUILT_IN_COMMAND_OPERATION,
  canExecuteBuiltIn,
} from '../../commands/built-in-command-policy.js';
import type { PasteSpecialOptions } from '../../commands/clipboard/paste-special.js';
import type { History as LegacyHistory } from '../../commands/history.js';
import type {
  HostCellChangeInput,
  InteractionOperation,
  ChangeBatchResult as LegacyBatchResult,
} from '../../commands/interaction-policy.js';
import { sheetPermits } from '../../commands/protection.js';
import { clearTraceArrows, traceDependents, tracePrecedents } from '../../commands/traces.js';
import { resolveCellInput, resolveRangeInput } from '../../engine/range-input.js';
import type { Addr, Range } from '../../engine/types.js';
import { fromEngineValue } from '../../engine/value.js';
import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import type {
  Extension,
  ExtensionHandle,
  ExtensionInput,
  FeatureFlags,
} from '../../extensions/index.js';
import type { DeepPartial, Strings } from '../../i18n/strings.js';
import type { BorderDrawHandle } from '../../interact/border-draw.js';
import type { ClipboardHandle } from '../../interact/clipboard.js';
import type { FormatPainterHandle } from '../../interact/format-painter.js';
import type { ViewportOptions } from '../../interact/navigation-policy.js';
import { shareSpreadsheetInternals, spreadsheetInternals } from '../../mount/internals.js';
import type { ToolbarInstance } from '../../mount/toolbar.js';
import { runRibbonCommand } from '../../mount/toolbar.js';
import type {
  SpreadsheetInstance as LegacyInstance,
  MountOptions,
  ScreenClipResult,
} from '../../mount/types.js';
import { Spreadsheet } from '../../mount.js';
import { mutators } from '../../store/store.js';
import type { TraceArrow } from '../../store/types.js';
import { type BatchRejection, type ChangeBatchResult, changeBatchResult } from '../change-batch.js';
import { spreadsheetFailure } from '../errors.js';
import { type Disposer, Emitter } from '../events.js';
import { Lifecycle } from '../lifecycle.js';
import {
  type ChangeOrigin,
  type SheetRef,
  Workbook,
  withChangeOrigin,
  workbookHandleOf,
} from '../workbook.js';
import { type CommandId, isCommandId } from './command-ids.js';
import { createDialogs, type Dialogs } from './dialogs.js';
import type {
  ApplyChangesOptions,
  CellChange,
  FormulaHint,
  SelectionState,
  SpreadsheetInstanceEvents,
  SpreadsheetOptions,
  SpreadsheetPatch,
  UpdateEvent,
  ViewportConfig,
} from './types.js';

/** The undo history of a mounted spreadsheet; facade writes on its workbook land here too. */
export interface History {
  readonly canUndo: boolean;
  readonly canRedo: boolean;
  undo(): void;
  redo(): void;
  /** Drops every undo and redo step. */
  clear(): void;
  on(name: 'change', listener: () => void): Disposer;
}

/** Selection of the mounted spreadsheet. Refs are A1 text and may be sheet-qualified. */
export interface Selection {
  get(): SelectionState;
  /** Selects `ref`; a sheet-qualified ref switches the active sheet. */
  set(ref: string): void;
  /** Extends the selection from its anchor to `ref`. */
  extend(ref: string): void;
  on(name: 'change', listener: (state: SelectionState) => void): Disposer;
}

/** The visible region. Change it with `update({ viewport })`. */
export interface Viewport {
  /** Scrolls `ref` into view. */
  scrollTo(ref: string): void;
  /** Zoom of the active sheet, in percent. */
  readonly zoomPercent: number;
  readonly options: ViewportConfig;
}

/** The ribbon commands, acting on the current selection. */
export interface Commands {
  /** Runs a command. Throws `policy-denied` or `protected` when it may not run. */
  execute(id: CommandId): void;
  /** Whether `execute(id)` would run now. */
  can(id: CommandId): boolean;
}

/** Formula auditing arrows. */
export interface Audit {
  /** Draws arrows from the cells `at` reads. */
  tracePrecedents(at: string): void;
  /** Draws arrows to the cells that read `at`. */
  traceDependents(at: string): void;
  /** Removes every arrow. */
  clear(): void;
  list(): readonly TraceArrow[];
}

/** Extensions mounted next to the built-in features. */
export interface Extensions {
  /** Mounts an extension; the returned disposer removes it. Adding an id already in use replaces
   *  that extension, a built-in one included. */
  add(extension: Extension): Disposer;
  /** Removes an extension; throws `NotFound` when none has that id. */
  remove(id: string): void;
  /** The handle of an extension, or `null`. */
  get(id: string): ExtensionHandle | null;
  /** Ids of the mounted extensions. */
  list(): string[];
}

/** UI language and strings. Change the language with `update({ locale })`. */
export interface I18n {
  readonly locale: string;
  readonly strings: Strings;
  /** Overrides strings of `locale`; takes effect at once when it is the current locale. */
  override(locale: string, strings: DeepPartial<Strings>): void;
  on(name: 'change', listener: (event: { locale: string; strings: Strings }) => void): Disposer;
}

/** How a cell's displayed text is replaced. */
export interface CellFormatter {
  readonly id: string;
  /** Lower runs first. Defaults to 50. */
  readonly priority?: number;
  readonly match: (cell: {
    sheet: number;
    at: CellAddress;
    value: Value;
    formula: string | null;
  }) => boolean;
  /** The text to show; `null` falls through to the next formatter. */
  readonly format: (cell: {
    sheet: number;
    at: CellAddress;
    value: Value;
    formula: string | null;
  }) => string | null;
}

/** Display formatters for cells. */
export interface CellRegistry {
  add(entry: CellFormatter): Disposer;
  /** Removes a formatter; throws `NotFound` when none has that id. */
  remove(id: string): void;
  get(id: string): CellFormatter | null;
  list(): CellFormatter[];
  on(name: 'change', listener: () => void): Disposer;
}

/** Formula hints for autocomplete and the argument helper. */
export interface FormulaHints {
  add(entry: FormulaHint): Disposer;
  /** Removes a hint; throws `NotFound` when none has that name. */
  remove(name: string): void;
  has(name: string): boolean;
  get(name: string): FormulaHint | null;
  list(): FormulaHint[];
  clear(): void;
  on(name: 'change', listener: () => void): Disposer;
}

interface InstanceState {
  legacy: LegacyInstance;
  workbook: Workbook;
  ownsWorkbook: boolean;
  options: SpreadsheetOptions;
  origin: ChangeOrigin | null;
  readonly life: Lifecycle;
  readonly events: Emitter<SpreadsheetInstanceEvents>;
  readonly selection: Selection;
  readonly history: History;
  readonly commands: Commands;
  readonly dialogs: Dialogs;
  readonly audit: Audit;
  readonly viewport: Viewport;
  readonly extensions: Extensions;
  readonly i18n: I18n;
  readonly cells: CellRegistry;
  readonly formulaHints: FormulaHints;
}

/** Workbooks a live instance shows, by engine binding. */
const mountedWorkbooks = new WeakSet<WorkbookHandle>();
const states = new WeakMap<SpreadsheetInstance, InstanceState>();
const CONSTRUCT = Symbol('construct');

const notFound = (what: string): FormulonError => new FormulonError(StatusCode.NotFound, what);

const stateOf = (instance: SpreadsheetInstance): InstanceState => {
  const state = states.get(instance);
  if (!state) throw new TypeError('not a SpreadsheetInstance');
  state.life.assertLive();
  return state;
};

const toRect = (r: Range): CellRect => ({
  firstRow: r.r0,
  firstCol: r.c0,
  lastRow: r.r1,
  lastCol: r.c1,
});

const legacyViewport = (config: ViewportConfig | null | undefined): ViewportOptions | undefined =>
  config && config.mode === 'fixed'
    ? {
        range: config.range,
        ...(config.tabNavigation ? { tabNavigation: config.tabNavigation } : {}),
      }
    : undefined;

const legacyBatch = (result: LegacyBatchResult): ChangeBatchResult => {
  const at = (a: Addr): { sheet: number; at: CellAddress } => ({
    sheet: a.sheet,
    at: { row: a.row, col: a.col },
  });
  const rejections: BatchRejection[] = result.rejected.map((r) => ({
    sheet: r.addr?.sheet ?? -1,
    at: { row: r.addr?.row ?? -1, col: r.addr?.col ?? -1 },
    code: r.code,
    reason: r.reason ?? r.code,
  }));
  return changeBatchResult(result.applied.map(at), rejections);
};

/** Protect Sheet option a command's operation needs on a protected sheet; `null` when the
 *  operation is checked cell by cell instead, `undefined` when no option allows it. */
const PROTECTION_PERMISSION: Partial<
  Record<InteractionOperation, Parameters<typeof sheetPermits>[2] | null>
> = {
  valueEdit: null,
  formulaEdit: null,
  clear: null,
  paste: null,
  fill: null,
  moveCells: null,
  format: 'formatCells',
  insertRows: 'insertRows',
  deleteRows: 'deleteRows',
  insertColumns: 'insertColumns',
  deleteColumns: 'deleteColumns',
  resizeRows: 'formatRows',
  resizeColumns: 'formatColumns',
  sort: 'sort',
  filter: 'autoFilter',
  hyperlink: 'insertHyperlinks',
  object: 'objects',
  merge: undefined,
  table: undefined,
  comment: undefined,
  validation: undefined,
};

type CommandRefusal = 'policy-denied' | 'protected' | null;

const commandRefusal = (legacy: LegacyInstance, id: string): CommandRefusal => {
  if (!canExecuteBuiltIn(legacy.store, id, 'instanceApi').allowed) return 'policy-denied';
  const operation = BUILT_IN_COMMAND_OPERATION[id];
  if (!operation || !(operation in PROTECTION_PERMISSION)) return null;
  const permission = PROTECTION_PERMISSION[operation];
  if (permission === null) return null;
  const state = legacy.store.getState();
  return sheetPermits(state, state.data.sheetIndex, permission) ? null : 'protected';
};

const selectionOf = (legacy: LegacyInstance): SelectionState => {
  const { selection, data } = legacy.store.getState();
  const ranges = [selection.range, ...(selection.extraRanges ?? [])].map(toRect);
  return {
    sheet: data.sheetIndex,
    active: { row: selection.active.row, col: selection.active.col },
    range: toRect(selection.range),
    ranges,
  };
};

const sameSelection = (a: SelectionState, b: SelectionState): boolean =>
  JSON.stringify(a) === JSON.stringify(b);

/**
 * A spreadsheet mounted on a host element: the grid, its chrome and the objects that drive it.
 * Get one from {@link mount}; release it with `dispose()`. After `dispose()` every other member
 * throws `disposed`.
 */
export class SpreadsheetInstance {
  /** @internal Use {@link mount}. */
  constructor(token: typeof CONSTRUCT, state: InstanceState) {
    if (token !== CONSTRUCT) throw new TypeError('use mount()');
    states.set(this, state);
    shareSpreadsheetInternals(this, state.legacy);
  }

  /** The element the spreadsheet is mounted on. */
  get host(): HTMLElement {
    return stateOf(this).legacy.host;
  }

  /** The workbook shown. */
  get workbook(): Workbook {
    return stateOf(this).workbook;
  }

  /** The current effective options. */
  get options(): Readonly<SpreadsheetOptions> {
    return stateOf(this).options;
  }

  get selection(): Selection {
    return stateOf(this).selection;
  }

  get viewport(): Viewport {
    return stateOf(this).viewport;
  }

  get history(): History {
    return stateOf(this).history;
  }

  get commands(): Commands {
    return stateOf(this).commands;
  }

  get dialogs(): Dialogs {
    return stateOf(this).dialogs;
  }

  get audit(): Audit {
    return stateOf(this).audit;
  }

  get extensions(): Extensions {
    return stateOf(this).extensions;
  }

  get i18n(): I18n {
    return stateOf(this).i18n;
  }

  get cells(): CellRegistry {
    return stateOf(this).cells;
  }

  get formulaHints(): FormulaHints {
    return stateOf(this).formulaHints;
  }

  /** Copy, cut and paste; `null` when the `clipboard` feature is off. */
  get clipboard():
    | (ClipboardHandle & { pasteSpecial(options: PasteSpecialOptions): boolean })
    | null {
    const legacy = stateOf(this).legacy;
    const handle = legacy.clipboard;
    if (!handle) return null;
    return Object.assign(Object.create(handle) as ClipboardHandle, {
      pasteSpecial: (options: PasteSpecialOptions) => legacy.pasteSpecial(options),
    });
  }

  /** Format Painter; `null` when the `formatPainter` feature is off. */
  get formatPainter(): FormatPainterHandle | null {
    return stateOf(this).legacy.formatPainter ?? null;
  }

  /** Draw Borders; `null` when the `borderDraw` feature is off. */
  get borderDraw(): BorderDrawHandle | null {
    return stateOf(this).legacy.borderDraw ?? null;
  }

  /** The ribbon mounted inside the host, or `null`. */
  get toolbar(): ToolbarInstance | null {
    return stateOf(this).legacy.toolbar;
  }

  /** Applies a patch to the options: keys present are replaced, `null` resets one to its default,
   *  and keys left out keep their values. Emits `update` with the keys it changed. */
  update(patch: SpreadsheetPatch): void {
    const state = stateOf(this);
    const keys = (Object.keys(patch) as (keyof SpreadsheetPatch)[]).filter(
      (key) => patch[key] !== undefined,
    );
    if (keys.length === 0) return;
    for (const key of keys) applyPatchKey(state, key, patch[key] ?? null);
    const event: UpdateEvent = { keys };
    state.events.emit('update', event);
  }

  /** Writes host cell changes as one batch. Cells the policy or protection refuses are reported in
   *  `rejections`; nothing throws for them. */
  applyChanges(changes: readonly CellChange[], options: ApplyChangesOptions): ChangeBatchResult {
    const state = stateOf(this);
    const legacy = state.legacy;
    const handle = workbookHandleOf(state.workbook);
    const active = legacy.store.getState().data.sheetIndex;
    const input: HostCellChangeInput[] = changes.map((change, index) => {
      const sheet =
        change.sheet === undefined ? active : state.workbook.sheet(change.sheet as SheetRef).index;
      const addr =
        typeof change.at === 'string'
          ? resolveCellInput(change.at, handle, sheet, `applyChanges[${index}].at`)
          : { sheet, row: change.at.row, col: change.at.col };
      return 'input' in change
        ? { addr, input: change.input }
        : {
            addr,
            value: fromEngineValue(change.value),
            ...(change.formula !== undefined ? { formula: change.formula } : {}),
          };
    });
    const origin: ChangeOrigin = options.origin ?? 'instanceApi';
    const previous = state.origin;
    state.origin = origin;
    try {
      const result = withChangeOrigin(state.workbook, origin, () =>
        legacy.applyChanges(input, {
          history: options.history,
          ...(options.origin ? { origin: options.origin } : {}),
        }),
      );
      return legacyBatch(result);
    } finally {
      state.origin = previous;
    }
  }

  /** Fitted widths of the columns of `range`, in px. */
  measureColumnWidths(range: string): number[] {
    const state = stateOf(this);
    const r = rangeOf(state, range);
    const out: number[] = [];
    for (let col = r.c0; col <= r.c1; col += 1) {
      out.push(autofitColWidth(autofitTarget(state), col, r.r0, r.r1, state.legacy.i18n.locale));
    }
    return out;
  }

  /** Fitted heights of the rows of `range`, in px. */
  measureRowHeights(range: string): number[] {
    const state = stateOf(this);
    const r = rangeOf(state, range);
    const out: number[] = [];
    for (let row = r.r0; row <= r.r1; row += 1) {
      out.push(autofitRowHeight(autofitTarget(state), row, r.c0, r.c1, state.legacy.i18n.locale));
    }
    return out;
  }

  /** Opens the print flow (or the PDF export). */
  print(mode: 'print' | 'pdf' = 'print'): void {
    stateOf(this).legacy.print(mode);
  }

  /** Re-reads the printer list through `host.refreshPrinterProfiles`. */
  refreshPrinterProfiles(): Promise<void> {
    return stateOf(this)
      .legacy.refreshPrinterProfiles()
      .then(() => undefined);
  }

  /** Captures a screen region through `host.captureScreenClip`. */
  captureScreenClip(): Promise<ScreenClipResult | null> {
    return stateOf(this).legacy.captureScreenClip();
  }

  /** Subscribes to instance events; the returned function unsubscribes. */
  on<K extends keyof SpreadsheetInstanceEvents>(
    name: K,
    listener: (event: SpreadsheetInstanceEvents[K]) => void,
  ): Disposer {
    return stateOf(this).events.on(name, listener);
  }

  /** True once `dispose()` has run. */
  get disposed(): boolean {
    return (states.get(this) as InstanceState).life.disposed;
  }

  /** Unmounts the spreadsheet; a workbook the instance created is disposed with it. Idempotent;
   *  never throws. */
  dispose(): void {
    states.get(this)?.life.dispose();
  }

  [Symbol.dispose](): void {
    this.dispose();
  }
}

const rangeOf = (state: InstanceState, ref: string): Range =>
  resolveRangeInput(
    ref,
    workbookHandleOf(state.workbook),
    state.legacy.store.getState().data.sheetIndex,
    'range',
  );

const autofitTarget = (state: InstanceState) => ({
  store: state.legacy.store,
  host: state.legacy.host,
  workbookHandle: workbookHandleOf(state.workbook),
});

const applyPatchKey = (state: InstanceState, key: keyof SpreadsheetPatch, value: unknown): void => {
  const legacy = state.legacy;
  const next = { ...state.options } as Record<string, unknown>;
  if (value === null) delete next[key];
  else next[key] = value;
  switch (key) {
    case 'workbook':
      swapWorkbook(state, (value as Workbook | null) ?? null);
      return;
    case 'ui':
      legacy.setUi((value as SpreadsheetOptions['ui']) ?? undefined);
      break;
    case 'policy':
      legacy.setPolicy((value as SpreadsheetOptions['policy']) ?? undefined);
      break;
    case 'viewport':
      legacy.setViewportOptions(legacyViewport(value as ViewportConfig | null));
      break;
    case 'contextMenu':
      legacy.setContextMenu((value as SpreadsheetOptions['contextMenu']) ?? undefined);
      break;
    case 'overlays':
      legacy.setOverlayOptions((value as SpreadsheetOptions['overlays']) ?? undefined);
      break;
    case 'toolbar':
      legacy.setToolbar(value !== null);
      break;
    case 'theme':
      legacy.setTheme((value as string | null) ?? 'paper');
      break;
    case 'locale':
      if (!legacy.setLocale((value as string | null) ?? 'en')) {
        throw new FormulonError(
          StatusCode.PreconditionFailed,
          `locale ${String(value)} could not be applied`,
        );
      }
      break;
    case 'strings':
      legacy.i18n.extend(legacy.i18n.locale, (value as DeepPartial<Strings> | null) ?? {});
      break;
    case 'extensions':
      legacy.setExtensions(
        ((value as readonly ExtensionInput[] | null) ?? undefined) as ExtensionInput[] | undefined,
      );
      break;
    case 'formulaHints':
      state.formulaHints.clear();
      for (const hint of (value as readonly FormulaHint[] | null) ?? [])
        state.formulaHints.add(hint);
      break;
    case 'host':
      applyHostProviders(legacy, (value as SpreadsheetOptions['host'] | null) ?? {});
      break;
    case 'uploadStatus':
      legacy.setUploadStatus((value as SpreadsheetOptions['uploadStatus'] | null) ?? null);
      break;
    case 'macroRecording':
      legacy.setMacroRecording((value as boolean | null) ?? null);
      break;
  }
  state.options = next as SpreadsheetOptions;
};

const applyHostProviders = (
  legacy: LegacyInstance,
  host: NonNullable<SpreadsheetOptions['host']>,
): void => {
  legacy.setPrinterProfiles(host.printerProfiles);
  legacy.setPrinterProfileId(host.printerProfileId);
};

const swapWorkbook = (state: InstanceState, next: Workbook | null): void => {
  if (next === state.workbook) return;
  const replacement = next;
  if (!replacement) throw new TypeError('update({ workbook: null }): a workbook is required');
  const nextHandle = workbookHandleOf(replacement);
  if (mountedWorkbooks.has(nextHandle)) {
    throw spreadsheetFailure('workbook-in-use', 'the workbook is mounted by another spreadsheet');
  }
  const previous = state.workbook;
  const previousHandle = workbookHandleOf(previous);
  void state.legacy
    .setWorkbook(nextHandle)
    .catch((error: unknown) => state.events.emit('error', { error, source: 'engine' }));
  mountedWorkbooks.delete(previousHandle);
  mountedWorkbooks.add(nextHandle);
  state.workbook = replacement;
  state.options = { ...state.options, workbook: replacement };
  if (state.ownsWorkbook) previous.dispose();
  state.ownsWorkbook = false;
  state.events.emit('workbook', { workbook: replacement });
};

const createHistory = (history: LegacyHistory): History => ({
  get canUndo() {
    return history.canUndo();
  },
  get canRedo() {
    return history.canRedo();
  },
  undo() {
    history.undo();
  },
  redo() {
    history.redo();
  },
  clear() {
    history.clear();
  },
  on: (_name, listener) => history.subscribe(listener),
});

const createSelection = (
  state: () => InstanceState,
  emitter: Emitter<{ change: SelectionState }>,
): Selection => {
  const resolve = (ref: string): Range => {
    const s = state();
    return resolveRangeInput(
      ref,
      workbookHandleOf(s.workbook),
      s.legacy.store.getState().data.sheetIndex,
      'selection',
    );
  };
  const activate = (s: InstanceState, sheet: number): void => {
    if (sheet === s.legacy.store.getState().data.sheetIndex) return;
    if (!spreadsheetInternals(s.legacy).binding.activateSheet(sheet)) {
      throw new FormulonError(StatusCode.PreconditionFailed, `sheet ${sheet} cannot be activated`);
    }
  };
  return {
    get: () => selectionOf(state().legacy),
    set(ref) {
      const s = state();
      const range = resolve(ref);
      activate(s, range.sheet);
      const store = s.legacy.store;
      mutators.setActive(store, { sheet: range.sheet, row: range.r0, col: range.c0 });
      if (range.r1 !== range.r0 || range.c1 !== range.c0) {
        mutators.extendRangeTo(store, { sheet: range.sheet, row: range.r1, col: range.c1 });
      }
    },
    extend(ref) {
      const s = state();
      const range = resolve(ref);
      const sheet = s.legacy.store.getState().data.sheetIndex;
      mutators.extendRangeTo(s.legacy.store, { sheet, row: range.r1, col: range.c1 });
    },
    on: (_name, listener) => emitter.on('change', listener),
  };
};

const createCommands = (legacy: () => LegacyInstance): Commands => ({
  execute(id) {
    if (!isCommandId(id)) throw notFound(`unknown command ${String(id)}`);
    const target = legacy();
    const refusal = commandRefusal(target, id);
    if (refusal)
      throw spreadsheetFailure(refusal, `command ${id} may not run`, { context: { command: id } });
    if (target.toolbar) target.toolbar.applyCommand(id);
    else runRibbonCommand(target, id);
  },
  can(id) {
    return isCommandId(id) && commandRefusal(legacy(), id) === null;
  },
});

const createAudit = (legacy: () => LegacyInstance, at: (ref: string) => Addr): Audit => ({
  tracePrecedents(ref) {
    const target = legacy();
    tracePrecedents(target.store, target.workbook, at(ref), target.history);
  },
  traceDependents(ref) {
    const target = legacy();
    traceDependents(target.store, target.workbook, at(ref), target.history);
  },
  clear() {
    const target = legacy();
    clearTraceArrows(target.store, target.history);
  },
  list: () =>
    legacy()
      .store.getState()
      .traces.items.map((arrow) => ({ ...arrow })),
});

const createExtensions = (legacy: () => LegacyInstance): Extensions => {
  const ids = new Set<string>();
  return {
    add(extension) {
      const target = legacy();
      target.use(extension);
      ids.add(extension.id);
      return () => {
        if (ids.delete(extension.id)) target.remove(extension.id);
      };
    },
    remove(id) {
      if (!legacy().remove(id)) throw notFound(`no extension ${id}`);
      ids.delete(id);
    },
    get: (id) => legacy().features[id] ?? null,
    list: () => Object.keys(legacy().features).filter((id) => legacy().features[id] !== undefined),
  };
};

const createI18n = (legacy: () => LegacyInstance): I18n => ({
  get locale() {
    return legacy().i18n.locale;
  },
  get strings() {
    return legacy().i18n.strings;
  },
  override(locale, strings) {
    legacy().i18n.extend(locale, strings);
  },
  on(_name, listener) {
    const i18n = legacy().i18n;
    return i18n.subscribe((strings) => listener({ locale: i18n.locale, strings }));
  },
});

const cellInput = (
  sheet: number,
  addr: Addr,
  value: Value,
  formula: string | null,
): { sheet: number; at: CellAddress; value: Value; formula: string | null } => ({
  sheet,
  at: { row: addr.row, col: addr.col },
  value,
  formula,
});

const createCellRegistry = (
  legacy: () => LegacyInstance,
  valueAt: (a: Addr) => Value,
): CellRegistry => {
  const entries = new Map<string, CellFormatter>();
  const registry = () => legacy().cells;
  return {
    add(entry) {
      entries.set(entry.id, entry);
      const input = (i: { addr: Addr; formula: string | null }) =>
        cellInput(i.addr.sheet, i.addr, valueAt(i.addr), i.formula);
      const off = registry().registerFormatter({
        id: entry.id,
        ...(entry.priority !== undefined ? { priority: entry.priority } : {}),
        match: (i) => entry.match(input(i)),
        format: (i) => entry.format(input(i)),
      });
      return () => {
        if (entries.get(entry.id) !== entry) return;
        entries.delete(entry.id);
        off();
      };
    },
    remove(id) {
      if (!entries.delete(id)) throw notFound(`no cell formatter ${id}`);
      registry().unregisterFormatter(id);
    },
    get: (id) => entries.get(id) ?? null,
    list: () => [...entries.values()],
    on: (_name, listener) => registry().subscribe(listener),
  };
};

const createFormulaHints = (legacy: () => LegacyInstance): FormulaHints => {
  const hints = new Map<string, FormulaHint>();
  const registry = () => legacy().formula;
  const key = (name: string) => name.toUpperCase();
  return {
    add(entry) {
      hints.set(key(entry.name), entry);
      const off = registry().register(
        entry.name,
        () => {
          throw new FormulonError(
            StatusCode.NotImplemented,
            `${entry.name} is a hint, not a function`,
          );
        },
        {
          ...(entry.args ? { args: entry.args } : {}),
          ...(entry.description ? { description: entry.description } : {}),
        },
      );
      return () => {
        if (hints.get(key(entry.name)) !== entry) return;
        hints.delete(key(entry.name));
        off();
      };
    },
    remove(name) {
      if (!hints.delete(key(name))) throw notFound(`no formula hint ${name}`);
      registry().unregister(name);
    },
    has: (name) => hints.has(key(name)),
    get: (name) => hints.get(key(name)) ?? null,
    list: () => [...hints.values()],
    clear() {
      hints.clear();
      registry().clear();
    },
    on: (_name, listener) => registry().subscribe(listener),
  };
};

const legacyOptionsOf = (options: SpreadsheetOptions, handle: WorkbookHandle): MountOptions => {
  const host = options.host ?? {};
  const refresh = host.refreshPrinterProfiles;
  return {
    workbook: handle,
    ...(options.ui ? { ui: options.ui } : {}),
    // A switch named after a feature id reaches that feature even where `ui` has no switch of its own.
    ...(options.ui?.features ? { features: options.ui.features as FeatureFlags } : {}),
    ...(options.policy ? { policy: options.policy } : {}),
    ...(legacyViewport(options.viewport) ? { viewport: legacyViewport(options.viewport) } : {}),
    ...(options.contextMenu ? { contextMenu: options.contextMenu } : {}),
    ...(options.overlays ? { overlays: options.overlays } : {}),
    toolbar: options.toolbar != null,
    ...(options.theme ? { theme: options.theme } : {}),
    locale: options.locale ?? 'en',
    ...(options.strings ? { strings: options.strings } : {}),
    ...(options.extensions ? { extensions: [...options.extensions] } : {}),
    ...(host.workbookName ? { getWorkbookName: host.workbookName } : {}),
    ...(host.printerProfiles ? { printerProfiles: host.printerProfiles } : {}),
    ...(host.printerProfileId ? { printerProfileId: host.printerProfileId } : {}),
    ...(refresh
      ? {
          refreshPrinterProfiles: async () => (await refresh()) ?? undefined,
        }
      : {}),
    ...(host.captureScreenClip ? { captureScreenClip: host.captureScreenClip } : {}),
    ...(host.functionArgumentHelp ? { getFunctionArgumentHelp: host.functionArgumentHelp } : {}),
    ...(options.uploadStatus !== undefined ? { uploadStatus: options.uploadStatus } : {}),
    ...(options.macroRecording !== undefined ? { macroRecording: options.macroRecording } : {}),
    renderError: false,
  };
};

/**
 * Mounts a spreadsheet on `host` and resolves with its {@link SpreadsheetInstance}. The host's
 * children are replaced. Without `options.workbook` the instance creates a workbook and disposes it
 * with itself; a workbook passed in stays the caller's.
 *
 * Rejects with `TypeError` when `workbook` and `seed` are both given, with `workbook-in-use` when
 * another live instance shows the workbook, and with `options.signal.reason` once aborted.
 */
export async function mount(
  host: HTMLElement,
  options: SpreadsheetOptions = {},
): Promise<SpreadsheetInstance> {
  options.signal?.throwIfAborted();
  if (!(host instanceof HTMLElement)) throw new TypeError('mount: host must be an HTMLElement');
  if (options.workbook && options.seed) {
    throw new TypeError('mount: workbook and seed cannot be combined');
  }
  if (options.workbook && mountedWorkbooks.has(workbookHandleOf(options.workbook))) {
    throw spreadsheetFailure('workbook-in-use', 'the workbook is mounted by another spreadsheet');
  }
  const ownsWorkbook = !options.workbook;
  const workbook =
    options.workbook ??
    (await Workbook.createDefault({
      locale: options.locale ?? 'en',
      ...(options.ui?.platform && options.ui.platform !== 'auto'
        ? { platform: options.ui.platform }
        : {}),
      ...(options.signal ? { signal: options.signal } : {}),
    }));
  const handle = workbookHandleOf(workbook);
  let legacy: LegacyInstance | null = null;
  try {
    if (options.seed) options.seed(workbook);
    legacy = await Spreadsheet.mount(host, legacyOptionsOf(options, handle));
  } finally {
    // A failed seed or mount releases the workbook the instance created; one passed in stays.
    if (!legacy && ownsWorkbook) workbook.dispose();
  }
  if (options.signal?.aborted) {
    legacy.dispose();
    if (ownsWorkbook) workbook.dispose();
    throw options.signal.reason;
  }
  mountedWorkbooks.add(handle);

  const life = new Lifecycle('SpreadsheetInstance');
  const events = new Emitter<SpreadsheetInstanceEvents>();
  const selectionEvents = new Emitter<{ change: SelectionState }>();
  let state!: InstanceState;
  const current = (): InstanceState => {
    life.assertLive();
    return state;
  };
  const currentLegacy = (): LegacyInstance => current().legacy;
  const valueAt = (a: Addr): Value =>
    workbookHandleOf(state.workbook).engineWorkbook().sheet(a.sheet).getTypedValue(a.row, a.col);
  const addrOf = (ref: string): Addr =>
    resolveCellInput(
      ref,
      workbookHandleOf(state.workbook),
      state.legacy.store.getState().data.sheetIndex,
      'at',
    );
  const formulaHints = createFormulaHints(currentLegacy);
  state = {
    legacy,
    workbook,
    ownsWorkbook,
    options: { ...options, locale: options.locale ?? 'en' },
    origin: null,
    life,
    events,
    selection: createSelection(current, selectionEvents),
    history: createHistory(legacy.history),
    commands: createCommands(currentLegacy),
    dialogs: createDialogs(currentLegacy),
    audit: createAudit(currentLegacy, addrOf),
    viewport: {
      scrollTo(ref) {
        const target = currentLegacy();
        const addr = addrOf(ref);
        mutators.setActive(target.store, addr);
      },
      get zoomPercent() {
        const target = currentLegacy();
        return Math.round((target.store.getState().viewport.zoom ?? 1) * 100);
      },
      get options() {
        return current().options.viewport ?? { mode: 'auto' };
      },
    },
    extensions: createExtensions(currentLegacy),
    i18n: createI18n(currentLegacy),
    cells: createCellRegistry(currentLegacy, valueAt),
    formulaHints,
  };
  for (const hint of options.formulaHints ?? []) formulaHints.add(hint);
  const instance = new SpreadsheetInstance(CONSTRUCT, state);

  let lastSelection = selectionOf(legacy);
  const offSelection = legacy.store.subscribe(() => {
    const next = selectionOf(state.legacy);
    if (sameSelection(next, lastSelection)) return;
    lastSelection = next;
    selectionEvents.emit('change', next);
  });
  const offCell = legacy.on('cellChange', (e) =>
    events.emit('cell', {
      sheet: e.addr.sheet,
      at: { row: e.addr.row, col: e.addr.col },
      value: valueAt(e.addr),
      origin: state.origin ?? 'editor',
    }),
  );
  const offBatch = legacy.on('changeBatch', (result) =>
    events.emit('change', { ...legacyBatch(result), origin: state.origin ?? 'editor' }),
  );
  const offRecalc = legacy.on('recalc', (e) =>
    events.emit('recalc', {
      cells: [...e.dirty].map((key) => {
        const [sheet, row, col] = key.split(':').map(Number) as [number, number, number];
        return { sheet, at: { row, col } };
      }),
    }),
  );

  life.onDispose(() => {
    mountedWorkbooks.delete(workbookHandleOf(state.workbook));
    if (state.ownsWorkbook) state.workbook.dispose();
  });
  life.onDispose(() => state.legacy.dispose());
  life.onDispose(() => {
    offSelection();
    offCell();
    offBatch();
    offRecalc();
    events.clear();
    selectionEvents.clear();
  });
  return instance;
}

export type { ScreenClipResult };
