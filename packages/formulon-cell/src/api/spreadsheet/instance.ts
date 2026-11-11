import {
  type CellAddress,
  type CellRect,
  FormulonError,
  StatusCode,
  type Value,
} from '@libraz/formulon';
import { autofitColWidth, autofitRowHeight } from '../../commands/autofit-measurement.js';
import { isBuiltInCommandId } from '../../commands/built-in-command-ids.js';
import {
  BUILT_IN_COMMAND_OPERATION,
  canExecuteBuiltInRequest,
} from '../../commands/built-in-command-policy.js';
import { parseCSV } from '../../commands/clipboard/csv.js';
import type { PasteSpecialOptions } from '../../commands/clipboard/paste-special.js';
import type { History as LegacyHistory } from '../../commands/history.js';
import { createHistoryView } from '../../commands/history-view.js';
import { importCSV, MAX_CSV_CELLS } from '../../commands/import-export.js';
import type {
  HostCellChangeInput,
  InteractionOperation,
} from '../../commands/interaction-policy.js';
import { refusedBatch } from '../../commands/interaction-policy.js';
import { sheetPermits } from '../../commands/protection.js';
import type { SpellingService } from '../../commands/spelling.js';
import { clearTraceArrows, traceDependents, tracePrecedents } from '../../commands/traces.js';
import { resolveCellInput, resolveRangeInput } from '../../engine/range-input.js';
import type { Addr, Range } from '../../engine/types.js';
import { fromEngineValue } from '../../engine/value.js';
import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import type { Extension, ExtensionHandle, ExtensionInput } from '../../extensions/index.js';
import type { I18n } from '../../i18n/public-view.js';
import type { DeepPartial, Strings } from '../../i18n/strings.js';
import type { BorderDrawHandle } from '../../interact/border-draw.js';
import type { ClipboardHandle } from '../../interact/clipboard.js';
import type { FormatPainterHandle } from '../../interact/format-painter.js';
import type { ViewportOptions } from '../../interact/navigation-policy.js';
import { shareSpreadsheetInternals, spreadsheetInternals } from '../../mount/internals.js';
import type { MountToolbarOptions, ToolbarInstance } from '../../mount/toolbar.js';
import { runRibbonCommand } from '../../mount/toolbar.js';
import type {
  SpreadsheetInstance as LegacyInstance,
  MountOptions,
  ScreenClipResult,
} from '../../mount/types.js';
import { Spreadsheet } from '../../mount.js';
import { mutators } from '../../store/store.js';
import type { TraceArrow } from '../../store/types.js';
import {
  type InstanceRibbon,
  type RibbonEvents,
  RibbonView,
  ribbonStateOf,
} from '../../toolbar/ribbon-view.js';
import type { BatchCell, ChangeBatchResult } from '../change-batch.js';
import { failureOfDecision, spreadsheetFailure } from '../errors.js';
import { type Disposer, Emitter } from '../events.js';
import { Lifecycle } from '../lifecycle.js';
import type { ImportCsvOptions } from '../sheet-cells.js';
import {
  type ChangeOrigin,
  type SheetRef,
  Workbook,
  withChangeOrigin,
  workbookHandleOf,
} from '../workbook.js';
import { type CELL_COMMAND_IDS, type CommandId, isCommandId } from './command-ids.js';

type CellCommandId = (typeof CELL_COMMAND_IDS)[number];

/** The options the user can change from the UI. */
type UserOptionKey = 'showAllNotes' | 'theme' | 'locale';

import { createDialogs, type Dialogs } from './dialogs.js';
import type {
  ApplyChangesOptions,
  CellChange,
  FormulaHint,
  HostProviders,
  RibbonToolbarConfig,
  SelectionState,
  SheetPrintDocument,
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
  /** Whether `repeat` would run an operation. */
  readonly canRepeat: boolean;
  undo(): void;
  redo(): void;
  /** Runs the latest repeatable action again on the current selection (F4); false when there is none. */
  repeat(): boolean;
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
/** A question {@link Commands.can} answers besides a command id: whether the user may select,
 *  may copy, or may make an edit of `operation` to the cells `at` names. `sheet` and `at` take
 *  what `applyChanges` and the `Sheet` facade take: an A1 cell or range (sheet-qualified or on
 *  `sheet`, the active sheet when omitted), a `CellAddress` or a `CellRect`. */
export type CommandIntent =
  | { readonly kind: 'select' }
  | { readonly kind: 'copy' }
  | {
      readonly kind: 'edit';
      readonly operation: InteractionOperation;
      readonly sheet?: SheetRef;
      readonly at: string | CellAddress | CellRect;
    };

export interface Commands {
  /** Runs a command. Throws `policy-denied` or `protected` when it may not run. */
  execute(id: CommandId): void;
  /**
   * Whether `execute(id)` would run now, or whether the policy and sheet protection allow an
   * intent: an `edit` is decided as `applyChanges` decides a batch over the same cells. The answer
   * follows the policy, the protection and the selection, so check again after an `update`, a
   * `change` and a selection `change`.
   */
  can(id: CommandId): boolean;
  can(intent: CommandIntent): boolean;
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

/** One cell, addressed as `applyChanges` and `commands.can` address cells: an A1 cell
 *  (sheet-qualified or on `sheet`, the active sheet when omitted) or a `CellAddress`. */
export interface CellTarget {
  readonly sheet?: SheetRef;
  readonly at: string | CellAddress;
}

/** The cells the Watch Window shows: a view of this spreadsheet, not saved with the workbook. */
export interface Watches {
  /** Watches a cell; a cell already watched stays where it is. */
  add(cell: CellTarget): void;
  /** Stops watching a cell; throws `NotFound` when it is not watched. */
  remove(cell: CellTarget): void;
  has(cell: CellTarget): boolean;
  /** The watched cells, in the Watch Window's order. */
  list(): readonly BatchCell[];
  clear(): void;
  /** The watched cells changed, from the API or the Watch Window. */
  on(name: 'change', listener: () => void): Disposer;
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

export type { I18n };

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
  /** The current host providers; the mounted spreadsheet reads them through it at call time. */
  readonly host: { current: HostProviders };
  /** The host's spell checker, read when Review > Spelling runs. */
  readonly spelling: { current: SpellingService | undefined };
  /** The workbook the latest `update({ workbook })` asked for; differs from `workbook` while the
   *  mount is still switching to it. */
  requestedWorkbook: Workbook;
  /** Workbook swaps, run one at a time in request order. */
  swaps: Promise<void>;
  origin: ChangeOrigin | null;
  /** True while `update` applies a patch: the view changes it causes are the API's, not the user's. */
  applyingPatch: boolean;
  readonly ribbonEvents: Emitter<RibbonEvents>;
  /** Reports the ribbon's state to `toolbar.on('change')` listeners. */
  readonly notifyRibbon: () => void;
  /** The view of the current ribbon handle; a new handle gets a new view. */
  ribbon: { readonly handle: ToolbarInstance; readonly view: InstanceRibbon } | null;
  readonly life: Lifecycle;
  readonly events: Emitter<SpreadsheetInstanceEvents>;
  readonly selection: Selection;
  readonly history: History;
  readonly commands: Commands;
  readonly dialogs: Dialogs;
  readonly audit: Audit;
  readonly watches: Watches;
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

/** The mount behind a public instance, for the ribbon placed on its own. Exported from no entry
 *  point. */
export const legacyInstanceOf = (instance: SpreadsheetInstance): LegacyInstance =>
  stateOf(instance).legacy;

/** The ribbon options `ribbon(options)` carries, under the names the ribbon mount reads, with
 *  `notify` called on every state change the ribbon reports. */
const ribbonOptionsOf = (config: RibbonToolbarConfig, notify: () => void): MountToolbarOptions => {
  const { kind: _kind, ...options } = config as RibbonToolbarConfig & MountToolbarOptions;
  return {
    ...options,
    onTabChange: notify,
    onDisplayModeChange: notify,
    onFormulaBarChange: notify,
    onBackstageOpenChange: notify,
  };
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
  const refused = failureOfDecision(canExecuteBuiltInRequest(legacy.store, id, 'instanceApi'));
  if (refused) return refused;
  const operation = isBuiltInCommandId(id) ? BUILT_IN_COMMAND_OPERATION[id] : null;
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

  get watches(): Watches {
    return stateOf(this).watches;
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
    | (ClipboardHandle & { pasteSpecial(options: PasteSpecialOptions): ChangeBatchResult })
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
  /** The ribbon `update({ toolbar: ribbon() })` put above the grid, or `null` when there is none. */
  get toolbar(): InstanceRibbon | null {
    const state = stateOf(this);
    const handle = state.legacy.toolbar;
    if (!handle) return null;
    if (state.ribbon?.handle !== handle) {
      const view = new RibbonView(handle, state.legacy, state.ribbonEvents, () =>
        state.life.assertLive(),
      );
      state.ribbon = { handle, view };
    }
    return state.ribbon.view;
  }

  /** Applies a patch to the options: keys present are replaced, `null` resets one to its default,
   *  and keys left out keep their values. Emits `update` with the keys it changed and
   *  `source: 'api'`. */
  update(patch: SpreadsheetPatch): void {
    const state = stateOf(this);
    const keys = (Object.keys(patch) as (keyof SpreadsheetPatch)[]).filter(
      (key) => patch[key] !== undefined,
    );
    if (keys.length === 0) return;
    state.applyingPatch = true;
    try {
      for (const key of keys) applyPatchKey(state, key, patch[key] ?? null);
    } finally {
      state.applyingPatch = false;
    }
    const event: UpdateEvent = { keys, source: 'api' };
    state.events.emit('update', event);
  }

  /**
   * Imports CSV text from `options.at` as the user's import does: decided by the interaction
   * policy and sheet protection, all or nothing, and one undo step. A refusal, or more fields
   * than the CSV limit, is reported in `rejections`; nothing throws for it. `Sheet.importCsv` is
   * the ungated import for host code.
   */
  importCsv(text: string, options: ImportCsvOptions): ChangeBatchResult {
    const state = stateOf(this);
    if (typeof text !== 'string') throw new TypeError('importCsv: text must be a string');
    if (options?.at === undefined) throw new TypeError('importCsv: options.at is required');
    const legacy = state.legacy;
    const handle = workbookHandleOf(state.workbook);
    const active = legacy.store.getState().data.sheetIndex;
    const at = cellOfTarget(state.workbook, active, { at: options.at }, 'importCsv.at');
    const parse = options.delimiter ? { delimiter: options.delimiter } : {};
    const rows = parseCSV(text, parse);
    const width = rows.reduce((widest, row) => Math.max(widest, row.length), 0);
    if (rows.length * width > MAX_CSV_CELLS) {
      return refusedBatch({
        code: 'effectTooLarge',
        reason: `importCsv: ${rows.length * width} cells exceed the limit of ${MAX_CSV_CELLS}`,
      });
    }
    return importCSV(legacy.store, handle, text, at, legacy.history, {
      ...parse,
      ...(options.locale !== undefined ? { locale: options.locale } : {}),
    });
  }

  /** Writes cell changes as one batch, decided by the interaction policy and sheet protection in
   *  every history mode: all or nothing, so a refused change writes none of them and is reported
   *  in `rejections`; nothing throws for it. Ungated writes go through the Sheet facade's setters
   *  and `Sheet.importCsv`. */
  applyChanges(changes: readonly CellChange[], options: ApplyChangesOptions): ChangeBatchResult {
    const state = stateOf(this);
    const legacy = state.legacy;
    const active = legacy.store.getState().data.sheetIndex;
    const input: HostCellChangeInput[] = changes.map((change, index) => {
      const addr = cellOfTarget(state.workbook, active, change, `applyChanges[${index}].at`);
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
      return result;
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

  /** The print document for the active sheet's page setup, the document `print()` renders.
   *  Throws an error with code `PRINT_RESOURCE_LIMIT` when the print range is too large to build. */
  printDocument(): SheetPrintDocument {
    return stateOf(this).legacy.printDocument();
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
      legacy.setToolbar(
        value === null ? false : ribbonOptionsOf(value as RibbonToolbarConfig, state.notifyRibbon),
      );
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
      legacy.i18n.override(legacy.i18n.locale, (value as DeepPartial<Strings> | null) ?? {});
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
    case 'spelling':
      state.spelling.current = (value as SpellingService | null) ?? undefined;
      break;
    case 'host':
      applyHostProviders(state, (value as SpreadsheetOptions['host'] | null) ?? {});
      break;
    case 'uploadStatus':
      legacy.setUploadStatus((value as SpreadsheetOptions['uploadStatus'] | null) ?? null);
      break;
    case 'macroRecording':
      legacy.setMacroRecording((value as boolean | null) ?? null);
      break;
    case 'showAllNotes':
      mutators.setShowAllNotes(legacy.store, (value as boolean | null) ?? false);
      break;
  }
  state.options = next as SpreadsheetOptions;
};

const applyHostProviders = (state: InstanceState, host: HostProviders): void => {
  state.host.current = host;
  state.legacy.setPrinterProfiles(host.printerProfiles);
  state.legacy.setPrinterProfileId(host.printerProfileId);
};

const swapWorkbook = (state: InstanceState, next: Workbook | null): void => {
  if (!next) throw new TypeError('update({ workbook: null }): a workbook is required');
  if (next === state.requestedWorkbook) return;
  const nextHandle = workbookHandleOf(next);
  if (mountedWorkbooks.has(nextHandle)) {
    throw spreadsheetFailure('workbook-in-use', 'the workbook is mounted by another spreadsheet');
  }
  // Reserved now so no other instance mounts it while the engine switches.
  mountedWorkbooks.add(nextHandle);
  state.requestedWorkbook = next;
  state.swaps = state.swaps.then(() => runWorkbookSwap(state, next));
};

/** Switches the mount to `next`, then commits it to the instance; a failure keeps the current one. */
const runWorkbookSwap = async (state: InstanceState, next: Workbook): Promise<void> => {
  const nextHandle = workbookHandleOf(next);
  if (state.life.disposed) {
    mountedWorkbooks.delete(nextHandle);
    return;
  }
  try {
    await state.legacy.setWorkbook(nextHandle);
  } catch (error: unknown) {
    mountedWorkbooks.delete(nextHandle);
    if (state.requestedWorkbook === next) state.requestedWorkbook = state.workbook;
    if (!state.life.disposed) state.events.emit('error', { error, source: 'workbook' });
    return;
  }
  if (state.life.disposed) {
    mountedWorkbooks.delete(nextHandle);
    return;
  }
  const previous = state.workbook;
  mountedWorkbooks.delete(workbookHandleOf(previous));
  state.workbook = next;
  state.options = { ...state.options, workbook: next };
  if (state.ownsWorkbook) previous.dispose();
  state.ownsWorkbook = false;
  state.events.emit('workbook', { workbook: next });
};

const createHistory = (history: LegacyHistory): History => createHistoryView(history);

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

/** The cell `target` names; its sheet defaults to `active`. */
const cellOfTarget = (
  workbook: Workbook,
  active: number,
  target: CellTarget,
  label: string,
): Addr => {
  const sheet = target.sheet === undefined ? active : workbook.sheet(target.sheet).index;
  return typeof target.at === 'string'
    ? resolveCellInput(target.at, workbookHandleOf(workbook), sheet, label)
    : { sheet, row: target.at.row, col: target.at.col };
};

const intentRange = (target: LegacyInstance, workbook: Workbook, intent: EditIntent): Range => {
  const sheet =
    intent.sheet === undefined
      ? target.store.getState().data.sheetIndex
      : workbook.sheet(intent.sheet).index;
  const { at } = intent;
  if (typeof at === 'string') return resolveRangeInput(at, workbookHandleOf(workbook), sheet, 'at');
  if ('firstRow' in at)
    return { sheet, r0: at.firstRow, c0: at.firstCol, r1: at.lastRow, c1: at.lastCol };
  return { sheet, r0: at.row, c0: at.col, r1: at.row, c1: at.col };
};

type EditIntent = Extract<CommandIntent, { kind: 'edit' }>;

const intentAllowed = (target: LegacyInstance, workbook: Workbook, intent: CommandIntent) => {
  const controller = target.commands;
  if (intent.kind === 'select') return controller.canSelect().allowed;
  if (intent.kind === 'copy') return controller.canCopy().allowed;
  const range = intentRange(target, workbook, intent);
  const base = { operation: intent.operation, origin: 'instanceApi' } as const;
  return (
    controller.canExecute({ ...base, effects: [{ kind: 'workbook' }] }).allowed &&
    controller.canExecute({ ...base, effects: [{ kind: 'range', range }] }).allowed
  );
};

/** How each cell command runs; ribbon commands run through the ribbon. */
const CELL_COMMANDS: Record<CellCommandId, (target: LegacyInstance) => void> = {
  openHyperlink: (target) => {
    target.followActiveHyperlink();
  },
};

const isCellCommandId = (id: CommandId): id is CellCommandId => id in CELL_COMMANDS;

const createCommands = (legacy: () => LegacyInstance, workbook: () => Workbook): Commands => ({
  execute(id) {
    if (!isCommandId(id)) throw notFound(`unknown command ${String(id)}`);
    const target = legacy();
    const refusal = commandRefusal(target, id);
    if (refusal)
      throw spreadsheetFailure(refusal, `command ${id} may not run`, { context: { command: id } });
    if (isCellCommandId(id)) CELL_COMMANDS[id](target);
    else if (target.toolbar) target.toolbar.applyCommand(id);
    else runRibbonCommand(target, id);
  },
  can(request: CommandId | CommandIntent) {
    if (typeof request === 'object') return intentAllowed(legacy(), workbook(), request);
    return isCommandId(request) && commandRefusal(legacy(), request) === null;
  },
});

const createWatches = (
  legacy: () => LegacyInstance,
  addrOf: (cell: CellTarget) => Addr,
): Watches => {
  const watched = () => legacy().store.getState().watch.watches;
  const isWatched = (addr: Addr): boolean =>
    watched().some((w) => w.sheet === addr.sheet && w.row === addr.row && w.col === addr.col);
  return {
    add: (cell) => mutators.addWatch(legacy().store, addrOf(cell)),
    remove(cell) {
      const addr = addrOf(cell);
      if (!isWatched(addr)) throw notFound(`cell ${addr.row}:${addr.col} is not watched`);
      mutators.removeWatch(legacy().store, addr);
    },
    has: (cell) => isWatched(addrOf(cell)),
    list: () => watched().map(({ sheet, row, col }) => ({ sheet, at: { row, col } })),
    clear: () => mutators.clearWatches(legacy().store),
    on(_name, listener) {
      let last = watched();
      return legacy().store.subscribe(() => {
        const next = watched();
        if (next === last) return;
        last = next;
        listener();
      });
    },
  };
};

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
    legacy().i18n.override(locale, strings);
  },
  on(_name, listener) {
    const i18n = legacy().i18n;
    return i18n.on('change', (strings) => listener({ locale: i18n.locale, strings }));
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
      const off = registry().register(entry.name, {
        ...(entry.args ? { args: entry.args } : {}),
        ...(entry.description ? { description: entry.description } : {}),
      });
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

const legacyOptionsOf = (
  options: SpreadsheetOptions,
  handle: WorkbookHandle,
  hostProviders: { readonly current: HostProviders },
  spelling: { readonly current: SpellingService | undefined },
  notifyRibbon: () => void,
): MountOptions => {
  const host = options.host ?? {};
  return {
    workbook: handle,
    ...(options.ui ? { ui: options.ui } : {}),
    // A switch named after a feature id reaches that feature even where `ui` has no switch of its own.
    ...(options.ui?.features ? { features: options.ui.features } : {}),
    ...(options.policy ? { policy: options.policy } : {}),
    ...(legacyViewport(options.viewport) ? { viewport: legacyViewport(options.viewport) } : {}),
    ...(options.contextMenu ? { contextMenu: options.contextMenu } : {}),
    ...(options.overlays ? { overlays: options.overlays } : {}),
    toolbar: options.toolbar != null && ribbonOptionsOf(options.toolbar, notifyRibbon),
    ...(options.theme ? { theme: options.theme } : {}),
    locale: options.locale ?? 'en',
    ...(options.strings ? { strings: options.strings } : {}),
    ...(options.extensions ? { extensions: [...options.extensions] } : {}),
    ...(host.printerProfiles ? { printerProfiles: host.printerProfiles } : {}),
    ...(host.printerProfileId ? { printerProfileId: host.printerProfileId } : {}),
    // Function providers forward to the current set, so `update({ host })` replaces them.
    getWorkbookName: () => hostProviders.current.workbookName?.() ?? null,
    refreshPrinterProfiles: async () =>
      (await hostProviders.current.refreshPrinterProfiles?.()) ?? undefined,
    captureScreenClip: () => hostProviders.current.captureScreenClip?.(),
    getFunctionArgumentHelp: (name, index, locale) =>
      hostProviders.current.functionArgumentHelp?.(name, index, locale),
    getSpellingService: () => spelling.current,
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
  const hostProviders = { current: options.host ?? {} };
  const spelling = { current: options.spelling };
  const ribbonEvents = new Emitter<RibbonEvents>();
  const notifyRibbon = (): void => {
    const ribbonHandle = legacy?.toolbar;
    if (ribbonHandle) ribbonEvents.emit('change', ribbonStateOf(ribbonHandle));
  };
  try {
    if (options.seed) options.seed(workbook);
    legacy = await Spreadsheet.mount(
      host,
      legacyOptionsOf(options, handle, hostProviders, spelling, notifyRibbon),
    );
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
  const targetOf = (cell: CellTarget): Addr =>
    cellOfTarget(state.workbook, state.legacy.store.getState().data.sheetIndex, cell, 'at');
  const formulaHints = createFormulaHints(currentLegacy);
  state = {
    legacy,
    workbook,
    ownsWorkbook,
    options: { ...options, locale: options.locale ?? 'en' },
    host: hostProviders,
    spelling,
    requestedWorkbook: workbook,
    swaps: Promise.resolve(),
    origin: null,
    applyingPatch: false,
    ribbonEvents,
    notifyRibbon,
    ribbon: null,
    life,
    events,
    selection: createSelection(current, selectionEvents),
    history: createHistory(legacy.history),
    commands: createCommands(currentLegacy, () => current().workbook),
    dialogs: createDialogs(currentLegacy),
    audit: createAudit(currentLegacy, addrOf),
    watches: createWatches(currentLegacy, targetOf),
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

  if (options.showAllNotes) mutators.setShowAllNotes(legacy.store, true);
  // A view preference the user changes from the UI updates the options and reports `update`
  // with `source: 'user'`; a change `update` makes is already reported as the API's.
  const userOption = <K extends UserOptionKey>(key: K, initial: SpreadsheetOptions[K]) => {
    let last = initial;
    return (value: SpreadsheetOptions[K]): void => {
      if (value === last) return;
      last = value;
      if (state.applyingPatch) return;
      state.options = { ...state.options, [key]: value };
      events.emit('update', { keys: [key], source: 'user' });
    };
  };
  const reportShowAllNotes = userOption('showAllNotes', legacy.store.getState().ui.showAllNotes);
  const offShowAllNotes = legacy.store.subscribe(() =>
    reportShowAllNotes(state.legacy.store.getState().ui.showAllNotes),
  );
  const reportTheme = userOption('theme', legacy.store.getState().ui.theme);
  const offTheme = legacy.on('themeChange', (e) => reportTheme(e.theme));
  const reportLocale = userOption('locale', legacy.i18n.locale);
  const offLocale = legacy.on('localeChange', (e) => reportLocale(e.locale));

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
    events.emit('change', { ...result, origin: state.origin ?? 'editor' }),
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
    offShowAllNotes();
    offTheme();
    offLocale();
    offCell();
    offBatch();
    offRecalc();
    events.clear();
    ribbonEvents.clear();
    selectionEvents.clear();
  });
  return instance;
}

export type { ScreenClipResult };
