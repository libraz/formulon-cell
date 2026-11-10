import type { CellAddress, CellRect, Value } from '@libraz/formulon';
import type { InteractionPolicy } from '../../commands/interaction-policy.js';
import type { PrinterProfile } from '../../commands/printer-profile.js';
import type { ExtensionInput, SpreadsheetUiOptions, ThemeName } from '../../extensions/index.js';
import type { DeepPartial, Locale, Strings } from '../../i18n/strings.js';
import type { ContextMenuOptions } from '../../interact/context-menu-options.js';
import type { FunctionArgumentHelpProvider } from '../../interact/mac-formula-palette.js';
import type { OverlayOptions } from '../../interact/overlay-portal.js';
import type { StatusBarUploadStatus } from '../../interact/status-bar.js';
import type { ScreenClipCapture } from '../../mount/types.js';
import type { ChangeBatchResult } from '../change-batch.js';
import type { ChangedCell, ChangeOrigin, SheetRef, Workbook } from '../workbook.js';

/** The visible region: the whole sheet, or a fixed range the selection stays inside. */
export type ViewportConfig =
  | { readonly mode: 'auto' }
  | {
      readonly mode: 'fixed';
      /** A1 range, optionally sheet-qualified (`'Input!A1:D20'`). */
      readonly range: string;
      /** `'editable'` moves Tab between the cells a policy lets the user edit. */
      readonly tabNavigation?: 'normal' | 'editable';
    };

/** The ribbon toolbar to mount inside the host; create it with `ribbon()` from `./toolbar`. */
export interface RibbonToolbarConfig {
  readonly kind: 'ribbon';
}

/** Host-side providers the spreadsheet calls; replace them with `update({ host })`. */
export interface HostProviders {
  /** The file name or title the Watch Window shows. */
  readonly workbookName?: () => string | null;
  /** Printer profiles the print and PDF flow chooses from. */
  readonly printerProfiles?: readonly PrinterProfile[];
  /** Id of the active printer profile; the best match for the page setup when omitted. */
  readonly printerProfileId?: string;
  /** Re-reads the printer list; called by `refreshPrinterProfiles()`. */
  readonly refreshPrinterProfiles?: () =>
    | readonly PrinterProfile[]
    | null
    | Promise<readonly PrinterProfile[] | null>;
  /** Captures a screen region for Insert > Screenshot; called by `captureScreenClip()`. */
  readonly captureScreenClip?: ScreenClipCapture;
  /** Per-argument help for the Mac formula palette. */
  readonly functionArgumentHelp?: FunctionArgumentHelpProvider;
}

/** A host formula hint: what autocomplete and the argument helper show for a name. It does not add
 *  a worksheet function; the calculator never calls it. */
export interface FormulaHint {
  /** Function name, matched case-insensitively. */
  readonly name: string;
  /** Argument labels; wrap an optional one in brackets (`'[step]'`). */
  readonly args?: readonly string[];
  readonly description?: string;
}

/** What `mount` takes. */
export interface SpreadsheetOptions {
  /** The workbook to show; the caller keeps owning it. Omitted: the instance creates and owns one. */
  readonly workbook?: Workbook;
  /** Fills the workbook the instance creates; not allowed together with `workbook`. */
  readonly seed?: (workbook: Workbook) => void;
  readonly ui?: SpreadsheetUiOptions;
  /** Restricts what the user may do; `null` lifts every restriction. */
  readonly policy?: InteractionPolicy | null;
  readonly viewport?: ViewportConfig;
  readonly contextMenu?: ContextMenuOptions;
  readonly overlays?: OverlayOptions;
  /** The ribbon to mount above the grid; `null` for none. */
  readonly toolbar?: RibbonToolbarConfig | null;
  readonly theme?: ThemeName;
  /** UI language and the workbook's calculation profile. Defaults to `'en'`. */
  readonly locale?: Locale | (string & {});
  /** Overrides on top of the locale's strings. */
  readonly strings?: DeepPartial<Strings>;
  readonly extensions?: readonly ExtensionInput[];
  readonly formulaHints?: readonly FormulaHint[];
  readonly host?: HostProviders;
  readonly uploadStatus?: StatusBarUploadStatus;
  readonly macroRecording?: boolean | null;
  /** Cancels the mount; an aborted signal rejects with its reason. */
  readonly signal?: AbortSignal;
}

/** What `update` takes: any option except the mount-only `seed` and `signal`. A `null` resets a
 *  nullable key to its default. */
export type SpreadsheetPatch = {
  readonly [K in Exclude<keyof SpreadsheetOptions, 'seed' | 'signal'>]?:
    | SpreadsheetOptions[K]
    | null;
};

/** The current selection. */
export interface SelectionState {
  /** Index of the active sheet. */
  readonly sheet: number;
  readonly active: CellAddress;
  /** The range the active cell is in. */
  readonly range: CellRect;
  /** Every selected range, the active one first. */
  readonly ranges: readonly CellRect[];
}

/** What a batch of host or user changes did, and where it came from. */
export interface ChangeEvent extends ChangeBatchResult {
  readonly origin: ChangeOrigin;
}

/** One cell whose value changed. */
export interface CellEvent {
  readonly sheet: number;
  readonly at: CellAddress;
  readonly value: Value;
  readonly origin: ChangeOrigin;
}

/** The instance now shows another workbook. */
export interface WorkbookEvent {
  readonly workbook: Workbook;
}

/** A recalculation ran; `cells` are the writes it answered (dependents are not listed). */
export interface RecalcEvent {
  readonly cells: readonly ChangedCell[];
}

/** `update` changed these option keys. */
export interface UpdateEvent {
  readonly keys: readonly (keyof SpreadsheetPatch)[];
}

/** A failure after `mount` resolved that no call could throw: an extension's setup, a host
 *  provider's rejection, a deferred engine failure. */
export interface ErrorEvent {
  readonly error: unknown;
  readonly source: 'extension' | 'host' | 'engine';
}

/** Events of a {@link SpreadsheetInstance}. */
export interface SpreadsheetInstanceEvents {
  change: ChangeEvent;
  cell: CellEvent;
  workbook: WorkbookEvent;
  recalc: RecalcEvent;
  error: ErrorEvent;
  update: UpdateEvent;
}

/** A host or import write: text as typed, or a value with an optional formula. `at` may be
 *  sheet-qualified; an omitted `sheet` means the active sheet. */
export type CellChange =
  | { readonly sheet?: SheetRef; readonly at: string | CellAddress; readonly input: string }
  | {
      readonly sheet?: SheetRef;
      readonly at: string | CellAddress;
      readonly value: Value;
      readonly formula?: string | null;
    };

/** How `applyChanges` treats the undo history: `'record'` makes the batch one undo step,
 *  `'reset'` clears the history afterwards. */
export interface ApplyChangesOptions {
  readonly history: 'reset' | 'record';
  readonly origin?: ChangeOrigin;
}
