import type { CellRegistry } from '../cells.js';
import type { PasteSpecialOptions } from '../commands/clipboard/paste-special.js';
import type { History } from '../commands/history.js';
import type { PrinterProfile } from '../commands/printer-profile.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SpreadsheetEventHandler, SpreadsheetEventName } from '../events.js';
import type {
  ExtensionHandle,
  ExtensionInput,
  FeatureFlags,
  SpreadsheetUiOptions,
  ThemeName,
} from '../extensions/index.js';
import type { CustomFunction, CustomFunctionMeta, FormulaRegistry } from '../formula.js';
import type { I18nController } from '../i18n/controller.js';
import type { DeepPartial, Locale, Strings } from '../i18n/strings.js';
import type { BorderDrawHandle } from '../interact/border-draw.js';
import type { ClipboardHandle } from '../interact/clipboard.js';
import type { ConditionalDialogOpenOptions } from '../interact/conditional-dialog.js';
import type { FormatPainterHandle } from '../interact/format-painter.js';
import type { PasteSpecialOpenOptions } from '../interact/paste-special.js';
import type { StatusBarUploadStatus } from '../interact/status-bar.js';
import type { SlicerSpec, SpreadsheetStore } from '../store/store.js';

export interface MountOptions {
  workbook?: WorkbookHandle;
  ui?: SpreadsheetUiOptions;
  theme?: ThemeName;
  seed?: (wb: WorkbookHandle) => void;
  locale?: Locale | (string & {});
  strings?: DeepPartial<Strings>;
  features?: FeatureFlags;
  extensions?: ExtensionInput[];
  functions?: readonly {
    name: string;
    impl: CustomFunction['impl'];
    meta?: CustomFunctionMeta;
  }[];
  /** Optional host-provided printer profiles. Browser print APIs do not expose
   *  physical printer non-printable areas, so hosts can pass paper/orientation
   *  profiles here for the built-in print/PDF flow. */
  printerProfiles?: readonly PrinterProfile[];
  /** Optional active host printer profile id. When omitted, the print flow
   *  chooses the best profile for the current paper size and orientation. */
  printerProfileId?: string;
  /** Optional host refresh hook for native/Electron printer discovery. Browser
   *  APIs do not expose the physical printer list or non-printable areas. */
  refreshPrinterProfiles?: () =>
    | readonly PrinterProfile[]
    | undefined
    | Promise<readonly PrinterProfile[] | undefined>;
  /** Optional host capture hook for Insert > Screenshot > Screen Clipping.
   *  Browsers cannot invoke Excel-style OS region capture directly, so native
   *  shells can return an image source here. */
  captureScreenClip?: ScreenClipCapture;
  /** Optional host-driven status bar Upload Status indicator. */
  uploadStatus?: StatusBarUploadStatus;
  /** Optional host-driven status bar Macro Recording indicator. */
  macroRecording?: boolean | null;
  /** Called when mount fails before an instance exists, most commonly when
   *  the WASM engine cannot start because SharedArrayBuffer is unavailable. */
  onError?: (error: unknown) => void;
  /** Whether core should render its built-in mount error panel into `host`
   *  before rejecting. Defaults to true. Wrappers can disable this and render
   *  their framework-native fallback instead. */
  renderError?: boolean;
}

export interface ScreenClipResult {
  src: string;
  alt?: string;
}

export type ScreenClipCaptureResult = string | ScreenClipResult | null | undefined;

export type ScreenClipCapture = () => ScreenClipCaptureResult | Promise<ScreenClipCaptureResult>;

export interface SpreadsheetInstance {
  readonly host: HTMLElement;
  readonly workbook: WorkbookHandle;
  readonly store: SpreadsheetStore;
  readonly history: History;
  readonly i18n: I18nController;
  readonly features: Readonly<Record<string, ExtensionHandle | undefined>>;
  /** Clipboard handle the engine binding produced. `null` when the
   *  `clipboard` feature flag is off (e.g. read-only embeds). Ribbon and
   *  context-menu actions read this so they can route Copy/Cut/Paste through
   *  `runShortcut` — `document.execCommand` doesn't fire copy/paste events
   *  on the non-editable grid host. */
  readonly clipboard: ClipboardHandle | null;
  readonly formatPainter: FormatPainterHandle | undefined;
  readonly borderDraw: BorderDrawHandle | undefined;
  readonly formula: FormulaRegistry;
  readonly cells: CellRegistry;
  use(ext: ExtensionInput): void;
  remove(id: string): boolean;
  setFeatures(next: FeatureFlags): void;
  setExtensions(next: ExtensionInput[] | undefined): void;
  openConditionalDialog(options?: ConditionalDialogOpenOptions): void;
  openNamedRangeDialog(): void;
  openFormatDialog(
    tab?: 'number' | 'align' | 'font' | 'border' | 'fill' | 'protection' | 'more',
  ): void;
  openDataValidationDialog(): void;
  openGoTo(): void;
  openGoToSpecial(): void;
  openFilterDropdown(range?: import('../engine/types.js').Range, col?: number): void;
  openIterativeDialog(): void;
  openExternalLinksDialog(): void;
  openCfRulesDialog(): void;
  openCellStylesGallery(): void;
  openEvaluateFormulaDialog(): void;
  openFunctionArguments(seedName?: string): void;
  openHyperlinkDialog(): void;
  openCommentDialog(): void;
  openDefineNameDialog(): void;
  openFindReplace(tab?: 'find' | 'replace'): void;
  closeFindReplace(): void;
  openPasteSpecial(opts?: PasteSpecialOpenOptions): void;
  pasteSpecial(options: PasteSpecialOptions, opts?: PasteSpecialOpenOptions): boolean;
  openInsertCopiedCells(): void;
  openPageSetup(): void;
  print(mode?: 'print' | 'pdf'): void;
  setPrinterProfiles(next: readonly PrinterProfile[] | undefined): void;
  setPrinterProfileId(next: string | undefined): void;
  refreshPrinterProfiles(): Promise<readonly PrinterProfile[] | undefined>;
  captureScreenClip(): Promise<ScreenClipResult | null>;
  setUploadStatus(next: StatusBarUploadStatus): void;
  setMacroRecording(next: boolean | null): void;
  recalc(): void;
  openWatchWindow(): void;
  closeWatchWindow(): void;
  toggleWatchWindow(): void;
  openQuickAnalysis(): void;
  openWorkbookObjects(): void;
  openPivotFieldList(sheetIndex: number, pivotIndex: number): boolean;
  openActivePivotFieldList(): boolean;
  openPivotTableDialog(): void;
  addSlicer(input: {
    tableName: string;
    column: string;
    selected?: readonly string[];
    x?: number;
    y?: number;
  }): SlicerSpec;
  removeSlicer(id: string): void;
  toggleSheetProtection(): void;
  setSheetProtected(on: boolean, password?: string): void;
  isSheetProtected(): boolean;
  tracePrecedents(): number;
  traceDependents(): number;
  clearTraces(): void;
  setTheme(t: ThemeName): void;
  undo(): boolean;
  redo(): boolean;
  setWorkbook(next: WorkbookHandle): Promise<void>;
  on<K extends SpreadsheetEventName>(name: K, fn: SpreadsheetEventHandler<K>): () => void;
  off<K extends SpreadsheetEventName>(name: K, fn: SpreadsheetEventHandler<K>): void;
  dispose(): void;
}
