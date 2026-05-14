import type { CellRegistry } from '../cells.js';
import type { History } from '../commands/history.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SpreadsheetEventHandler, SpreadsheetEventName } from '../events.js';
import type {
  ExtensionHandle,
  ExtensionInput,
  FeatureFlags,
  ThemeName,
} from '../extensions/index.js';
import type { CustomFunction, CustomFunctionMeta, FormulaRegistry } from '../formula.js';
import type { I18nController } from '../i18n/controller.js';
import type { DeepPartial, Locale, Strings } from '../i18n/strings.js';
import type { FormatPainterHandle } from '../interact/format-painter.js';
import type { SlicerSpec, SpreadsheetStore } from '../store/store.js';

export interface MountOptions {
  workbook?: WorkbookHandle;
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
  /** Called when mount fails before an instance exists, most commonly when
   *  the WASM engine cannot start because SharedArrayBuffer is unavailable. */
  onError?: (error: unknown) => void;
  /** Whether core should render its built-in mount error panel into `host`
   *  before rejecting. Defaults to true. Wrappers can disable this and render
   *  their framework-native fallback instead. */
  renderError?: boolean;
}

export interface SpreadsheetInstance {
  readonly host: HTMLElement;
  readonly workbook: WorkbookHandle;
  readonly store: SpreadsheetStore;
  readonly history: History;
  readonly i18n: I18nController;
  readonly features: Readonly<Record<string, ExtensionHandle | undefined>>;
  readonly formatPainter: FormatPainterHandle | undefined;
  readonly formula: FormulaRegistry;
  readonly cells: CellRegistry;
  use(ext: ExtensionInput): void;
  remove(id: string): boolean;
  setFeatures(next: FeatureFlags): void;
  setExtensions(next: ExtensionInput[] | undefined): void;
  openConditionalDialog(): void;
  openNamedRangeDialog(): void;
  openFormatDialog(): void;
  openGoToSpecial(): void;
  openIterativeDialog(): void;
  openExternalLinksDialog(): void;
  openCfRulesDialog(): void;
  openCellStylesGallery(): void;
  openFunctionArguments(seedName?: string): void;
  openHyperlinkDialog(): void;
  openCommentDialog(): void;
  openFindReplace(): void;
  closeFindReplace(): void;
  openPasteSpecial(): void;
  openPageSetup(): void;
  print(): void;
  recalc(): void;
  openWatchWindow(): void;
  closeWatchWindow(): void;
  toggleWatchWindow(): void;
  openQuickAnalysis(): void;
  openWorkbookObjects(): void;
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
  tracePrecedents(): void;
  traceDependents(): void;
  clearTraces(): void;
  setTheme(t: ThemeName): void;
  undo(): boolean;
  redo(): boolean;
  setWorkbook(next: WorkbookHandle): Promise<void>;
  on<K extends SpreadsheetEventName>(name: K, fn: SpreadsheetEventHandler<K>): () => void;
  off<K extends SpreadsheetEventName>(name: K, fn: SpreadsheetEventHandler<K>): void;
  dispose(): void;
}
