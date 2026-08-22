// Re-export of the formulon-typed surface plus our adapter shapes.
export type {
  CellEntry,
  CellResult,
  CellXf,
  ColorSpec,
  ConditionalFormatEntry,
  ConditionalFormatInput,
  DataValidationEntry,
  DataValidationInput,
  DataValidationRange,
  DxfResult,
  EvalArrayResult,
  EvalResult,
  FormulonModule,
  FunctionMetadataEntry,
  FunctionMetadataLocalized,
  FunctionMetadataProvider,
  FunctionMetadataResult,
  MergedFunctionMetadataResult,
  PivotCell,
  PivotLayoutResult,
  PivotReportLayoutResult,
  PivotWorksheetSource,
  PivotWorksheetSourceResult,
  SaveResult,
  Status,
  StringResult,
  TableInput,
  Value,
  Workbook,
} from '@libraz/formulon';

/** Marks the fields a record only carries when the engine read it back out of
 * a workbook: the original OOXML `<color>` selector and the "was this element
 * present in the source" flags. The cell layer authors records from UI state,
 * where every colour is literal RGB and no such provenance exists, so these
 * stay optional at the adapter boundary. `WorkbookHandle` fills them in on the
 * way into the engine — see `completeFontRecord` and friends. */
type EngineFilled<T, K extends keyof T> = Omit<T, K> & Partial<Pick<T, K>>;

export type FontRecord = EngineFilled<
  import('@libraz/formulon').FontRecord,
  | 'vertAlign'
  | 'hasBold'
  | 'hasItalic'
  | 'hasStrike'
  | 'hasFamily'
  | 'family'
  | 'hasCharset'
  | 'charset'
  | 'color'
>;

export type FillRecord = EngineFilled<import('@libraz/formulon').FillRecord, 'fg' | 'bg'>;

export type BorderSide = EngineFilled<import('@libraz/formulon').BorderSide, 'color'>;

export interface BorderRecord {
  left: BorderSide;
  right: BorderSide;
  top: BorderSide;
  bottom: BorderSide;
  diagonal: BorderSide;
  diagonalUp: boolean;
  diagonalDown: boolean;
}

/** Differential format whose sub-records use the adapter's authoring shapes. */
export type DxfRecord = Omit<import('@libraz/formulon').DxfRecord, 'font' | 'fill' | 'border'> & {
  font?: FontRecord;
  fill?: FillRecord;
  border?: BorderRecord;
};

export type SpreadsheetProfileId = 'windows-ja_JP' | 'mac-ja_JP';

/** PivotTable axis ordinals. Mirrors `fm_pivot_axis_t`. */
export const PivotAxis = {
  Row: 0,
  Col: 1,
  Value: 2,
  Page: 3,
} as const;
export type PivotAxis = (typeof PivotAxis)[keyof typeof PivotAxis];

export interface PivotFieldSpec {
  readonly sourceName: string;
  readonly axis: PivotAxis;
  readonly subtotalTop?: boolean;
}

/** Aggregation function ordinals for value-axis fields. */
export const PivotAggregation = {
  Sum: 0,
  Count: 1,
  Average: 2,
  Max: 3,
  Min: 4,
  Product: 5,
  CountNumbers: 6,
  StdDev: 7,
  StdDevP: 8,
  Var: 9,
  VarP: 10,
} as const;
export type PivotAggregation = (typeof PivotAggregation)[keyof typeof PivotAggregation];

/** Show-values-as derivation ordinals for PivotTable data fields. */
export const PivotShowValuesAs = {
  Normal: 0,
  PercentOfRow: 1,
  PercentOfCol: 2,
  PercentOfTotal: 3,
  RunningTotalInRow: 4,
  RunningTotalInCol: 5,
  Index: 6,
  DifferenceFrom: 7,
  PercentDifferenceFrom: 8,
  PercentOfParentRow: 9,
  PercentOfParentCol: 10,
  PercentOfParent: 11,
} as const;
export type PivotShowValuesAs = (typeof PivotShowValuesAs)[keyof typeof PivotShowValuesAs];

/** Sentinel values for `PivotDataFieldSpec.showAsBaseItem`. */
export const PIVOT_SHOW_AS_BASE_PREVIOUS = 1048828;
export const PIVOT_SHOW_AS_BASE_NEXT = 1048829;

export interface PivotDataFieldSpec {
  readonly name?: string;
  readonly fieldIndex: number;
  readonly aggregation: PivotAggregation;
  readonly numberFormat?: string;
  readonly showValuesAs?: PivotShowValuesAs;
  readonly showAsBaseField?: number;
  readonly showAsBaseItem?: number;
}

/** PivotTable filter type ordinals. */
export const PivotFilterType = {
  ValueTop10: 0,
  ValueGreaterThan: 1,
  ValueBetween: 2,
  LabelContains: 3,
  LabelBeginsWith: 4,
  LabelDate: 5,
  LabelEquals: 6,
  LabelDoesNotEqual: 7,
  LabelDoesNotContain: 8,
  LabelEndsWith: 9,
  ValueLessThan: 10,
  ValueEquals: 11,
  ValueNotBetween: 12,
  DateBefore: 13,
  DateAfter: 14,
  DateBetween: 15,
} as const;
export type PivotFilterType = (typeof PivotFilterType)[keyof typeof PivotFilterType];

/** PivotTable date grouping ordinals. */
export const PivotDateGrouping = {
  Day: 0,
  Month: 1,
  Quarter: 2,
  Year: 3,
  Week: 4,
  Hour: 5,
  Minute: 6,
  Second: 7,
} as const;
export type PivotDateGrouping = (typeof PivotDateGrouping)[keyof typeof PivotDateGrouping];

/** PivotTable calendar ordinals. */
export const PivotCalendar = {
  Gregorian: 0,
  Japanese: 1,
} as const;
export type PivotCalendar = (typeof PivotCalendar)[keyof typeof PivotCalendar];

/** Pivot report layout form. Mirrors `fm_pivot_layout_t`. */
export const PivotReportLayout = {
  Compact: 0,
  Tabular: 1,
  Outline: 2,
} as const;
export type PivotReportLayout = (typeof PivotReportLayout)[keyof typeof PivotReportLayout];

/** Worksheet page setup as the engine models it. `orientation` follows OOXML
 *  `<pageSetup orientation>`: 0 leaves the attribute off (the printer
 *  default), 1 is portrait, 2 is landscape. `scale` is a percentage, not the
 *  0..1 fraction the store keeps. The `*Stated` flags say whether the sheet
 *  declares the attribute at all, which is the only way to tell a written
 *  `scale="100"` from an absent one. */
export type EngineOrientation = 0 | 1 | 2;

export interface EnginePageSetup {
  readonly orientation: EngineOrientation;
  readonly paperSize: number;
  readonly scale: number;
  readonly fitToWidth: number;
  readonly fitToHeight: number;
  readonly fitToPage: boolean;
  readonly orientationStated: boolean;
  readonly paperSizeStated: boolean;
  readonly scaleStated: boolean;
  readonly fitToPageStated: boolean;
}

/**
 * The `<pageSetup>` attributes the engine's typed setter does not model.
 *
 * They survive a load and save untouched — the engine keeps the raw fragment
 * as the writer's source of truth — but authoring one means going through that
 * fragment. Values use OOXML's own vocabulary rather than the store's, so the
 * translation stays in one place (`print-sync`). An empty string or `null`
 * means the attribute is absent.
 */
export interface EnginePageSetupExtras {
  readonly blackAndWhite: boolean;
  readonly draft: boolean;
  /** `none` | `asDisplayed` | `atEnd`. */
  readonly cellComments: string;
  /** `displayed` | `blank` | `dash` | `NA`. */
  readonly errors: string;
  /** `downThenOver` | `overThenDown`. */
  readonly pageOrder: string;
  readonly firstPageNumber: number | null;
  readonly useFirstPageNumber: boolean;
  readonly horizontalDpi: number | null;
  readonly verticalDpi: number | null;
}

/** Worksheet page margins in inches. */
export interface EnginePageMargins {
  readonly left: number;
  readonly right: number;
  readonly top: number;
  readonly bottom: number;
  readonly header: number;
  readonly footer: number;
}

/** Worksheet `<printOptions>` flags. */
export interface EnginePrintOptions {
  readonly gridLines: boolean;
  readonly headings: boolean;
  readonly horizontalCentered: boolean;
  readonly verticalCentered: boolean;
}

/** Worksheet `<headerFooter>` state. Section strings are decoded — the
 *  formatting codes appear plainly (`&C`, `&P`), not XML-escaped. Only the
 *  odd (i.e. every-page) sections are surfaced; the even and first-page
 *  sections round-trip through the engine untouched. */
export interface EngineHeaderFooter {
  readonly oddHeader: string;
  readonly oddFooter: string;
  readonly differentOddEven: boolean;
  readonly differentFirst: boolean;
  readonly scaleWithDoc: boolean;
  readonly alignWithMargins: boolean;
}

/** Manual page breaks on a sheet. Each entry is the 0-based row or column the
 *  break precedes — the same convention the store's page-setup slice uses. */
export interface EnginePageBreaks {
  readonly rows: readonly number[];
  readonly cols: readonly number[];
}

/** Sheet tab visibility, mirroring OOXML `<sheet state>` and the engine's
 *  `SheetVisibility`. Declared here rather than imported so reading a tab's
 *  state does not pull the WASM module glue into a caller's bundle.
 *
 *  `VeryHidden` is the state a workbook uses to keep a settings or lookup
 *  sheet out of reach: a spreadsheet leaves such a sheet out of its "Unhide"
 *  list entirely. */
export const SheetVisibility = {
  Visible: 0,
  Hidden: 1,
  VeryHidden: 2,
} as const;
export type SheetVisibility = (typeof SheetVisibility)[keyof typeof SheetVisibility];

/** Discriminator ordinals for PivotTable filter payload values. */
export const PivotFilterValueKind = {
  None: -1,
  Int: 0,
  Double: 1,
  Text: 2,
} as const;
export type PivotFilterValueKind = (typeof PivotFilterValueKind)[keyof typeof PivotFilterValueKind];

export interface PivotFilterSpec {
  readonly axis: PivotAxis;
  readonly fieldName: string;
  readonly type: PivotFilterType;
  readonly valueKind?: PivotFilterValueKind;
  readonly valueInt?: number;
  readonly valueDouble?: number;
  readonly valueText?: string;
  readonly valueHighKind?: PivotFilterValueKind;
  readonly valueHighInt?: number;
  readonly valueHighDouble?: number;
  readonly valueHighText?: string;
}

/** Value kind ordinals — mirror of `fm_value_kind_t`. We redeclare here to
 *  avoid `const enum` cross-module hazards under `isolatedModules`. */
export const ValueKind = {
  Blank: 0,
  Number: 1,
  Bool: 2,
  Text: 3,
  Error: 4,
  Array: 5,
  Ref: 6,
  Lambda: 7,
} as const;
export type ValueKindT = (typeof ValueKind)[keyof typeof ValueKind];

/** A1-style cell coordinate. Zero-indexed. */
export interface Addr {
  readonly sheet: number;
  readonly row: number;
  readonly col: number;
}

/** Inclusive rectangular range. */
export interface Range {
  readonly sheet: number;
  readonly r0: number;
  readonly c0: number;
  readonly r1: number;
  readonly c1: number;
}

/** Tagged value the UI displays. Mirrors the six error sentinels. */
export type CellValue =
  | { readonly kind: 'blank' }
  | { readonly kind: 'number'; readonly value: number }
  | { readonly kind: 'bool'; readonly value: boolean }
  | { readonly kind: 'text'; readonly value: string }
  | { readonly kind: 'error'; readonly code: number; readonly text: string };

/** Probe-discovered capability flags. Designed for forward-compat: every
 *  field is `readonly boolean`, future fields are added as optional.
 *
 *  Each flag mirrors a specific subset of methods on the Workbook handle.
 *  Probes are conservative: a flag flips on only when *every* method
 *  required for round-tripping that feature is present. */
export interface EngineCapabilities {
  /** `addMerge` + `getMerges` + `removeMerge` + `clearMerges`. Round-trips
   *  through both the store and the engine, including unmerge. */
  readonly merges: boolean;
  /** Full XF-table round-trip: `getCellXfIndex`, `setCellXfIndex`, `getCellXf`,
   *  plus the resolver/dedup writers (`getFont`/`getFill`/`getBorder`/`getNumFmt`,
   *  `addFont`/`addFill`/`addBorder`/`addNumFmt`/`addXf`). The `numFmtId` field
   *  on the XF record carries number-format ids, so a separate `numberFormat`
   *  flag is unnecessary. */
  readonly cellFormatting: boolean;
  /** Cell-level OOXML phonetic guide (`rPh`) read/write. */
  readonly phonetic?: boolean;
  /** Worksheet-level `<autoFilter>` definition read/write. */
  readonly autoFilter?: boolean;
  /** OOXML worksheet-table create/update/delete. */
  readonly tableMutate?: boolean;
  /** `evaluateCfRange` (read-only evaluation). */
  readonly conditionalFormat: boolean;
  /** Full data-validation round-trip: `getValidations` + `addValidation` +
   *  `clearValidations` (plus the implicit `removeValidationAt`). */
  readonly dataValidation: boolean;
  /** `renameSheet` + `removeSheet` + `moveSheet`. */
  readonly sheetMutate: boolean;
  /** `insertRows` + `deleteRows` + `insertCols` + `deleteCols`. */
  readonly insertDeleteRowsCols: boolean;
  /** `setRowHidden` + `setColumnHidden`. */
  readonly hiddenRowsCols: boolean;
  /** `setColumnWidth` + `setRowHeight`. */
  readonly colRowSize: boolean;
  /** `setSheetFreeze`. */
  readonly freeze: boolean;
  /** `getSheetView` readback for zoom, frozen panes, and tab visibility. */
  readonly sheetView?: boolean;
  /** `setSheetZoom`. */
  readonly sheetZoom: boolean;
  /** `setSheetTabHidden` — the two-state view of tab visibility. */
  readonly sheetTabHidden: boolean;
  /** `setSheetVisibility` three-state tab visibility. Required to state
   *  `veryHidden`, which `setSheetTabHidden` can neither set nor clear. */
  readonly sheetVisibility?: boolean;
  /** Read/write display flags on `<sheetView>`: gridlines, row/column
   *  headers, zeros, and right-to-left direction. */
  readonly sheetViewFlags?: boolean;
  /** `setColumnOutline` + `setRowOutline`. */
  readonly outlines: boolean;
  /** `getComment` + `setComment`. */
  readonly comments: boolean;
  /** Optional sheet-wide comment enumerator. Enables comments on blank cells
   *  to hydrate without probing physical cells. */
  readonly commentsEnumerable?: boolean;
  /** Full hyperlink round-trip: `getHyperlinks`, `addHyperlink`, and
   *  `clearHyperlinks`. */
  readonly hyperlinks: boolean;
  /** `setDefinedNameScoped` for workbook and sheet-scoped names. */
  readonly definedNameMutate: boolean;
  /** `setDefinedNameScoped` plus `DefinedNameEntry.localSheetId`. */
  readonly definedNameScopes: boolean;
  /** `setError` static error-value authoring. */
  readonly staticErrorValues: boolean;
  /** `partialRecalc` viewport-scoped recalculation. */
  readonly partialRecalc: boolean;
  /** `setIterativeProgress` callback for cancellable iterative solves. */
  readonly iterativeProgress: boolean;
  /** `getIterative` readback of the stored iterative-calculation settings, so
   *  the dialog can open on what the workbook actually carries instead of on
   *  the engine defaults. */
  readonly iterativeSettings?: boolean;
  /** `spillInfo` returns precise dynamic-array region info per cell. When
   *  off, the renderer falls back to a heuristic that walks right/down
   *  from likely anchor formulas. */
  readonly spillInfo: boolean;
  /** `precedents` + `dependents` graph traversal at the engine level.
   *  Cross-sheet refs are surfaced when this flag is on. */
  readonly traceArrows: boolean;
  /** `functionNames` + `functionMetadata` enumerable function catalog. */
  readonly functionMetadata: boolean;
  /** `localizeFunctionName` + `canonicalizeFunctionName` round-trip. */
  readonly functionLocale: boolean;
  /** `calcMode` + `setCalcMode` round-trip metadata for `<calcPr>`. */
  readonly calcMode: boolean;
  /** Workbook formula-behaviour host profile. */
  readonly spreadsheetProfile: boolean;
  /** `getSheetProtection` + `setSheetProtection` round-trip. */
  readonly sheetProtectionRoundtrip: boolean;
  /** `getExternalLinks` enumeration of `<externalReferences>` records. */
  readonly externalLinks: boolean;
  /** `getLambdaText` rendering of lambda values back to formula text. */
  readonly lambdaText: boolean;
  /** Read-only formula evaluation anchored at a workbook cell. */
  readonly formulaTextEvaluation: boolean;
  /** Read-only ad-hoc evaluation that returns the whole dynamic-array /
   *  spilled result via `evaluateFormulaArray` instead of reducing to the
   *  top-left element. When off, array-returning F9 previews collapse to a
   *  single value. */
  readonly arrayFormulaEvaluation: boolean;
  /** Read-only conditional-format predicate evaluation with a range anchor. */
  readonly conditionalFormulaEvaluation: boolean;
  /** `cellStyleCount` + `getCellStyle` + `getCellStyleXf` named-style
   *  enumeration. */
  readonly cellStyles: boolean;
  /** `addCellStyleXf` + `setCellStyle` named-style authoring. When off, a
   *  style applied from the gallery is saved as direct formatting only. */
  readonly cellStyleMutate?: boolean;
  /** `getConditionalFormats` + `addConditionalFormat` (non-visual) +
   *  `removeConditionalFormatAt` + `clearConditionalFormats` authoring
   *  surface. Read-only `evaluateCfRange` is gated by `conditionalFormat`. */
  readonly conditionalFormatMutate: boolean;
  /** `getDxf` + `addDxf` + `dxfCount` differential-format authoring. */
  readonly conditionalFormatDxf: boolean;
  /** Visual CF payload authoring for color scales, data bars, and icon sets. */
  readonly conditionalFormatVisualMutate: boolean;
  /** `pivotCount` + `pivotLayout` projection of loaded workbook PivotTables. */
  readonly pivotTables: boolean;
  /** PivotCache + PivotTable mutation APIs. Enables low-level PivotTable
   *  authoring; UI wizards can layer on top of `WorkbookHandle` wrappers. */
  readonly pivotTableMutate: boolean;
  /** `pivotFieldAddItemAt` — a manual-filter item addressed by its cache
   *  index rather than its label. Required to express the blank member of a
   *  pivot axis, which has no label to be named by. */
  readonly pivotItemByCacheIndex?: boolean;
  /** Pivot cache worksheet source metadata read/write. */
  readonly pivotCacheSource: boolean;
  /** Pivot compact / tabular / outline report layout read/write. */
  readonly pivotReportLayout: boolean;
  /** Worksheet print settings round-trip: page setup, margins, print options,
   *  header/footer, print area and print titles. Without it the page-setup
   *  slice stays a session-only UI record. */
  readonly printSettings?: boolean;
  /** Raw `<pageSetup>` fragment read/write. The typed page-setup setter models
   *  only what the paginator needs; the attributes beside it are preserved on
   *  a load and save but can be authored only through the fragment. */
  readonly printSettingsXml?: boolean;
  /** Manual page-break read/write on both axes. Separate from `printSettings`
   *  because the break table lives on the worksheet rather than in the print
   *  elements. */
  readonly pageBreaks?: boolean;
}
