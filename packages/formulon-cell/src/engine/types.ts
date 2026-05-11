// Re-export of the formulon-typed surface plus our adapter shapes.
export type {
  BorderRecord,
  BorderSide,
  CellEntry,
  CellResult,
  CellXf,
  DataValidationEntry,
  DataValidationInput,
  DataValidationRange,
  EvalResult,
  FillRecord,
  FontRecord,
  FormulonModule,
  PivotCell,
  PivotLayoutResult,
  SaveResult,
  Status,
  StringResult,
  Value,
  Workbook,
} from '@libraz/formulon';

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
  /** `setSheetZoom`. */
  readonly sheetZoom: boolean;
  /** `setSheetTabHidden`. */
  readonly sheetTabHidden: boolean;
  /** `setColumnOutline` + `setRowOutline`. */
  readonly outlines: boolean;
  /** `getComment` + `setComment`. */
  readonly comments: boolean;
  /** Full hyperlink round-trip: `getHyperlinks`, `addHyperlink`, and
   *  `clearHyperlinks`. */
  readonly hyperlinks: boolean;
  /** `setDefinedName` (no remove yet — pass empty formula to clear via
   *  the engine convention). */
  readonly definedNameMutate: boolean;
  /** `partialRecalc` viewport-scoped recalculation. */
  readonly partialRecalc: boolean;
  /** `setIterativeProgress` callback for cancellable iterative solves. */
  readonly iterativeProgress: boolean;
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
  /** `cellStyleCount` + `getCellStyle` + `getCellStyleXf` named-style
   *  enumeration. */
  readonly cellStyles: boolean;
  /** `getConditionalFormats` + `addConditionalFormat` (non-visual) +
   *  `removeConditionalFormatAt` + `clearConditionalFormats` authoring
   *  surface. Read-only `evaluateCfRange` is gated by `conditionalFormat`. */
  readonly conditionalFormatMutate: boolean;
  /** `pivotCount` + `pivotLayout` projection of loaded workbook PivotTables. */
  readonly pivotTables: boolean;
  /** PivotCache + PivotTable mutation APIs. Enables low-level PivotTable
   *  authoring; UI wizards can layer on top of `WorkbookHandle` wrappers. */
  readonly pivotTableMutate: boolean;
}
