// Re-export of the formulon-typed surface plus our adapter shapes.
// Vendored during pre-publish; once @libraz/formulon is on npm this points there.
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
  SaveResult,
  Status,
  StringResult,
  Value,
  Workbook,
} from '../../vendor/formulon/formulon.js';

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

/** Tagged value the UI displays. Mirrors Excel's six error sentinels. */
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
}
