import type {
  Pivot as EnginePivot,
  PivotInput as EnginePivotInput,
  PivotRef as EnginePivotRef,
} from '@libraz/formulon';
import type { HandleContext, MemberSpec, SheetContext, WorkbookContext } from './facade-context.js';
import {
  enginePivotCacheRef,
  type PivotCache,
  type PivotCacheRef,
  pivotCacheOf,
} from './pivot-cache.js';
import type { Sheet, Workbook } from './workbook.js';

/**
 * A PivotTable of a sheet. Get one from `Sheet.addPivot`, `Sheet.pivot` or `Sheet.listPivots`; the
 * workbook keeps one `Pivot` per pivot table. Members are the engine `Pivot` members of the same
 * name. Once the pivot is removed every member throws `NotFound`.
 */
// biome-ignore lint/suspicious/noUnsafeDeclarationMerging: the interface below declares the engine members installed on the prototype
export class Pivot {
  /** @internal */
  constructor() {
    throw new TypeError('get a Pivot from Sheet.addPivot() or Sheet.pivot()');
  }
}

/** A pivot table passed by its position on the sheet, its name, or its `Pivot`. */
export type PivotRef = number | string | Pivot;

/** What `Sheet.addPivot` takes; `cacheId` takes a {@link PivotCacheRef}. */
export interface PivotInput extends Omit<EnginePivotInput, 'cacheId'> {
  cacheId: PivotCacheRef;
}

const PIVOT_PROPS = ['index', 'name'] as const;
const PIVOT_READ = [
  'getProperties',
  'getLayout',
  'getGrandTotals',
  'getStyleOptions',
  'getReportLayout',
  'getOptions',
  'isProjectionPending',
  'listFields',
  'getField',
  'listFieldItems',
  'getFieldItem',
  'getFieldDateGroup',
  'getFieldNumericGroup',
  'getRowFieldOrder',
  'getColumnFieldOrder',
  'getPageFieldOrder',
  'listDataFields',
  'getDataField',
  'getDataFieldDefaultName',
  'listFilters',
  'getFilter',
] as const;
const PIVOT_CHANGES = {
  pivot: [
    'rename',
    'setAnchor',
    'updateGrandTotals',
    'updateStyleOptions',
    'setReportLayout',
    'updateOptions',
    'project',
    'updateField',
    'clearFields',
    'setFieldAxis',
    'updateFieldItem',
    'addFieldItem',
    'clearFieldItems',
    'setFieldSubtotalFunctions',
    'setFieldDateGroup',
    'setFieldNumericGroup',
    'setRowFieldOrder',
    'setColumnFieldOrder',
    'setPageFieldOrder',
    'addDataField',
    'setDataField',
    'clearDataFields',
    'addFilter',
    'removeFilter',
    'clearFilters',
  ],
} as const;

export const PIVOT_MEMBERS: MemberSpec<HandleContext<EnginePivot>> = {
  props: PIVOT_PROPS,
  getters: {
    sheet: (ctx) => ctx.sheet(ctx.target.sheet),
    workbook: (ctx) => ctx.workbook,
  },
  read: PIVOT_READ,
  change: PIVOT_CHANGES,
  methods: {
    cache: (ctx: HandleContext<EnginePivot>): PivotCache => pivotCacheOf(ctx, ctx.target.cache()),
    showDetails: (ctx: HandleContext<EnginePivot>, ...at: (number | string)[]): Sheet =>
      ctx.run(() => {
        const show = (pivot: EnginePivot) =>
          (pivot.showDetails as (...args: unknown[]) => ReturnType<EnginePivot['showDetails']>)(
            ...at,
          );
        const { index } = ctx.target;
        const sheetIndex = ctx.target.sheet.index;
        const added = ctx.sheetStep(
          () => show(ctx.target),
          (sheet) => ({ kind: 'sheet-add', index: sheet.index, name: sheet.name }),
          () => show(ctx.engine.sheet(sheetIndex).listPivots()[index] as EnginePivot),
        );
        return ctx.sheet(added);
      }),
  },
};

const pivotOf = (ctx: WorkbookContext, engine: EnginePivot): Pivot => ctx.facade(engine, Pivot);
const enginePivotRef = (ctx: WorkbookContext, ref: PivotRef): EnginePivotRef =>
  ctx.engineOf(ref) as EnginePivotRef;

export const SHEET_PIVOT_MEMBERS: MemberSpec<SheetContext> = {
  methods: {
    listPivots: (ctx: SheetContext): Pivot[] =>
      ctx.sheetEngine.listPivots().map((engine) => pivotOf(ctx, engine)),
    pivot: (ctx: SheetContext, ref: PivotRef): Pivot =>
      pivotOf(ctx, ctx.sheetEngine.pivot(enginePivotRef(ctx, ref))),
    addPivot: (ctx: SheetContext, input: PivotInput): Pivot =>
      ctx.sheetChange('pivot', () =>
        pivotOf(
          ctx,
          ctx.sheetEngine.addPivot({ ...input, cacheId: enginePivotCacheRef(ctx, input.cacheId) }),
        ),
      ),
    removePivot: (ctx: SheetContext, ref: PivotRef): void =>
      ctx.sheetChange('pivot', () => ctx.sheetEngine.removePivot(enginePivotRef(ctx, ref))),
  },
};

type PivotMemberName =
  | (typeof PIVOT_PROPS)[number]
  | (typeof PIVOT_READ)[number]
  | (typeof PIVOT_CHANGES.pivot)[number];

export interface Pivot extends Pick<EnginePivot, PivotMemberName> {
  /** The sheet the pivot is anchored on. */
  readonly sheet: Sheet;
  /** The workbook the pivot belongs to. */
  readonly workbook: Workbook;
  /** The pivot cache the pivot reads. */
  cache(): PivotCache;
  /** Writes the source records behind a value cell to a new sheet and returns that sheet. */
  showDetails(row: number, col: number): Sheet;
  showDetails(a1: string): Sheet;
}

declare module './workbook.js' {
  interface Sheet {
    /** The sheet's pivot tables, in position order. */
    listPivots(): Pivot[];
    /** The pivot at a position, of a name, or behind a `Pivot`; throws `NotFound` when there is none. */
    pivot(ref: PivotRef): Pivot;
    /** Adds a pivot table over `input.cacheId` with its top-left cell at `input.anchor`. */
    addPivot(input: PivotInput): Pivot;
    /** Removes a pivot table. */
    removePivot(ref: PivotRef): void;
  }
}
