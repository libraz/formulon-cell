import type {
  PivotCache as EnginePivotCache,
  PivotCacheRef as EnginePivotCacheRef,
  PivotWorksheetSourceInput as EnginePivotWorksheetSourceInput,
  Sheet as EngineSheet,
} from '@libraz/formulon';
import type { HandleContext, MemberSpec, WorkbookContext } from './facade-context.js';
import type { SheetRef, Workbook } from './workbook.js';

/**
 * A pivot cache of a workbook: the source and records one or more pivot tables read. Get one from
 * `Workbook.addPivotCache`, `Workbook.pivotCache`, `Workbook.listPivotCaches` or `Pivot.cache`; the
 * workbook keeps one `PivotCache` per cache. Members are the engine `PivotCache` members of the same
 * name. Once the cache is removed every member throws `NotFound`.
 */
// biome-ignore lint/suspicious/noUnsafeDeclarationMerging: the interface below declares the engine members installed on the prototype
export class PivotCache {
  /** @internal */
  constructor() {
    throw new TypeError('get a PivotCache from Workbook.addPivotCache() or Workbook.pivotCache()');
  }
}

/** A pivot cache passed by its `cacheId` or its `PivotCache`. */
export type PivotCacheRef = number | PivotCache;

/** The worksheet source of a pivot cache; `sheet` takes a {@link SheetRef}. */
export interface PivotWorksheetSourceInput extends Omit<EnginePivotWorksheetSourceInput, 'sheet'> {
  sheet?: SheetRef;
}

const engineSource = (
  ctx: WorkbookContext,
  { sheet, ...rest }: PivotWorksheetSourceInput,
): EnginePivotWorksheetSourceInput =>
  sheet === undefined
    ? rest
    : { ...rest, sheet: ctx.engineOf(sheet) as EngineSheet | number | string };

const PIVOT_CACHE_PROPS = ['cacheId'] as const;
const PIVOT_CACHE_READ = [
  'getWorksheetSource',
  'getOptions',
  'listFields',
  'listCalculatedItems',
  'listFieldSharedItems',
  'getRecordCount',
  'getRecordValue',
] as const;
const PIVOT_CACHE_CHANGES = {
  pivot: [
    'refresh',
    'updateOptions',
    'addField',
    'clearFields',
    'addCalculatedField',
    'setCalculatedFieldFormula',
    'removeCalculatedField',
    'addCalculatedItem',
    'setCalculatedItemFormula',
    'moveCalculatedItem',
    'removeCalculatedItem',
    'addFieldSharedItem',
    'clearFieldSharedItems',
    'addRecord',
    'clearRecords',
    'setRecordValue',
  ],
} as const;

export const PIVOT_CACHE_MEMBERS: MemberSpec<HandleContext<EnginePivotCache>> = {
  props: PIVOT_CACHE_PROPS,
  getters: { workbook: (ctx) => ctx.workbook },
  read: PIVOT_CACHE_READ,
  change: PIVOT_CACHE_CHANGES,
  methods: {
    setWorksheetSource: (
      ctx: HandleContext<EnginePivotCache>,
      source: PivotWorksheetSourceInput | null,
    ): void =>
      ctx.change('pivot', null, () =>
        ctx.target.setWorksheetSource(source && engineSource(ctx, source)),
      ),
  },
};

/** The facade of an engine pivot cache handle. */
export const pivotCacheOf = (ctx: WorkbookContext, engine: EnginePivotCache): PivotCache =>
  ctx.facade(engine, PivotCache);

/** The engine reference behind a {@link PivotCacheRef}. */
export const enginePivotCacheRef = (
  ctx: WorkbookContext,
  ref: PivotCacheRef,
): EnginePivotCacheRef => ctx.engineOf(ref) as EnginePivotCacheRef;

export const WORKBOOK_PIVOT_CACHE_MEMBERS: MemberSpec<WorkbookContext> = {
  methods: {
    listPivotCaches: (ctx: WorkbookContext): PivotCache[] =>
      ctx.engine.listPivotCaches().map((engine) => pivotCacheOf(ctx, engine)),
    pivotCache: (ctx: WorkbookContext, ref: PivotCacheRef): PivotCache =>
      pivotCacheOf(ctx, ctx.engine.pivotCache(enginePivotCacheRef(ctx, ref))),
    addPivotCache: (ctx: WorkbookContext, requestedId?: number | null): PivotCache =>
      ctx.change('pivot', null, () => pivotCacheOf(ctx, ctx.engine.addPivotCache(requestedId))),
    removePivotCache: (ctx: WorkbookContext, ref: PivotCacheRef): void =>
      ctx.change('pivot', null, () => ctx.engine.removePivotCache(enginePivotCacheRef(ctx, ref))),
  },
};

type PivotCacheMemberName =
  | (typeof PIVOT_CACHE_PROPS)[number]
  | (typeof PIVOT_CACHE_READ)[number]
  | (typeof PIVOT_CACHE_CHANGES.pivot)[number];

export interface PivotCache extends Pick<EnginePivotCache, PivotCacheMemberName> {
  /** The workbook the cache belongs to. */
  readonly workbook: Workbook;
  /** Sets the worksheet range (or defined name) the cache reads; `null` removes it. */
  setWorksheetSource(source: PivotWorksheetSourceInput | null): void;
}

declare module './workbook.js' {
  interface Workbook {
    /** The workbook's pivot caches. */
    listPivotCaches(): PivotCache[];
    /** The pivot cache of an id or behind a `PivotCache`; throws `NotFound` when there is none. */
    pivotCache(ref: PivotCacheRef): PivotCache;
    /** Adds an empty pivot cache; give it a source with `setWorksheetSource` before a pivot uses it. */
    addPivotCache(requestedId?: number | null): PivotCache;
    /** Removes a pivot cache; throws `PreconditionFailed` while a pivot still uses it. */
    removePivotCache(ref: PivotCacheRef): void;
  }
}
