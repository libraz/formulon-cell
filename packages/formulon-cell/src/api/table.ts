import type {
  Table as EngineTable,
  TableRef as EngineTableRef,
  Workbook as EngineWorkbook,
  TableInput,
} from '@libraz/formulon';
import type { HandleContext, MemberSpec, SheetContext, WorkbookContext } from './facade-context.js';
import type { Sheet, Workbook } from './workbook.js';

/**
 * A table (ListObject) of a sheet. Get one from `Sheet.addTable`, `Sheet.table`, `Sheet.listTables`
 * or `Workbook.table`; the workbook keeps one `Table` per table. Members are the engine `Table`
 * members of the same name. Once the table is removed every member throws `NotFound`.
 */
// biome-ignore lint/suspicious/noUnsafeDeclarationMerging: the interface below declares the engine members installed on the prototype
export class Table {
  /** @internal */
  constructor() {
    throw new TypeError('get a Table from Sheet.addTable() or Sheet.table()');
  }
}

/** A table passed by its position on the sheet, its name, or its `Table`. */
export type TableRef = number | string | Table;

const TABLE_PROPS = ['index', 'name'] as const;
const TABLE_READ = [
  'getProperties',
  'getStyleOptions',
  'getAutoFilter',
  'evaluateAutoFilter',
] as const;
const TABLE_CHANGES = {
  table: ['update', 'rename', 'renameColumn', 'updateStyleOptions'],
  autoFilter: ['setAutoFilter', 'applyAutoFilter', 'clearAutoFilterCriteria'],
} as const;

export const TABLE_MEMBERS: MemberSpec<HandleContext<EngineTable>> = {
  props: TABLE_PROPS,
  getters: {
    sheet: (ctx) => ctx.sheet(ctx.target.sheet),
    workbook: (ctx) => ctx.workbook,
  },
  read: TABLE_READ,
  change: TABLE_CHANGES,
};

const tableOf = (ctx: WorkbookContext, engine: EngineTable): Table => ctx.facade(engine, Table);

export const SHEET_TABLE_MEMBERS: MemberSpec<SheetContext> = {
  methods: {
    addTable: (ctx: SheetContext, input: TableInput): Table =>
      ctx.sheetChange('table', () => tableOf(ctx, ctx.sheetEngine.addTable(input))),
    table: (ctx: SheetContext, ref: TableRef): Table =>
      tableOf(ctx, ctx.sheetEngine.table(ctx.engineOf(ref) as EngineTableRef)),
    listTables: (ctx: SheetContext): Table[] =>
      ctx.sheetEngine.listTables().map((engine) => tableOf(ctx, engine)),
    removeTable: (ctx: SheetContext, ref: TableRef): void =>
      ctx.sheetChange('table', () =>
        ctx.sheetEngine.removeTable(ctx.engineOf(ref) as EngineTableRef),
      ),
  },
};

const WORKBOOK_TABLE_READ = ['getTableStyle', 'listTableStyles', 'getTableStyleDefaults'] as const;
const WORKBOOK_TABLE_CHANGES = {
  table: ['setTableStyle', 'removeTableStyle', 'updateTableStyleDefaults'],
} as const;

export const WORKBOOK_TABLE_MEMBERS: MemberSpec<WorkbookContext> = {
  read: WORKBOOK_TABLE_READ,
  change: WORKBOOK_TABLE_CHANGES,
  methods: {
    table: (ctx: WorkbookContext, name: string): Table => tableOf(ctx, ctx.engine.table(name)),
  },
};

type TableMemberName =
  | (typeof TABLE_PROPS)[number]
  | (typeof TABLE_READ)[number]
  | (typeof TABLE_CHANGES)[keyof typeof TABLE_CHANGES][number];
type WorkbookTableMemberName =
  | (typeof WORKBOOK_TABLE_READ)[number]
  | (typeof WORKBOOK_TABLE_CHANGES.table)[number];

export interface Table extends Pick<EngineTable, TableMemberName> {
  /** The sheet the table lives on. */
  readonly sheet: Sheet;
  /** The workbook the table belongs to. */
  readonly workbook: Workbook;
}

declare module './workbook.js' {
  interface Sheet {
    /** Adds a table over `input.range` and returns it. */
    addTable(input: TableInput): Table;
    /** The table at a position, of a name, or behind a `Table`; throws `NotFound` when there is none. */
    table(ref: TableRef): Table;
    /** The sheet's tables, in position order. */
    listTables(): Table[];
    /** Removes a table; its cells keep their values. */
    removeTable(ref: TableRef): void;
  }
  interface Workbook extends Pick<EngineWorkbook, WorkbookTableMemberName> {
    /** The table of a name, on any sheet; throws `NotFound` when there is none. */
    table(name: string): Table;
  }
}
