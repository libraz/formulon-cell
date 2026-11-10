import type {
  CellAddress,
  CellRect,
  Sheet as EngineSheet,
  Workbook as EngineWorkbook,
} from '@libraz/formulon';
import type { FacadeChangeKind, WorkbookHandle } from '../engine/workbook-handle.js';
import type { Sheet, Workbook } from './workbook.js';

/**
 * What a facade member file needs to implement a member: the live engine handles and the write
 * paths that journal, schedule the recalc and announce the change. Member files import this type
 * only, so they carry no runtime dependency on workbook.ts.
 */
export interface WorkbookContext {
  readonly workbook: Workbook;
  readonly handle: WorkbookHandle;
  /** The live engine workbook; throws `disposed` once the facade is disposed. */
  readonly engine: EngineWorkbook;
  /** Runs `fn` as one facade operation: one recalc, one `change` event, one undo step. */
  run<T>(fn: () => T): T;
  /** Runs a non-cell write and announces the store slice it touched. */
  change<T>(kind: FacadeChangeKind, sheet: number | null, fn: () => T): T;
  /** The facade of an engine sheet handle. */
  sheet(engine: EngineSheet): Sheet;
  /** The facade of an engine object handle (a table, a picture, …): one facade per handle. */
  facade<F extends object>(engine: object, cls: FacadeClass<F>): F;
  /** The engine handle behind a facade argument; any other value as it is. */
  engineOf(value: unknown): unknown;
}

/** A facade class: its instances are made by the workbook, never by `new`. */
export type FacadeClass<F> = abstract new (...args: never[]) => F;

/** The context of a facade over an engine object handle. */
export interface HandleContext<H> extends WorkbookContext {
  /** The engine handle the facade fronts. */
  readonly target: H;
  /** The sheet the object lives on; `null` for a workbook-level object. */
  readonly sheetIndex: number | null;
}

export interface SheetContext extends WorkbookContext {
  /** The engine sheet handle this facade fronts. */
  readonly sheetEngine: EngineSheet;
  /** Writes the one cell named by `args` (`(row, col, ...)` or `(a1, ...)`) through `write`. */
  cellWrite(args: readonly unknown[], write: () => void): void;
  /** Writes every cell of `range` through one engine call. */
  rangeWrite<T>(range: CellRect, write: () => T): T;
  /** The cell an A1 address or a `{ row, col }` names, on this sheet. */
  cellAt(at: string | CellAddress): CellAddress;
  /** The rectangle an A1 range or a `CellRect` names, on this sheet. */
  rectOf(range: string | CellRect): CellRect;
  /** `change` on this sheet. */
  sheetChange<T>(kind: FacadeChangeKind, fn: () => T): T;
}

/** A hand-written facade member: receives the context first, then the caller's arguments. */
export type ContextMethod<C> = (ctx: C, ...args: never[]) => unknown;

/**
 * The members one file contributes to a facade class. Pass-through members forward to the engine
 * handle of the same name with the same arguments; `methods` are written by hand.
 */
export interface MemberSpec<C> {
  /** Reads: forwarded as is. */
  readonly read?: readonly string[];
  /** Writes of one cell's content: journaled and announced as a value change. */
  readonly cellWrite?: readonly string[];
  /** Other writes, by the store slice they touch. */
  readonly change?: Readonly<Partial<Record<FacadeChangeKind, readonly string[]>>>;
  /** Hand-written members. */
  readonly methods?: Readonly<Record<string, ContextMethod<C>>>;
  /** Engine properties forwarded as read-only getters. */
  readonly props?: readonly string[];
  /** Hand-written read-only getters. */
  readonly getters?: Readonly<Record<string, (ctx: C) => unknown>>;
}
