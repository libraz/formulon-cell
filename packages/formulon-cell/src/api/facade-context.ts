import type { Sheet as EngineSheet, Workbook as EngineWorkbook } from '@libraz/formulon';
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
}

export interface SheetContext extends WorkbookContext {
  /** The engine sheet handle this facade fronts. */
  readonly sheetEngine: EngineSheet;
  /** Writes the one cell named by `args` (`(row, col, ...)` or `(a1, ...)`) through `write`. */
  cellWrite(args: readonly unknown[], write: () => void): void;
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
}
