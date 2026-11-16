import type { Sheet } from '@libraz/formulon';
import type { Workbook } from './types.js';

/** The members of `WorkbookHandle` the feature mixins reach into. A mixin that
 *  needs more declares `WorkbookHandleInternals & { ... }`. */
export type WorkbookHandleInternals = {
  wb: Workbook;
  assertAlive(): void;
};

/** `handle` viewed through its internal members. */
export function internals(handle: unknown): WorkbookHandleInternals {
  return handle as WorkbookHandleInternals;
}

export function assertAlive(handle: unknown): void {
  internals(handle).assertAlive();
}

export function wb(handle: unknown): Workbook {
  return internals(handle).wb;
}

export function ws(handle: unknown, sheet: number): Sheet {
  return wb(handle).sheet(sheet);
}
