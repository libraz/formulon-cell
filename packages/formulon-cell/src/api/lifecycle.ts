import { spreadsheetFailure } from './errors.js';
import { reportUncaught } from './events.js';

/** The members every disposable public object carries. */
export interface Disposable {
  /** True once `dispose()` has run. */
  readonly disposed: boolean;
  /** Releases the object. Idempotent; never throws. */
  dispose(): void;
  [Symbol.dispose](): void;
}

/**
 * Disposal state for one public object: guards members after disposal and runs teardown steps once,
 * newest first, reporting a failing step instead of throwing.
 */
export class Lifecycle {
  private isDisposed = false;
  private readonly teardown: Array<() => void> = [];

  constructor(private readonly owner: string) {}

  get disposed(): boolean {
    return this.isDisposed;
  }

  /** Throws `FormulonError(BindingInvalidHandle)` with the `disposed` failure once disposed. */
  assertLive(): void {
    if (this.isDisposed) throw spreadsheetFailure('disposed', `${this.owner} is disposed`);
  }

  /** Registers a teardown step; after disposal the step runs at once. */
  onDispose(step: () => void): void {
    if (this.isDisposed) {
      runStep(step);
      return;
    }
    this.teardown.push(step);
  }

  /** Marks the object disposed and runs the teardown steps. Idempotent. */
  dispose(): void {
    if (this.isDisposed) return;
    this.isDisposed = true;
    for (let i = this.teardown.length - 1; i >= 0; i -= 1) {
      const step = this.teardown[i];
      if (step) runStep(step);
    }
    this.teardown.length = 0;
  }
}

const runStep = (step: () => void): void => {
  try {
    step();
  } catch (error) {
    reportUncaught(error);
  }
};
