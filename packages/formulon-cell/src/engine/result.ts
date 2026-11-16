import type { FormulonError, StatusCode } from '@libraz/formulon';

/** True for an error the engine threw. Matched by name so errors from the
 *  stub engine and from any copy of the engine package are recognised. */
export function isFormulonError(e: unknown): e is FormulonError {
  return e instanceof Error && e.name === 'FormulonError';
}

/** The engine status code of `e`, or null when `e` is not an engine error. */
export function formulonErrorCode(e: unknown): StatusCode | null {
  return isFormulonError(e) ? e.code : null;
}

/** `call()`, or `fallback` when the engine refuses it; any other error propagates. */
export function orFallback<T, F>(fallback: F, call: () => T): T | F {
  try {
    return call();
  } catch (e) {
    if (!isFormulonError(e)) throw e;
    return fallback;
  }
}

/** `write()` as `true`, or `false` when the engine refuses it; any other error propagates. */
export function attemptWrite(write: () => void): boolean {
  return orFallback(false, () => {
    write();
    return true;
  });
}
