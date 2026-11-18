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

/** Outcome of an engine write: `status` is the engine code, or null when no engine call was made. */
export type WriteResult = { ok: true } | { ok: false; status: StatusCode | null };

/** `write()` as `{ ok: true }`, or the engine status when it refuses; any other error propagates. */
export function attemptWriteStatus(write: () => void): WriteResult {
  try {
    write();
    return { ok: true };
  } catch (e) {
    if (!isFormulonError(e)) throw e;
    return { ok: false, status: e.code };
  }
}

/** `write()` as `true`, or `false` when the engine refuses it; any other error propagates. */
export function attemptWrite(write: () => void): boolean {
  return attemptWriteStatus(write).ok;
}
