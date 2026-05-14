// The formulon package is plain ESM and resolves the .wasm itself via its
// own `import.meta.url` + `locateFile` mechanism, so we don't pass an
// explicit wasm URL here — Emscripten finds the file next to formulon.js
// at the consumer's resolved path.
//
// formulon's WASM uses pthread/SharedArrayBuffer. Browsers without a
// crossOriginIsolated context (missing COOP+COEP, ad-hoc demos, SSR shells)
// will fail at instantiation. Treat that as a host configuration error by
// default: the in-memory stub is intentionally opt-in for tests and demos,
// because silently pretending to calculate can corrupt user expectations.
import createFormulon from '@libraz/formulon';
import { createStubModule } from './stub-engine.js';
import type { FormulonModule } from './types.js';

let cached: Promise<FormulonModule> | null = null;
let usedStub = false;

export interface LoadOptions {
  /** Force the JS stub even if the WASM could load. Useful for tests and
   *  explicit demos only; production code should leave this off. */
  preferStub?: boolean;
  /** Called when `preferStub` explicitly selects the stub engine. */
  onFallback?: (reason: string) => void;
}

export function loadFormulon(opts: LoadOptions = {}): Promise<FormulonModule> {
  if (cached) return cached;

  if (opts.preferStub) {
    usedStub = true;
    opts.onFallback?.('preferStub set — using stub engine');
    cached = Promise.resolve(createStubModule());
    return cached;
  }

  const unavailableReason = wasmUnavailableReason();
  if (unavailableReason) {
    return Promise.reject(new Error(unavailableReason));
  }

  const promise: Promise<FormulonModule> = createFormulon().catch((reason: unknown) => {
    throw new Error(`formulon WASM init failed: ${String(reason)}`);
  });
  cached = promise;

  return promise;
}

export function isUsingStub(): boolean {
  return usedStub;
}

function wasmUnavailableReason(): string | null {
  if (typeof WebAssembly === 'undefined') {
    return 'formulon WASM unavailable: WebAssembly is not supported in this environment';
  }
  // pthreaded WASM needs SAB. If we're in a browser context that doesn't
  // expose it, fail before invoking the Emscripten loader so callers see the
  // missing COOP/COEP setup instead of a partial spreadsheet.
  if (typeof SharedArrayBuffer === 'undefined') {
    return 'formulon WASM unavailable: SharedArrayBuffer is missing; serve the page with COOP: same-origin and COEP: require-corp, or pass preferStub: true explicitly for tests/demos';
  }
  return null;
}
