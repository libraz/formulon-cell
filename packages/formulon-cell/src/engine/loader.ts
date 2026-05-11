// The formulon package is plain ESM and resolves the .wasm itself via its
// own `import.meta.url` + `locateFile` mechanism, so we don't pass an
// explicit wasm URL here — Emscripten finds the file next to formulon.js
// at the consumer's resolved path.
//
// formulon's WASM uses pthread/SharedArrayBuffer. Browsers without a
// crossOriginIsolated context (missing COOP+COEP, ad-hoc demos, SSR shells)
// will fail at instantiation; in that case we fall back to an in-memory
// `stub` engine so the UI keeps working — formulas degrade gracefully.
import createFormulon from '@libraz/formulon';
import { createStubModule } from './stub-engine.js';
import type { FormulonModule } from './types.js';

let cached: Promise<FormulonModule> | null = null;
let usedStub = false;

export interface LoadOptions {
  /** Force the JS stub even if the WASM could load. Useful for tests. */
  preferStub?: boolean;
  /** Called when the engine falls back to the stub (e.g. no SAB). */
  onFallback?: (reason: string) => void;
}

export function loadFormulon(opts: LoadOptions = {}): Promise<FormulonModule> {
  if (cached) return cached;

  if (opts.preferStub || !canLoadWasm()) {
    usedStub = true;
    opts.onFallback?.(
      opts.preferStub
        ? 'preferStub set — using stub engine'
        : 'crossOriginIsolated unavailable — using stub engine',
    );
    cached = Promise.resolve(createStubModule());
    return cached;
  }

  const promise: Promise<FormulonModule> = createFormulon().catch((reason: unknown) => {
    usedStub = true;
    opts.onFallback?.(`WASM init failed: ${String(reason)}`);
    return createStubModule();
  });
  cached = promise;

  return promise;
}

export function isUsingStub(): boolean {
  return usedStub;
}

function canLoadWasm(): boolean {
  if (typeof WebAssembly === 'undefined') return false;
  // pthreaded WASM needs SAB. If we're in a browser context that doesn't
  // expose it, save the round-trip and use the stub up front.
  if (typeof SharedArrayBuffer === 'undefined') return false;
  return true;
}
