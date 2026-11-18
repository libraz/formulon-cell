// The formulon package is plain ESM and resolves the .wasm next to
// formulon.js via its own `import.meta.url`; hosts that serve the binary
// elsewhere (a CDN, a sub-path) pass `locateFile` through to Emscripten.
//
// The default formulon entry is single-threaded and needs no cross-origin
// isolation. Initialization errors still reject: the in-memory stub is
// intentionally opt-in for tests and demos.
import createFormulon, { type Formulon, type FormulonOptions } from '@libraz/formulon';
import { engineProfileIdFor } from './spreadsheet-profile.js';
import { createStubModule } from './stub-engine.js';

let cachedWasm: Promise<Formulon> | null = null;
let cachedStub: Promise<Formulon> | null = null;
let loadedWasm: Formulon | null = null;
const stubModules = new WeakSet<Formulon>();

export interface LoadOptions {
  /** Force the JS stub even if the WASM could load. Useful for tests and
   *  explicit demos only; production code should leave this off. */
  preferStub?: boolean;
  /** Called when `preferStub` explicitly selects the stub engine. */
  onFallback?: (reason: string) => void;
  /** Resolve the URL of the engine's `.wasm` file, for hosts that serve it
   *  from a CDN or a sub-path. The engine is loaded once per page, so only
   *  the first load's resolver takes effect; later loads reuse that module. */
  locateFile?: FormulonOptions['locateFile'];
}

export function loadFormulon(opts: LoadOptions = {}): Promise<Formulon> {
  if (opts.preferStub) {
    opts.onFallback?.('preferStub set — using stub engine');
    if (!cachedStub) {
      const module = createStubModule();
      stubModules.add(module);
      cachedStub = Promise.resolve(module);
    }
    return cachedStub;
  }

  if (cachedWasm) return cachedWasm;

  const unavailableReason = wasmUnavailableReason();
  if (unavailableReason) {
    return Promise.reject(new Error(unavailableReason));
  }

  const moduleOptions: FormulonOptions = opts.locateFile ? { locateFile: opts.locateFile } : {};
  const promise: Promise<Formulon> = createFormulon(moduleOptions).catch((reason: unknown) => {
    cachedWasm = null;
    throw new Error(`formulon WASM init failed: ${String(reason)}`);
  });
  cachedWasm = promise;
  promise.then(
    (module) => {
      loadedWasm = module;
    },
    () => {},
  );

  return promise;
}

/**
 * The spreadsheet's name for a function in `locale` (a BCP-47 tag) on the Mac profile, read from
 * the loaded engine module itself so no workbook is needed. Returns `null` until the real engine
 * has loaded, and for a function it has no name for.
 */
export function localizedFunctionNameFromEngine(
  canonicalName: string,
  locale: string,
): string | null {
  if (!loadedWasm) return null;
  try {
    const name = loadedWasm.localizeFunctionName(canonicalName, engineProfileIdFor(locale, 'mac'));
    return name || null;
  } catch {
    return null;
  }
}

export function isUsingStub(module?: Formulon): boolean {
  if (module) return stubModules.has(module);
  return cachedStub !== null && cachedWasm === null;
}

function wasmUnavailableReason(): string | null {
  if (typeof WebAssembly === 'undefined') {
    return 'formulon WASM unavailable: WebAssembly is not supported in this environment';
  }
  return null;
}
