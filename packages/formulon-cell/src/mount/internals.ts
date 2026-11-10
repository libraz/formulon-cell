import type { InteractionController } from '../commands/interaction-controller.js';
import type { SpreadsheetStore } from '../store/store.js';
import type { EngineBinding } from './engine-binding.js';

/** What tests, the demos and e2e probes reach behind a mounted spreadsheet. Exported from no
 *  entry point; import it by path. */
export interface SpreadsheetInternals {
  readonly store: SpreadsheetStore;
  readonly binding: EngineBinding;
  readonly controller: InteractionController;
}

interface InternalsSource {
  readonly store: SpreadsheetStore;
  readonly binding: () => EngineBinding;
  readonly controller: InteractionController;
}

const sources = new WeakMap<object, InternalsSource>();

/** Records the internals behind a mounted instance; `mount` calls it once per instance. */
export const registerSpreadsheetInternals = (instance: object, source: InternalsSource): void => {
  sources.set(instance, source);
};

/** Lets another instance object (the public one wrapping a mount) share `from`'s internals. */
export const shareSpreadsheetInternals = (instance: object, from: object): void => {
  const source = sources.get(from);
  if (!source) throw new TypeError('not a mounted spreadsheet');
  sources.set(instance, source);
};

/** The store, engine binding and interaction controller behind `instance`. */
export function spreadsheetInternals(instance: object): SpreadsheetInternals {
  const source = sources.get(instance);
  if (!source) throw new TypeError('not a mounted spreadsheet');
  return { store: source.store, binding: source.binding(), controller: source.controller };
}
