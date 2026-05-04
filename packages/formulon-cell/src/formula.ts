// Custom-function registry — host-side Excel-like functions consumers can
// add at runtime. The formulon engine does not currently expose a
// callback-based user-function surface, so v0.1 keeps the registry on the
// JS side: the spreadsheet uses it for autocomplete suggestions and
// signature hints, and exposes `evaluate(name, args)` so application code
// can wire its own derived-cell flow via `inst.on('cellChange', …)`.
//
// When the engine grows native user-function support the registry
// becomes the single source of truth bridging both ends — the public API
// won't change.
import type { CellValue } from './engine/types.js';

export type CustomFunctionReturn = CellValue | number | string | boolean | null;

export interface CustomFunctionMeta {
  /** Short description shown in tooltips / docs. */
  description?: string;
  /** Argument labels — used by the autocomplete signature popover. Mark
   *  optional args with surrounding brackets: `'[step]'`. */
  args?: readonly string[];
  /** Optional return-type hint. Doesn't affect runtime. */
  returnType?: 'number' | 'text' | 'bool' | 'any';
}

export interface CustomFunction {
  /** Upper-cased function name. Stored case-insensitively. */
  readonly name: string;
  readonly meta: CustomFunctionMeta;
  readonly impl: (...args: CellValue[]) => CustomFunctionReturn;
}

const normalizeReturn = (raw: CustomFunctionReturn): CellValue => {
  if (raw === null || raw === undefined) return { kind: 'blank' };
  if (typeof raw === 'number') return { kind: 'number', value: raw };
  if (typeof raw === 'string') return { kind: 'text', value: raw };
  if (typeof raw === 'boolean') return { kind: 'bool', value: raw };
  return raw;
};

export class FormulaRegistry {
  private readonly entries = new Map<string, CustomFunction>();
  private readonly listeners = new Set<() => void>();

  /** Register or replace a function under `name`. Returns a disposer that
   *  removes only this exact registration. */
  register(name: string, impl: CustomFunction['impl'], meta: CustomFunctionMeta = {}): () => void {
    const upper = name.toUpperCase();
    if (!/^[A-Z][A-Z0-9_]*$/.test(upper)) {
      throw new Error(`formulon-cell: invalid function name "${name}"`);
    }
    const entry: CustomFunction = { name: upper, meta, impl };
    this.entries.set(upper, entry);
    this.notify();
    return () => {
      // Only remove if we still own the slot — protects against leaking
      // a previous disposer after a later register replaces the entry.
      if (this.entries.get(upper) === entry) {
        this.entries.delete(upper);
        this.notify();
      }
    };
  }

  unregister(name: string): boolean {
    const ok = this.entries.delete(name.toUpperCase());
    if (ok) this.notify();
    return ok;
  }

  has(name: string): boolean {
    return this.entries.has(name.toUpperCase());
  }

  get(name: string): CustomFunction | undefined {
    return this.entries.get(name.toUpperCase());
  }

  /** Sorted, upper-cased list of registered names. Stable for autocomplete. */
  list(): string[] {
    return [...this.entries.keys()].sort();
  }

  /** Synchronously invoke `name` with `args`. Throws when the function is
   *  unknown. Returns the impl's value coerced into a `CellValue`. */
  evaluate(name: string, args: readonly CellValue[]): CellValue {
    const fn = this.get(name);
    if (!fn) throw new Error(`formulon-cell: unknown custom function "${name}"`);
    return normalizeReturn(fn.impl(...args));
  }

  /** Subscribe to registry changes (register / unregister). Returns
   *  unsubscribe. Used by autocomplete to keep its menu fresh. */
  subscribe(fn: () => void): () => void {
    this.listeners.add(fn);
    return () => {
      this.listeners.delete(fn);
    };
  }

  /** Drop every registered function. Mostly useful for tests. */
  clear(): void {
    if (this.entries.size === 0) return;
    this.entries.clear();
    this.notify();
  }

  private notify(): void {
    for (const fn of [...this.listeners]) {
      try {
        fn();
      } catch (err) {
        console.error('formulon-cell: formula-registry listener threw', err);
      }
    }
  }
}
