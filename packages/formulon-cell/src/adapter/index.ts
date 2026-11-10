/**
 * Helpers for framework adapters (React, Vue, ...) that bind component props to a
 * `SpreadsheetInstance`. Application code that mounts a spreadsheet directly does not need them.
 *
 * @module
 */
import type { SpreadsheetInstance } from '../api/spreadsheet/instance.js';
import type { SpreadsheetPatch } from '../api/spreadsheet/types.js';

type PatchKey = keyof SpreadsheetPatch;

const isPlainObject = (value: unknown): value is Record<string, unknown> => {
  if (typeof value !== 'object' || value === null) return false;
  const proto = Object.getPrototypeOf(value);
  return proto === Object.prototype || proto === null;
};

const sameDeep = (a: unknown, b: unknown, seen: WeakMap<object, WeakSet<object>>): boolean => {
  if (Object.is(a, b)) return true;
  const bothArrays = Array.isArray(a) && Array.isArray(b);
  const bothPlain = isPlainObject(a) && isPlainObject(b);
  if (!bothArrays && !bothPlain) return false;
  const left = a as object;
  const right = b as object;
  // A pair already under comparison is assumed equal; the outer walk decides.
  let partners = seen.get(left);
  if (partners?.has(right)) return true;
  if (!partners) {
    partners = new WeakSet();
    seen.set(left, partners);
  }
  partners.add(right);
  if (bothArrays) {
    const x = a as readonly unknown[];
    const y = b as readonly unknown[];
    return x.length === y.length && x.every((item, i) => sameDeep(item, y[i], seen));
  }
  const x = a as Record<string, unknown>;
  const y = b as Record<string, unknown>;
  const keys = Object.keys(x);
  if (keys.length !== Object.keys(y).length) return false;
  return keys.every((key) => Object.hasOwn(y, key) && sameDeep(x[key], y[key], seen));
};

/**
 * Structural equality for adapter props: primitives by `Object.is`, arrays element-wise, plain
 * objects by own enumerable keys; everything else (functions, class instances such as a
 * `Workbook`, `AbortSignal`, DOM nodes, `Map`, `Set`) by identity. Cyclic structures terminate.
 *
 * @example
 * ```ts
 * sameProp({ a: [1, 2] }, { a: [1, 2] }); // true
 * sameProp(new Map(), new Map()); // false
 * ```
 */
export const sameProp = (a: unknown, b: unknown): boolean => sameDeep(a, b, new WeakMap());

/**
 * The `update` patch that takes an instance from `previous` props to `next`: every key whose value
 * differs by {@link sameProp}. A key that was set and is now `undefined` or absent maps to `null`
 * (reset to its default); unchanged keys are omitted.
 *
 * @example
 * ```ts
 * propsPatch({ theme: 'ink', locale: 'en' }, { theme: 'paper', locale: 'en' }); // { theme: 'paper' }
 * propsPatch({ theme: 'ink' }, {}); // { theme: null }
 * ```
 */
export const propsPatch = (
  previous: Partial<SpreadsheetPatch>,
  next: Partial<SpreadsheetPatch>,
): SpreadsheetPatch => {
  const patch: Record<string, unknown> = {};
  const keys = new Set<string>([...Object.keys(previous), ...Object.keys(next)]);
  for (const key of keys) {
    const before = (previous as Record<string, unknown>)[key];
    const after = (next as Record<string, unknown>)[key];
    if (sameProp(before, after)) continue;
    if (after === undefined) {
      if (before !== undefined) patch[key] = null;
      continue;
    }
    patch[key] = after;
  }
  return patch as SpreadsheetPatch;
};

/**
 * Applies the difference between two prop sets to an instance, one `update` call per key so that
 * a key the instance refuses does not hold back the others. Refusals go to `onError`; without it
 * the first one is rethrown once the remaining keys have been applied.
 *
 * @returns The keys that were applied.
 * @example
 * ```ts
 * reconcileProps(instance, { theme: 'ink' }, { theme: 'paper' }, (error, key) => report(key, error));
 * ```
 */
export const reconcileProps = (
  instance: SpreadsheetInstance,
  previous: Partial<SpreadsheetPatch>,
  next: Partial<SpreadsheetPatch>,
  onError?: (error: unknown, key: PatchKey) => void,
): readonly PatchKey[] => {
  const patch = propsPatch(previous, next);
  const applied: PatchKey[] = [];
  let first: { error: unknown } | null = null;
  for (const key of Object.keys(patch) as PatchKey[]) {
    try {
      instance.update({ [key]: patch[key] } as SpreadsheetPatch);
      applied.push(key);
    } catch (error) {
      if (onError) onError(error, key);
      else first ??= { error };
    }
  }
  if (first) throw first.error;
  return applied;
};
