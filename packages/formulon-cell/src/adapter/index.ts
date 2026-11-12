/**
 * Helpers for framework adapters (React, Vue, ...) that bind component props to a
 * `SpreadsheetInstance`. Application code that mounts a spreadsheet directly does not need them.
 *
 * @module
 */
import type { Disposer } from '../api/events.js';
import type { SpreadsheetInstance } from '../api/spreadsheet/instance.js';
import type { SpreadsheetOptions, SpreadsheetPatch } from '../api/spreadsheet/types.js';
import type { RibbonHooks } from '../toolbar/ribbon/apply-ribbon-command.js';

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

/**
 * Every `SpreadsheetOptions` key, in the order an adapter component reads its props.
 * Compilation fails when `SpreadsheetOptions` gains a key this list lacks.
 */
export const SPREADSHEET_OPTION_KEYS = [
  'workbook',
  'seed',
  'ui',
  'policy',
  'viewport',
  'contextMenu',
  'overlays',
  'toolbar',
  'theme',
  'locale',
  'strings',
  'extensions',
  'formulaHints',
  'spelling',
  'autoCorrect',
  'host',
  'uploadStatus',
  'macroRecording',
  'showAllNotes',
  'signal',
] as const satisfies readonly (keyof SpreadsheetOptions)[];

type MissingOptionKey = Exclude<keyof SpreadsheetOptions, (typeof SPREADSHEET_OPTION_KEYS)[number]>;
// Fails to compile when SpreadsheetOptions gains a key the list above lacks.
const optionKeysComplete: [MissingOptionKey] extends [never] ? true : MissingOptionKey = true;
void optionKeysComplete;

/**
 * Picks the mount options out of a component's props.
 *
 * @param props - The component's props, which may carry keys that are not options.
 * @returns The props named in {@link SPREADSHEET_OPTION_KEYS}, leaving out absent and `undefined` ones.
 */
export const optionsFromProps = (props: SpreadsheetOptions): SpreadsheetOptions => {
  const options: Record<string, unknown> = {};
  for (const key of SPREADSHEET_OPTION_KEYS)
    if (props[key] !== undefined) options[key] = props[key];
  return options as SpreadsheetOptions;
};

/**
 * Narrows mount options to what `update` accepts.
 *
 * @param options - Mount options, as {@link optionsFromProps} returns them.
 * @returns The options without `seed` and `signal`, which only apply at mount.
 */
export const patchFromOptions = (options: SpreadsheetOptions): Partial<SpreadsheetPatch> => {
  const { seed: _seed, signal: _signal, ...patch } = options;
  return patch;
};

type HookGroups = Record<string, Record<string, unknown> | undefined>;

/**
 * The remount key of a ribbon's hooks: the ribbon remounts only when this key changes.
 *
 * @param hooks - The ribbon hooks a component received.
 * @returns The `group.name` of every hook present, sorted and comma-joined; `''` for none.
 */
export const ribbonHooksKey = (hooks: RibbonHooks | undefined): string =>
  Object.entries((hooks ?? {}) as HookGroups)
    .flatMap(([group, members]) =>
      Object.entries(members ?? {})
        .filter(([, fn]) => typeof fn === 'function')
        .map(([name]) => `${group}.${name}`),
    )
    .sort()
    .join(',');

/**
 * Stable ribbon hooks that call whatever hooks a component holds when a command runs.
 *
 * @param get - Reads the component's current hooks.
 * @returns Hooks with the same `group.name` members as `get()` has now, each forwarding to the
 *   member `get()` returns at call time.
 */
export const forwardRibbonHooks = (get: () => RibbonHooks | undefined): RibbonHooks => {
  const out: Record<string, Record<string, unknown>> = {};
  for (const key of ribbonHooksKey(get()).split(',')) {
    if (!key) continue;
    const [group, name] = key.split('.') as [string, string];
    out[group] ??= {};
    out[group][name] = (...args: unknown[]) => {
      const fn = (get() as HookGroups | undefined)?.[group]?.[name];
      return typeof fn === 'function' ? fn(...args) : undefined;
    };
  }
  return out as RibbonHooks;
};

/**
 * Listens to a spreadsheet for the events a derived value depends on.
 *
 * @param spreadsheet - The mounted spreadsheet.
 * @param notify - Called after each event of interest.
 * @returns A disposer that stops listening.
 */
export type SpreadsheetSubscriber = (
  spreadsheet: SpreadsheetInstance,
  notify: () => void,
) => Disposer;

/**
 * Whether a derived value can be read from a spreadsheet.
 *
 * @param spreadsheet - The spreadsheet, or `null` before it mounts.
 * @returns `true` when it is mounted and not disposed.
 */
export const isLiveSpreadsheet = (
  spreadsheet: SpreadsheetInstance | null,
): spreadsheet is SpreadsheetInstance => spreadsheet !== null && !spreadsheet.disposed;

/**
 * A {@link SpreadsheetSubscriber} that notifies after every change, recalc, workbook swap,
 * `update`, selection move and locale change.
 */
export const subscribeToSpreadsheet: SpreadsheetSubscriber = (s, notify) => {
  const offs = [
    s.on('change', notify),
    s.on('cell', notify),
    s.on('workbook', notify),
    s.on('recalc', notify),
    s.on('update', notify),
    s.selection.on('change', notify),
    s.i18n.on('change', notify),
  ];
  return () => {
    for (const off of offs) off();
  };
};

/** A {@link SpreadsheetSubscriber} that notifies when the selection moves. */
export const subscribeToSelection: SpreadsheetSubscriber = (s, notify) =>
  s.selection.on('change', notify);

/** A {@link SpreadsheetSubscriber} that notifies when the locale or its strings change. */
export const subscribeToLocale: SpreadsheetSubscriber = (s, notify) => s.i18n.on('change', notify);
