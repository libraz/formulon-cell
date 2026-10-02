import type { SpreadsheetStore } from '../store/store.js';
import { FUNCTION_SIGNATURES } from './refs.js';

const RECENT_FUNCTION_LIMIT = 12;

export type RecentFunctionsListener = (names: readonly string[]) => void;

interface RecentFunctionsState {
  names: string[];
  listeners: Set<RecentFunctionsListener>;
}

const recentFunctionsByStore = new WeakMap<SpreadsheetStore, RecentFunctionsState>();

const stateFor = (store: SpreadsheetStore): RecentFunctionsState => {
  const current = recentFunctionsByStore.get(store);
  if (current) return current;
  const next: RecentFunctionsState = { names: [], listeners: new Set() };
  recentFunctionsByStore.set(store, next);
  return next;
};

/** Return the store's most-recently-inserted function names, newest first. */
export const getRecentFunctions = (
  store: SpreadsheetStore,
  knownNames?: ReadonlySet<string>,
): readonly string[] =>
  Object.freeze(
    knownNames === undefined
      ? stateFor(store).names.filter((name) => Object.hasOwn(FUNCTION_SIGNATURES, name))
      : stateFor(store).names.filter((name) => knownNames.has(name)),
  );

/**
 * Record one successfully used picker function. Unknown or non-canonical
 * names are ignored. Returns true only when the MRU list actually changes.
 */
export const recordRecentFunction = (
  store: SpreadsheetStore,
  name: string,
  knownNames?: ReadonlySet<string>,
): boolean => {
  const recognized =
    knownNames === undefined ? Object.hasOwn(FUNCTION_SIGNATURES, name) : knownNames.has(name);
  if (name !== name.toUpperCase() || !recognized) return false;

  const state = stateFor(store);
  if (state.names[0] === name) return false;
  state.names = [name, ...state.names.filter((entry) => entry !== name)].slice(
    0,
    RECENT_FUNCTION_LIMIT,
  );
  const snapshot = Object.freeze([...state.names]);
  for (const listener of [...state.listeners]) {
    try {
      listener(snapshot);
    } catch {
      // A UI subscriber must not interrupt the operation that recorded usage.
    }
  }
  return true;
};

/** Subscribe to MRU list changes. The returned function removes the listener. */
export const subscribeRecentFunctions = (
  store: SpreadsheetStore,
  listener: RecentFunctionsListener,
): (() => void) => {
  const listeners = stateFor(store).listeners;
  listeners.add(listener);
  return () => listeners.delete(listener);
};
