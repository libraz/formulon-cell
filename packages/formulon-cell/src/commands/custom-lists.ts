import type { SpreadsheetStore } from '../store/store.js';
import { customListsFor } from './autofill-series.js';

/** The user's custom lists of a mount: read from the host's option, and written back when the
 *  Custom Lists dialog adds or deletes one. The built-in day and month lists are not among them. */
export interface CustomListsSource {
  get(): readonly (readonly string[])[];
  set(next: readonly (readonly string[])[]): void;
}

const sources = new WeakMap<SpreadsheetStore, CustomListsSource>();

export function registerCustomLists(store: SpreadsheetStore, source: CustomListsSource): void {
  sources.set(store, source);
}

export function customListsSourceFor(store: SpreadsheetStore): CustomListsSource | undefined {
  return sources.get(store);
}

/** The user's lists of `store`'s mount; none without a mount. */
export const userCustomLists = (
  store: SpreadsheetStore | undefined,
): readonly (readonly string[])[] => (store ? sources.get(store)?.get() : undefined) ?? [];

/** Every list AutoFill and custom-list sorting use: the locale's built-in lists, then the user's,
 *  so a user list wins where both hold the same entry. */
export const allCustomLists = (
  store: SpreadsheetStore | undefined,
  locale: string | undefined,
): readonly (readonly string[])[] => [...customListsFor(locale), ...userCustomLists(store)];
