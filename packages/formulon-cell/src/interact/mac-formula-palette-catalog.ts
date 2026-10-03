import {
  type FunctionCatalogEntry,
  type FunctionCatalogSnapshot,
  type FunctionCategory,
  isFunctionUnavailableForInsertion,
  supportedFunctionNames,
} from '../commands/function-categories.js';
import type { Strings } from '../i18n/strings.js';

export type MacPaletteStrings = Strings['fxDialog']['macPalette'];

/** Host-supplied help for one function argument. Omitted fields fall back to catalog data. */
export interface FunctionArgumentHelp {
  /** Replaces the catalog argument label. */
  label?: string;
  /** One-line hint rendered under the argument field. */
  description?: string;
  /** Reference page linked from the help section (read from the first argument). */
  url?: string;
}

/** Resolves help for `functionName`'s zero-based argument in the active UI locale. */
export type FunctionArgumentHelpProvider = (
  functionName: string,
  argumentIndex: number,
  locale: string,
) => FunctionArgumentHelp | null | undefined;

/** One picker list section: `key` is `recent`, `all` or a function category. */
export interface PickerSection {
  key: string;
  names: string[];
}

export const paletteStrings = (strings: Strings): MacPaletteStrings => strings.fxDialog.macPalette;

export const sameAvailability = (a: FunctionCatalogEntry, b: FunctionCatalogEntry): boolean =>
  Object.is(a.availability, b.availability);

/** The catalog entry for `name` when it is known and insertable, else null. */
export const insertableEntry = (
  catalog: Pick<FunctionCatalogSnapshot, 'entries' | 'knownNames'>,
  name: string,
): FunctionCatalogEntry | null => {
  const entry = catalog.entries.get(name);
  if (
    !catalog.knownNames.has(name) ||
    !entry ||
    isFunctionUnavailableForInsertion(entry.availability)
  )
    return null;
  return entry;
};

/** Argument help for `index`, merging the host provider over catalog labels. */
export const resolveArgumentHelp = (
  provider: FunctionArgumentHelpProvider | undefined,
  entry: FunctionCatalogEntry,
  index: number,
  locale: string,
  argumentLabel: string,
): FunctionArgumentHelp => {
  const provided = provider?.(entry.canonicalName, index, locale);
  const fallbackLabel = entry.argumentLabels[index] ?? `${argumentLabel} ${index + 1}`;
  return {
    label: provided?.label ?? fallbackLabel.replace(/^\[|\]$/g, ''),
    description: provided?.description,
    url: provided?.url,
  };
};

/** Number of argument fields to show for `entry` given the current argument count. */
export const argumentFieldCount = (entry: FunctionCatalogEntry, argCount: number): number =>
  Math.max(entry.minArity, argCount, entry.argumentLabels.length);

/** Picker sections for `category`, with names narrowed by the search `query`. */
export const pickerSections = (
  category: FunctionCategory,
  catalog: Pick<FunctionCatalogSnapshot, 'names' | 'knownNames'>,
  recentNames: readonly string[],
  query: string,
): PickerSection[] => {
  const needle = query.trim().toUpperCase();
  const section = (key: string, names: readonly string[]): PickerSection => ({
    key,
    names: names.filter((name) => !needle || name.includes(needle)),
  });
  if (category === 'all') {
    return [section('recent', recentNames), section('all', catalog.names)];
  }
  if (category === 'recent') return [section('recent', recentNames)];
  return [section(category, supportedFunctionNames(category, catalog.knownNames))];
};
