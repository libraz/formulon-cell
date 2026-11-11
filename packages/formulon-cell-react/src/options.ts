import type { SpreadsheetOptions, SpreadsheetPatch } from '@libraz/formulon-cell';

/** Every `SpreadsheetOptions` key; the component reads its props through this list. */
export const OPTION_KEYS = [
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
  'host',
  'uploadStatus',
  'macroRecording',
  'showAllNotes',
  'signal',
] as const satisfies readonly (keyof SpreadsheetOptions)[];

type MissingOptionKey = Exclude<keyof SpreadsheetOptions, (typeof OPTION_KEYS)[number]>;
// Fails to compile when SpreadsheetOptions gains a key the list above lacks.
const optionKeysComplete: [MissingOptionKey] extends [never] ? true : MissingOptionKey = true;
void optionKeysComplete;

/** The mount options among `props`; absent and `undefined` props are left out. */
export const optionsOf = (props: SpreadsheetOptions): SpreadsheetOptions => {
  const options: Record<string, unknown> = {};
  for (const key of OPTION_KEYS) if (props[key] !== undefined) options[key] = props[key];
  return options as SpreadsheetOptions;
};

/** The keys of `options` that `update` accepts. */
export const patchOf = (options: SpreadsheetOptions): Partial<SpreadsheetPatch> => {
  const { seed: _seed, signal: _signal, ...patch } = options;
  return patch;
};
