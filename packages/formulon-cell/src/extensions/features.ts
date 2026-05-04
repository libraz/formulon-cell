// Built-in feature ids — the on/off keys consumers pass via
// `MountOptions.features`. These map 1:1 to the attach* calls inside
// mount.ts, gated at mount time.
//
// `formulaBar`, `nameBox`, `editor`, `pointer`, `renderer`, and the host
// keymap are non-toggleable — they are the spreadsheet itself, removing
// them yields no UI. Everything else is opt-out.
export const ALL_FEATURE_IDS = [
  'formulaBar',
  'statusBar',
  'contextMenu',
  'findReplace',
  'formatDialog',
  'formatPainter',
  'conditional',
  'namedRanges',
  'hyperlink',
  'pasteSpecial',
  'validation',
  'autocomplete',
  'hoverComment',
  'clipboard',
  'wheel',
  'shortcuts',
] as const;

export type FeatureId = (typeof ALL_FEATURE_IDS)[number];

export type FeatureFlags = Partial<Record<FeatureId, boolean>>;

/** Flags built-ins inside mount.ts gate against. Defaults to `true` for
 *  every feature unless explicitly disabled. */
export const resolveFlags = (input?: FeatureFlags): Record<FeatureId, boolean> => {
  const out = {} as Record<FeatureId, boolean>;
  for (const id of ALL_FEATURE_IDS) {
    out[id] = input?.[id] !== false;
  }
  return out;
};
