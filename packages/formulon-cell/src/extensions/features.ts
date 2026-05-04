// Built-in feature ids — the on/off keys consumers pass via
// `MountOptions.features`. These map 1:1 to the attach* calls inside
// mount.ts and the chrome elements appended to the host. Disabling a flag
// removes both the behavior *and* the DOM nodes (e.g. an empty status bar
// no longer reserves vertical space).
//
// `nameBox`, `editor`, `pointer`, and `renderer` are non-toggleable —
// they are the spreadsheet itself, removing them yields no UI.
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
  'fxDialog',
  'pasteSpecial',
  'validation',
  'autocomplete',
  'hoverComment',
  'clipboard',
  'wheel',
  'shortcuts',
  'watchWindow',
  'errorIndicators',
  'gotoSpecial',
] as const;

export type FeatureId = (typeof ALL_FEATURE_IDS)[number];

export type FeatureFlags = Partial<Record<FeatureId, boolean>>;

/** Features that ship default-off — adding them to the chrome opt-in lets us
 *  introduce new panels without expanding the default UI surface. */
const DEFAULT_OFF: ReadonlySet<FeatureId> = new Set(['watchWindow']);

/** Flags built-ins inside mount.ts gate against. Defaults to `true` for
 *  every feature unless explicitly disabled, except for `DEFAULT_OFF`
 *  members which start disabled and require explicit opt-in. */
export const resolveFlags = (input?: FeatureFlags): Record<FeatureId, boolean> => {
  const out = {} as Record<FeatureId, boolean>;
  for (const id of ALL_FEATURE_IDS) {
    if (DEFAULT_OFF.has(id)) {
      out[id] = input?.[id] === true;
    } else {
      out[id] = input?.[id] !== false;
    }
  }
  return out;
};
