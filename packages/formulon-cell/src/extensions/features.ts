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
  'viewToolbar',
  'sheetTabs',
  'statusBar',
  'workbookObjects',
  'contextMenu',
  'findReplace',
  'formatDialog',
  'formatPainter',
  'borderDraw',
  'conditional',
  'iterative',
  'namedRanges',
  'hyperlink',
  'commentDialog',
  'fxDialog',
  'pasteSpecial',
  'quickAnalysis',
  'charts',
  'illustrations',
  'pivotTableDialog',
  'validation',
  'autocomplete',
  'hoverComment',
  'clipboard',
  'wheel',
  'shortcuts',
  'watchWindow',
  'errorIndicators',
  'gotoSpecial',
  'pageSetup',
  'slicer',
] as const;

export type FeatureId = (typeof ALL_FEATURE_IDS)[number];

export type FeatureFlags = Partial<Record<FeatureId, boolean>>;

/** Features that ship default-off — adding them to the chrome opt-in lets us
 *  introduce new panels without expanding the default UI surface.
 *
 *  `viewToolbar` is here because the ribbon's View tab already carries every
 *  control it offers, and a second always-visible strip for the same commands
 *  has no counterpart in the desktop app. Hosts that want the flat strip can
 *  still switch it on. */
export const DEFAULT_OFF_FEATURE_IDS: ReadonlySet<FeatureId> = new Set([
  'watchWindow',
  'slicer',
  'viewToolbar',
]);

/** Whether `id` starts disabled. Hosts rendering a feature-toggle UI need this
 *  to read the current value the same way `resolveFlags` does — an absent flag
 *  means "on" for most features and "off" for these. */
export const isFeatureDefaultOff = (id: FeatureId): boolean => DEFAULT_OFF_FEATURE_IDS.has(id);

/** Flags built-ins inside mount.ts gate against. Defaults to `true` for
 *  every feature unless explicitly disabled, except for `DEFAULT_OFF`
 *  members which start disabled and require explicit opt-in. */
export const resolveFlags = (input?: FeatureFlags): Record<FeatureId, boolean> => {
  const out = {} as Record<FeatureId, boolean>;
  for (const id of ALL_FEATURE_IDS) {
    if (DEFAULT_OFF_FEATURE_IDS.has(id)) {
      out[id] = input?.[id] === true;
    } else {
      out[id] = input?.[id] !== false;
    }
  }
  return out;
};
