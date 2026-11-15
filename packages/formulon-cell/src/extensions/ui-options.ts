import { ALL_FEATURE_IDS, type FeatureFlags } from './features.js';
import { full, minimal, standard } from './presets.js';
import type { ThemeName } from './types.js';

/**
 * Named starting point for `features`: `'embedded'` switches every feature off,
 * `'minimal'` and `'standard'` apply the preset of that name, `'full'` leaves each at its default.
 */
export type SpreadsheetUiProfile = 'embedded' | 'minimal' | 'standard' | 'full';

/** Platform preference for desktop-spreadsheet interaction details. */
export type SpreadsheetPlatform = 'auto' | 'mac' | 'default';

/** Platform selected after resolving an explicit preference or the browser. */
export type ResolvedSpreadsheetPlatform = 'mac' | 'default';

interface PlatformNavigator {
  platform?: string;
  userAgent?: string;
  maxTouchPoints?: number;
}

/**
 * Resolve the platform used by the spreadsheet keyboard and chrome layers.
 *
 * Hosts get the default platform unless they pass `'mac'` or opt in to
 * browser detection with `'auto'`.
 *
 * iPadOS can report itself as `MacIntel` when desktop-site mode is enabled.
 * A touch-capable MacIntel navigator is therefore treated as the default
 * platform so that touch iPad users do not receive Mac-only shortcuts.
 */
export function resolveSpreadsheetPlatform(
  requested: SpreadsheetPlatform = 'default',
  navigatorLike: PlatformNavigator | undefined = typeof navigator === 'undefined'
    ? undefined
    : navigator,
): ResolvedSpreadsheetPlatform {
  if (requested === 'mac') return 'mac';
  if (requested === 'default') return 'default';

  const platform = navigatorLike?.platform?.toLowerCase() ?? '';
  const userAgent = navigatorLike?.userAgent?.toLowerCase() ?? '';
  const touchPoints = navigatorLike?.maxTouchPoints ?? 0;
  const touchIpad =
    /ipad|iphone|ipod/.test(userAgent) ||
    ((platform === 'macintel' || platform.includes('mac')) && touchPoints > 0);
  if (touchIpad) return 'default';

  return platform.includes('mac') || userAgent.includes('macintosh') ? 'mac' : 'default';
}

/** Whether `el` sits in a spreadsheet host rendered with the Mac platform profile. */
export const isMacHost = (el: HTMLElement): boolean =>
  (el.closest<HTMLElement>('.fc-host') ?? el).dataset.fcPlatform === 'mac';

/**
 * Switches for the spreadsheet's chrome and features, keyed by feature id.
 * `print` is a chrome switch; `conditionalFormatting`, `comments` and `pivotTable` are the
 * friendly names of `conditional`, `commentDialog` and `pivotTableDialog`. An omitted switch
 * means the feature's default (see {@link FeatureFlags}). The explicit switches win over the
 * profile. No switch shows the ribbon: mount it with the `toolbar` option.
 */
export interface SpreadsheetFeatureSwitches extends FeatureFlags {
  /** Offer printing and page setup. Defaults to `true` except for the `'embedded'` profile. */
  print?: boolean;
  conditionalFormatting?: boolean;
  comments?: boolean;
  pivotTable?: boolean;
}

export interface SpreadsheetUiOptions {
  /** Starting set of features; explicit `features` entries are merged over it. Defaults to `'full'`. */
  profile?: SpreadsheetUiProfile;
  /**
   * Interaction profile. Defaults to `'default'`; pass `'mac'` to force the
   * Mac profile or `'auto'` to detect it from the browser.
   */
  platform?: SpreadsheetPlatform;
  features?: SpreadsheetFeatureSwitches;
}

export interface ResolvedSpreadsheetUiOptions {
  profile: SpreadsheetUiProfile;
  platform: ResolvedSpreadsheetPlatform;
  /** The theme applied when the host passes no `theme` option: always `'paper'`. */
  theme: ThemeName;
  print: boolean;
  features: FeatureFlags;
}

const profileFlags = (profile: SpreadsheetUiProfile): FeatureFlags => {
  switch (profile) {
    case 'embedded':
      return Object.fromEntries(ALL_FEATURE_IDS.map((id) => [id, false])) as FeatureFlags;
    case 'minimal':
      return minimal();
    case 'standard':
      return standard();
    case 'full':
      return full();
  }
};

/** The profile, platform, theme and feature switches `mount` resolves from `ui`: what a host's
 *  options panel shows as the current state before the user changes anything. */
export function resolveSpreadsheetUiOptions(
  opts: SpreadsheetUiOptions = {},
): ResolvedSpreadsheetUiOptions {
  const profile = opts.profile ?? 'full';
  const switches = opts.features ?? {};
  const features: FeatureFlags = { ...profileFlags(profile) };

  const apply = <K extends keyof FeatureFlags>(target: K, value: boolean | undefined): void => {
    if (value !== undefined) features[target] = value;
  };

  for (const id of ALL_FEATURE_IDS) apply(id, switches[id]);
  apply('conditional', switches.conditionalFormatting);
  apply('commentDialog', switches.comments);
  apply('pivotTableDialog', switches.pivotTable);

  if (switches.print === false && switches.pageSetup === undefined) {
    features.pageSetup = false;
  }

  return {
    profile,
    platform: resolveSpreadsheetPlatform(opts.platform),
    theme: 'paper',
    print: switches.print ?? profile !== 'embedded',
    features,
  };
}
