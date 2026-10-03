import { ALL_FEATURE_IDS, type FeatureFlags } from './features.js';
import { full, minimal, standard } from './presets.js';
import type { ThemeName } from './types.js';

export type SpreadsheetUiProfile = 'embedded' | 'minimal' | 'standard' | 'excel365' | 'full';

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

export interface SpreadsheetFeatureSwitches {
  ribbon?: boolean;
  formulaBar?: boolean;
  viewToolbar?: boolean;
  sheetTabs?: boolean;
  statusBar?: boolean;
  workbookObjects?: boolean;
  contextMenu?: boolean;
  findReplace?: boolean;
  formatDialog?: boolean;
  formatPainter?: boolean;
  conditionalFormatting?: boolean;
  namedRanges?: boolean;
  hyperlink?: boolean;
  comments?: boolean;
  pasteSpecial?: boolean;
  quickAnalysis?: boolean;
  charts?: boolean;
  print?: boolean;
  pageSetup?: boolean;
  pivotTable?: boolean;
  validation?: boolean;
  autocomplete?: boolean;
  clipboard?: boolean;
  shortcuts?: boolean;
  wheel?: boolean;
  watchWindow?: boolean;
  slicer?: boolean;
}

export interface SpreadsheetUiOptions {
  profile?: SpreadsheetUiProfile;
  /**
   * Interaction profile. Defaults to `'default'`; pass `'mac'` to force the
   * Mac profile or `'auto'` to detect it from the browser.
   */
  platform?: SpreadsheetPlatform;
  theme?: ThemeName;
  lockTheme?: boolean;
  features?: SpreadsheetFeatureSwitches;
  advancedFeatures?: FeatureFlags;
}

export interface ResolvedSpreadsheetUiOptions {
  profile: SpreadsheetUiProfile;
  platform: ResolvedSpreadsheetPlatform;
  theme: ThemeName;
  lockTheme: boolean;
  ribbon: boolean;
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
    case 'excel365':
    case 'full':
      return full();
  }
};

export function resolveSpreadsheetUiOptions(
  opts: SpreadsheetUiOptions = {},
): ResolvedSpreadsheetUiOptions {
  const profile = opts.profile ?? 'excel365';
  const switches = opts.features ?? {};
  const features: FeatureFlags = {
    ...profileFlags(profile),
    ...opts.advancedFeatures,
  };

  const apply = <K extends keyof FeatureFlags>(target: K, value: boolean | undefined): void => {
    if (value !== undefined) features[target] = value;
  };

  apply('formulaBar', switches.formulaBar);
  apply('viewToolbar', switches.viewToolbar);
  apply('sheetTabs', switches.sheetTabs);
  apply('statusBar', switches.statusBar);
  apply('workbookObjects', switches.workbookObjects);
  apply('contextMenu', switches.contextMenu);
  apply('findReplace', switches.findReplace);
  apply('formatDialog', switches.formatDialog);
  apply('formatPainter', switches.formatPainter);
  apply('conditional', switches.conditionalFormatting);
  apply('namedRanges', switches.namedRanges);
  apply('hyperlink', switches.hyperlink);
  apply('commentDialog', switches.comments);
  apply('pasteSpecial', switches.pasteSpecial);
  apply('quickAnalysis', switches.quickAnalysis);
  apply('charts', switches.charts);
  apply('pageSetup', switches.pageSetup);
  apply('pivotTableDialog', switches.pivotTable);
  apply('validation', switches.validation);
  apply('autocomplete', switches.autocomplete);
  apply('clipboard', switches.clipboard);
  apply('shortcuts', switches.shortcuts);
  apply('wheel', switches.wheel);
  apply('watchWindow', switches.watchWindow);
  apply('slicer', switches.slicer);

  if (switches.print === false && switches.pageSetup === undefined) {
    features.pageSetup = false;
  }

  return {
    profile,
    platform: resolveSpreadsheetPlatform(opts.platform),
    theme: opts.theme ?? 'paper',
    lockTheme: opts.lockTheme ?? false,
    ribbon: switches.ribbon ?? profile !== 'embedded',
    print: switches.print ?? profile !== 'embedded',
    features,
  };
}
