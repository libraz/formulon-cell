import type { FeatureFlags } from './features.js';
import { full, minimal, standard } from './presets.js';
import type { ThemeName } from './types.js';

export type SpreadsheetUiProfile = 'minimal' | 'standard' | 'excel365' | 'full';

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
  theme?: ThemeName;
  lockTheme?: boolean;
  features?: SpreadsheetFeatureSwitches;
  advancedFeatures?: FeatureFlags;
}

export interface ResolvedSpreadsheetUiOptions {
  profile: SpreadsheetUiProfile;
  theme: ThemeName;
  lockTheme: boolean;
  ribbon: boolean;
  print: boolean;
  features: FeatureFlags;
}

const profileFlags = (profile: SpreadsheetUiProfile): FeatureFlags => {
  switch (profile) {
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
    theme: opts.theme ?? 'paper',
    lockTheme: opts.lockTheme ?? false,
    ribbon: switches.ribbon ?? true,
    print: switches.print ?? true,
    features,
  };
}
