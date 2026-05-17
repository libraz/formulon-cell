import type {
  DynamicDropdownsCtx,
  FeatureFlags,
  RibbonTab,
  SpreadsheetInstance,
} from '@libraz/formulon-cell';

export type { ActiveState, BorderPreset, RibbonTab } from '@libraz/formulon-cell';
export {
  BORDER_PRESETS,
  BORDER_STYLES,
  EMPTY_ACTIVE_STATE,
  FONT_FAMILIES,
  FONT_SIZES,
  localizeBorderPresets,
  localizeBorderStyles,
  projectActiveState,
  RIBBON_KEYSHORTCUTS,
  RIBBON_TAB_LABELS,
  RIBBON_TABS,
  ribbonTabLabel,
} from '@libraz/formulon-cell';

export interface SpreadsheetToolbarProps {
  instance: SpreadsheetInstance | null;
  features?: FeatureFlags;
  activeTab: RibbonTab;
  onTabChange: (tab: RibbonTab) => void;
  locale: string;
  onSpellingReview?: () => void;
  onAccessibilityCheck?: () => void;
  onRunScript?: () => void;
  onDrawPen?: () => void;
  onDrawEraser?: () => void;
  onTranslate?: () => void;
  onAddIn?: () => void;
  onNewWorkbook?: () => void;
  onOpenWorkbook?: () => void;
  onSaveWorkbook?: () => void;
  onSaveWorkbookAs?: () => void;
  /** Override one or more entries in the core's default dynamic-dropdowns
   *  context. Use for dialog-opening handlers (sort, protect, file picker,
   *  etc.) that the wrapper can't represent as a named prop. Handlers
   *  supplied here win over the wrapper's built-in script/addIn wiring. */
  dropdownActions?: Partial<DynamicDropdownsCtx>;
}
