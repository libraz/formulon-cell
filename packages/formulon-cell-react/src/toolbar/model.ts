import type {
  DynamicDropdownsCtx,
  RibbonTab,
  SpreadsheetInstance,
  ToolbarInstance,
} from '@libraz/formulon-cell';

export type { ActiveState, BorderPreset, RibbonTab } from '@libraz/formulon-cell';
export {
  BORDER_PRESETS,
  BORDER_STYLES,
  EMPTY_ACTIVE_STATE,
  EXCEL365_STANDARD_RIBBON_TABS,
  FONT_FAMILIES,
  FONT_SIZES,
  localizeBorderPresets,
  localizeBorderStyles,
  OPTIONAL_RIBBON_TABS,
  projectActiveState,
  RIBBON_KEYSHORTCUTS,
  RIBBON_TAB_LABELS,
  RIBBON_TABS,
  ribbonTabLabel,
} from '@libraz/formulon-cell';

export interface SpreadsheetToolbarProps {
  instance: SpreadsheetInstance | null;
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
  /** Called when the core toolbar fails to mount. */
  onError?: (error: unknown) => void;
  /** Receives the mounted core toolbar instance so hosts can dispatch shared
   *  commands from titlebar search / Tell me without querying DOM buttons. */
  onToolbarReady?: (toolbar: ToolbarInstance | null) => void;
  /** Override one or more entries in the core's default dynamic-dropdowns
   *  context. Use for dialog-opening handlers (sort, protect, file picker,
   *  etc.) that the wrapper can't represent as a named prop. Handlers
   *  supplied here win over the wrapper's built-in script/addIn wiring. */
  dropdownActions?: Partial<DynamicDropdownsCtx>;
  /** Shared tab surface. Use `EXCEL365_STANDARD_RIBBON_TABS` for the
   *  baseline Excel profile; append `OPTIONAL_RIBBON_TABS` only when those
   *  add-in/automation tabs are intentionally exposed. */
  ribbonTabs?: readonly RibbonTab[];
}
