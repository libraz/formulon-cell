import type { RibbonTab, SpreadsheetInstance } from '@libraz/formulon-cell';

export type { ActiveState, BorderPreset, RibbonTab } from '@libraz/formulon-cell';
export {
  BORDER_PRESETS,
  BORDER_STYLES,
  EMPTY_ACTIVE_STATE,
  FONT_FAMILIES,
  FONT_SIZES,
  projectActiveState,
  RIBBON_KEYSHORTCUTS,
  RIBBON_TAB_LABELS,
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
}
