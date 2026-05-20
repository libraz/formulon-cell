import type {
  DynamicDropdownsCtx,
  RibbonTab,
  SpreadsheetInstance,
  ToolbarInstance,
} from '@libraz/formulon-cell';
import type { DefineComponent } from 'vue';

declare const SpreadsheetToolbar: DefineComponent<{
  instance: SpreadsheetInstance | null;
  activeTab: RibbonTab;
  locale: string;
  onSpellingReview?: () => void;
  onAccessibilityCheck?: () => void;
  onRunScript?: () => void;
  onDrawPen?: () => void;
  onDrawEraser?: () => void;
  onTranslate?: () => void;
  onAddIn?: () => void;
  onToolbarReady?: (toolbar: ToolbarInstance | null) => void;
  dropdownActions?: Partial<DynamicDropdownsCtx>;
  ribbonTabs?: readonly RibbonTab[];
}>;

export default SpreadsheetToolbar;
