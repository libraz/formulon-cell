import type { SpreadsheetInstance, ToolbarInstance } from '@libraz/formulon-cell';
import type { DefineComponent } from 'vue';

type RibbonTab =
  | 'file'
  | 'home'
  | 'insert'
  | 'draw'
  | 'pageLayout'
  | 'formulas'
  | 'data'
  | 'review'
  | 'view'
  | 'help'
  | 'automate'
  | 'acrobat';

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
  ribbonTabs?: readonly RibbonTab[];
}>;

export default SpreadsheetToolbar;
