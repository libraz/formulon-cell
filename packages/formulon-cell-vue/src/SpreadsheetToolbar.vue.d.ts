import type { SpreadsheetInstance } from '@libraz/formulon-cell';
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
  | 'automate'
  | 'acrobat';

declare const SpreadsheetToolbar: DefineComponent<{
  instance: SpreadsheetInstance | null;
  activeTab: RibbonTab;
  locale: string;
}>;

export default SpreadsheetToolbar;
