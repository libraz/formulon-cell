import type { DefineComponent } from 'vue';
import type { SpreadsheetToolbarProps } from './toolbar.js';

declare const SpreadsheetToolbar: DefineComponent<SpreadsheetToolbarProps>;

export const Toolbar: typeof SpreadsheetToolbar;
export default SpreadsheetToolbar;
