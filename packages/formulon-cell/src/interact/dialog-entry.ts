import type { InteractionOrigin } from '../commands/interaction-policy.js';
import type { SpreadsheetStore } from '../store/store.js';
import type { TabId } from './format-dialog-model.js';
import type { PageSetupDialogTab } from './page-setup-dialog.js';

/** The UI entry point that opened a dialog: its commit, repeat and their
 *  undo/redo are authorized under this origin and command id. */
export interface DialogEntry {
  readonly origin: InteractionOrigin;
  readonly commandId: string;
}

/** The entry of a Format Cells open through the public instance method. */
export const INSTANCE_API_FORMAT_ENTRY: DialogEntry = {
  origin: 'instanceApi',
  commandId: 'formatCells',
};

/** The entry of a Data Validation open through the public instance method. */
export const INSTANCE_API_VALIDATION_ENTRY: DialogEntry = {
  origin: 'instanceApi',
  commandId: 'dataValidation',
};

/** The entry of a Name Manager open through the public instance method. */
export const INSTANCE_API_NAME_MANAGER_ENTRY: DialogEntry = {
  origin: 'instanceApi',
  commandId: 'namedRanges',
};

/** The entry of a Define Name open through the public instance method. */
export const INSTANCE_API_DEFINE_NAME_ENTRY: DialogEntry = {
  origin: 'instanceApi',
  commandId: 'defineName',
};

/** The entry of a Page Setup open through the public instance method. */
export const INSTANCE_API_PAGE_SETUP_ENTRY: DialogEntry = {
  origin: 'instanceApi',
  commandId: 'pageSetup',
};

/** The entry of an Insert Hyperlink open through the public instance method. */
export const INSTANCE_API_HYPERLINK_ENTRY: DialogEntry = {
  origin: 'instanceApi',
  commandId: 'insertHyperlink',
};

/** The entry of an Insert PivotTable open through the public instance method. */
export const INSTANCE_API_PIVOT_TABLE_ENTRY: DialogEntry = {
  origin: 'instanceApi',
  commandId: 'pivotTableInsert',
};

/** The mount's dialog openers that take the entry of an internal UI caller. */
export interface DialogOpeners {
  readonly formatDialog: (tab: TabId | undefined, entry: DialogEntry) => void;
  readonly dataValidationDialog: (entry: DialogEntry) => void;
  readonly namedRangeDialog: (entry: DialogEntry) => void;
  readonly defineNameDialog: (entry: DialogEntry) => void;
  readonly pageSetupDialog: (tab: PageSetupDialogTab | undefined, entry: DialogEntry) => void;
  readonly hyperlinkDialog: (entry: DialogEntry) => void;
  readonly pivotTableDialog: (
    placement: 'new' | 'existing' | undefined,
    entry: DialogEntry,
  ) => void;
}

const openers = new WeakMap<SpreadsheetStore, DialogOpeners>();

/** Register the mounted instance's openers; they live as long as its store. */
export function registerDialogOpeners(store: SpreadsheetStore, next: DialogOpeners): void {
  openers.set(store, next);
}

/** Open Format Cells on behalf of an internal UI entry point. */
export function openFormatDialogFrom(
  store: SpreadsheetStore,
  tab: TabId | undefined,
  entry: DialogEntry,
): void {
  openers.get(store)?.formatDialog(tab, entry);
}

/** Open Data Validation on behalf of an internal UI entry point. */
export function openDataValidationDialogFrom(store: SpreadsheetStore, entry: DialogEntry): void {
  openers.get(store)?.dataValidationDialog(entry);
}

/** Open the Name Manager on behalf of an internal UI entry point. */
export function openNamedRangeDialogFrom(store: SpreadsheetStore, entry: DialogEntry): void {
  openers.get(store)?.namedRangeDialog(entry);
}

/** Open Define Name on behalf of an internal UI entry point. */
export function openDefineNameDialogFrom(store: SpreadsheetStore, entry: DialogEntry): void {
  openers.get(store)?.defineNameDialog(entry);
}

/** Open Page Setup on behalf of an internal UI entry point. */
export function openPageSetupDialogFrom(
  store: SpreadsheetStore,
  tab: PageSetupDialogTab | undefined,
  entry: DialogEntry,
): void {
  openers.get(store)?.pageSetupDialog(tab, entry);
}

/** Open Insert Hyperlink on behalf of an internal UI entry point. */
export function openHyperlinkDialogFrom(store: SpreadsheetStore, entry: DialogEntry): void {
  openers.get(store)?.hyperlinkDialog(entry);
}

/** Open Insert PivotTable on behalf of an internal UI entry point. */
export function openPivotTableDialogFrom(
  store: SpreadsheetStore,
  placement: 'new' | 'existing' | undefined,
  entry: DialogEntry,
): void {
  openers.get(store)?.pivotTableDialog(placement, entry);
}
