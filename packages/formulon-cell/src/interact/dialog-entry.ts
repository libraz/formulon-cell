import type { InteractionOrigin } from '../commands/interaction-policy.js';
import type { SpreadsheetStore } from '../store/store.js';
import type { TabId } from './format-dialog-model.js';

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

/** The mount's dialog openers that take the entry of an internal UI caller. */
export interface DialogOpeners {
  readonly formatDialog: (tab: TabId | undefined, entry: DialogEntry) => void;
  readonly dataValidationDialog: (entry: DialogEntry) => void;
  readonly namedRangeDialog: (entry: DialogEntry) => void;
  readonly defineNameDialog: (entry: DialogEntry) => void;
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
