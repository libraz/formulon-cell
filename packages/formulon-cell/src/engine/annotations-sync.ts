import { hydrateThreadedComments } from '../commands/comment.js';
import type { SpreadsheetStore } from '../store/store.js';
import type { WorkbookHandle } from './workbook-handle.js';

/** Hydrate native threaded comments into the active UI cache. Legacy notes are
 * handled by format-sync and intentionally remain a separate model. */
export function hydrateThreadedCommentsFromEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
): void {
  if (!wb.capabilities.threadedComments) return;
  const comments = wb.getThreadedComments(sheet);
  const persons = wb.getPersons();
  hydrateThreadedComments(store, sheet, comments);
  store.setState((state) => ({
    ...state,
    threadedComments: {
      ...state.threadedComments,
      persons,
    },
  }));
}
