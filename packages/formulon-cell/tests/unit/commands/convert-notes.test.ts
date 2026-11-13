import { afterEach, describe, expect, it } from 'vitest';
import { listThreadedComments } from '../../../src/commands/comment.js';
import { convertNotesToThreads, workbookHasNotes } from '../../../src/commands/convert-notes.js';
import { History } from '../../../src/commands/history.js';
import { hydrateCommentsAndHyperlinksFromEngine } from '../../../src/engine/format-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const handles: WorkbookHandle[] = [];
afterEach(() => {
  for (const wb of handles.splice(0)) wb.dispose();
});

const seeded = async () => {
  const wb = await WorkbookHandle.createDefault();
  handles.push(wb);
  wb.addSheet('Second');
  wb.setCommentEntry(0, 0, 0, 'Ann', 'Check the total');
  wb.setCommentEntry(0, 2, 1, 'Bo', 'Late payment');
  wb.setCommentEntry(1, 4, 3, 'Ann', 'Second sheet note');
  const store = createSpreadsheetStore();
  hydrateCommentsAndHyperlinksFromEngine(wb, store, 0);
  const history = new History();
  const ctx = { store, workbook: wb, history, origin: 'ribbon', commandId: 'test' } as const;
  return { wb, store, history, ctx };
};

/** Each sheet's threads as cell, text and author. */
const threadsOf = (wb: WorkbookHandle, store: ReturnType<typeof createSpreadsheetStore>) => {
  const people = new Map(wb.getPersons().map((person) => [person.id, person.displayName]));
  return [0, 1].map((sheet) =>
    listThreadedComments(store, sheet, wb).map((entry) => [
      entry.cell.row,
      entry.cell.col,
      entry.text,
      people.get(entry.personId),
    ]),
  );
};

const notesOf = (wb: WorkbookHandle) =>
  [0, 1].map((sheet) => wb.getComments(sheet).map((note) => [note.row, note.col, note.text]));

describe('Convert to Comments', () => {
  it('turns every note of the workbook into a thread under its author', async () => {
    const { wb, store, ctx } = await seeded();
    expect(workbookHasNotes(wb)).toBe(true);
    expect(convertNotesToThreads(ctx)).toBe(3);
    expect(notesOf(wb)).toEqual([[], []]);
    expect(threadsOf(wb, store)).toEqual([
      [
        [0, 0, 'Check the total', 'Ann'],
        [2, 1, 'Late payment', 'Bo'],
      ],
      [[4, 3, 'Second sheet note', 'Ann']],
    ]);
    expect(workbookHasNotes(wb)).toBe(false);
  });

  it('undoes and redoes the whole conversion as one step', async () => {
    const { wb, store, history, ctx } = await seeded();
    const before = notesOf(wb);
    convertNotesToThreads(ctx);
    const after = threadsOf(wb, store);
    expect(history.undo()).toBe(true);
    expect(notesOf(wb)).toEqual(before);
    expect(threadsOf(wb, store)).toEqual([[], []]);
    expect(history.canUndo()).toBe(false);
    expect(history.redo()).toBe(true);
    expect(notesOf(wb)).toEqual([[], []]);
    expect(threadsOf(wb, store)).toEqual(after);
  });

  it('saves the threads and reads them back after a reload', async () => {
    const { wb, ctx } = await seeded();
    convertNotesToThreads(ctx);
    const reopened = await WorkbookHandle.loadBytes(wb.save());
    handles.push(reopened);
    const store = createSpreadsheetStore();
    expect(notesOf(reopened)).toEqual([[], []]);
    expect(threadsOf(reopened, store)).toEqual([
      [
        [0, 0, 'Check the total', 'Ann'],
        [2, 1, 'Late payment', 'Bo'],
      ],
      [[4, 3, 'Second sheet note', 'Ann']],
    ]);
  });

  it('does nothing in a workbook without notes', async () => {
    const wb = await WorkbookHandle.createDefault();
    handles.push(wb);
    const history = new History();
    const store = createSpreadsheetStore();
    expect(workbookHasNotes(wb)).toBe(false);
    expect(
      convertNotesToThreads({ store, workbook: wb, history, origin: 'ribbon', commandId: 'test' }),
    ).toBe(0);
    expect(history.canUndo()).toBe(false);
  });
});
