import { beforeEach, describe, expect, it } from 'vitest';
import { pasteSpecial } from '../../../../../src/commands/clipboard/paste-special.js';
import { captureSnapshot } from '../../../../../src/commands/clipboard/snapshot.js';
import { History } from '../../../../../src/commands/history.js';
import { addrKey, type WorkbookHandle } from '../../../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../../src/store/store.js';
import { assertSnap, newWb, seedAndMirror, setActive } from './fixtures.js';

describe('pasteSpecial', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
  });

  it('copies All comments and authors with the source payload', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 7 },
      { row: 3, col: 3, value: 11 },
    ]);
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        bold: true,
        comment: 'source note',
        commentAuthor: 'Alice',
      },
    );
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 3, col: 3 },
      {
        italic: true,
        comment: 'destination note',
        commentAuthor: 'Bob',
      },
    );
    const snap = captureSnapshot(store.getState(), {
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 0,
      c1: 0,
    });
    setActive(store, 3, 3);
    assertSnap(snap);

    pasteSpecial(store.getState(), store, wb, snap, {
      what: 'all',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });

    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))).toEqual({
      bold: true,
      comment: 'source note',
      commentAuthor: 'Alice',
    });
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 3, col: 3 }))).toEqual({
      bold: true,
      comment: 'source note',
      commentAuthor: 'Alice',
    });
  });

  it('Formats paste preserves destination comment metadata', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 7 }]);
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        bold: true,
        comment: 'source note',
        commentAuthor: 'Alice',
      },
    );
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 3, col: 3 },
      {
        italic: true,
        comment: 'destination note',
        commentAuthor: 'Bob',
      },
    );
    const snap = captureSnapshot(store.getState(), {
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 0,
      c1: 0,
    });
    setActive(store, 3, 3);
    assertSnap(snap);

    pasteSpecial(store.getState(), store, wb, snap, {
      what: 'formats',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });

    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 3, col: 3 }))).toEqual({
      bold: true,
      comment: 'destination note',
      commentAuthor: 'Bob',
    });
  });

  it('cuts All comments and restores them with one undo action', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 7 },
      { row: 3, col: 3, value: 11 },
    ]);
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        comment: 'source note',
        commentAuthor: 'Alice',
      },
    );
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 3, col: 3 },
      {
        comment: 'destination note',
        commentAuthor: 'Bob',
      },
    );
    const snap = captureSnapshot(
      store.getState(),
      {
        sheet: 0,
        r0: 0,
        c0: 0,
        r1: 0,
        c1: 0,
      },
      'cut',
    );
    setActive(store, 3, 3);
    assertSnap(snap);
    const history = new History();
    history.begin();
    pasteSpecial(
      store.getState(),
      store,
      wb,
      snap,
      { what: 'all', operation: 'none', skipBlanks: false, transpose: false },
      history,
    );
    history.end();

    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.comment,
    ).toBeUndefined();
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 3, col: 3 }))).toEqual({
      comment: 'source note',
      commentAuthor: 'Alice',
    });
    expect(history.undo()).toBe(true);
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))).toEqual({
      comment: 'source note',
      commentAuthor: 'Alice',
    });
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 3, col: 3 }))).toEqual({
      comment: 'destination note',
      commentAuthor: 'Bob',
    });
    expect(history.redo()).toBe(true);
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.comment,
    ).toBeUndefined();
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 3, col: 3 }))).toEqual({
      comment: 'source note',
      commentAuthor: 'Alice',
    });
  });
});
