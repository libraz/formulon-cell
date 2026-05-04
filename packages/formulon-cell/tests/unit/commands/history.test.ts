import { beforeEach, describe, expect, it } from 'vitest';
import {
  applyFormatSnapshot,
  applyLayoutSnapshot,
  captureFormatSnapshot,
  captureLayoutSnapshot,
  History,
  recordFormatChange,
  recordLayoutChange,
} from '../../../src/commands/history.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

describe('History stack', () => {
  let h: History;

  beforeEach(() => {
    h = new History();
  });

  it('starts empty', () => {
    expect(h.canUndo()).toBe(false);
    expect(h.canRedo()).toBe(false);
    expect(h.undo()).toBe(false);
    expect(h.redo()).toBe(false);
  });

  it('replays one entry', () => {
    let value = 0;
    h.push({
      undo: () => {
        value = 1;
      },
      redo: () => {
        value = 2;
      },
    });
    expect(h.canUndo()).toBe(true);
    expect(h.canRedo()).toBe(false);

    expect(h.undo()).toBe(true);
    expect(value).toBe(1);
    expect(h.canRedo()).toBe(true);

    expect(h.redo()).toBe(true);
    expect(value).toBe(2);
  });

  it('clears redo stack on new push', () => {
    let v = 0;
    h.push({
      undo: () => {
        v = -1;
      },
      redo: () => {
        v = 1;
      },
    });
    h.undo();
    expect(h.canRedo()).toBe(true);
    h.push({
      undo: () => {
        v = -2;
      },
      redo: () => {
        v = 2;
      },
    });
    expect(h.canRedo()).toBe(false);
    void v;
  });

  it('suppresses pushes during replay', () => {
    let inner = 0;
    let pushedDuringReplay = 0;
    h.push({
      undo: () => {
        inner = -1;
        // Simulate a nested push performed by an undo handler.
        h.push({
          undo: () => {
            pushedDuringReplay += 1;
          },
          redo: () => {},
        });
      },
      redo: () => {
        inner = 1;
      },
    });
    h.undo();
    expect(inner).toBe(-1);
    expect(pushedDuringReplay).toBe(0);
    expect(h.canUndo()).toBe(false); // suppressed entry must not exist
  });

  describe('transactions', () => {
    it('commits a single combined entry on end()', () => {
      const log: string[] = [];
      h.begin();
      h.push({
        undo: () => log.push('u1'),
        redo: () => log.push('r1'),
      });
      h.push({
        undo: () => log.push('u2'),
        redo: () => log.push('r2'),
      });
      h.end();

      expect(h.canUndo()).toBe(true);
      h.undo();
      // Undo runs in reverse insertion order.
      expect(log).toEqual(['u2', 'u1']);
      log.length = 0;
      h.redo();
      expect(log).toEqual(['r1', 'r2']);
    });

    it('end() with no entries is a no-op', () => {
      h.begin();
      h.end();
      expect(h.canUndo()).toBe(false);
    });

    it('handles nested begin/end correctly', () => {
      const log: string[] = [];
      h.begin();
      h.begin();
      h.push({
        undo: () => log.push('u'),
        redo: () => log.push('r'),
      });
      h.end(); // inner end — entry still buffered
      expect(h.canUndo()).toBe(false);
      h.end(); // outer end — commit
      expect(h.canUndo()).toBe(true);
    });
  });

  it('notifies subscribers on stack changes', () => {
    let notifications = 0;
    const off = h.subscribe(() => {
      notifications += 1;
    });
    h.push({ undo: () => {}, redo: () => {} });
    h.undo();
    h.redo();
    expect(notifications).toBeGreaterThanOrEqual(3);
    off();
  });

  it('clear() empties stacks and notifies', () => {
    let notified = false;
    h.subscribe(() => {
      notified = true;
    });
    h.push({ undo: () => {}, redo: () => {} });
    h.clear();
    expect(h.canUndo()).toBe(false);
    expect(h.canRedo()).toBe(false);
    expect(notified).toBe(true);
  });
});

describe('snapshot helpers', () => {
  let store: SpreadsheetStore;

  beforeEach(() => {
    store = createSpreadsheetStore();
  });

  it('captureFormatSnapshot returns a detached copy', () => {
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { bold: true });
    const snap = captureFormatSnapshot(store.getState());
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { italic: true });
    expect(snap.get('0:0:0')?.bold).toBe(true);
    expect(snap.get('0:0:0')?.italic).toBeUndefined();
  });

  it('applyFormatSnapshot restores prior state', () => {
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { bold: true });
    const before = captureFormatSnapshot(store.getState());
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { bold: false, italic: true });
    applyFormatSnapshot(store, before);
    expect(store.getState().format.formats.get('0:0:0')).toEqual({ bold: true });
  });

  it('captureLayoutSnapshot includes hidden sets and freeze panes', () => {
    store.setState((s) => ({
      ...s,
      layout: {
        ...s.layout,
        colWidths: new Map([[2, 200]]),
        rowHeights: new Map([[3, 40]]),
        freezeRows: 1,
        freezeCols: 2,
        hiddenRows: new Set([5]),
        hiddenCols: new Set([7, 8]),
      },
    }));
    const snap = captureLayoutSnapshot(store.getState());

    // Mutating live state must not bleed into the snapshot.
    store.setState((s) => ({
      ...s,
      layout: {
        ...s.layout,
        hiddenRows: new Set(),
        hiddenCols: new Set(),
        freezeRows: 0,
        freezeCols: 0,
      },
    }));

    expect(snap.colWidths.get(2)).toBe(200);
    expect(snap.rowHeights.get(3)).toBe(40);
    expect(snap.freezeRows).toBe(1);
    expect(snap.freezeCols).toBe(2);
    expect(Array.from(snap.hiddenRows)).toEqual([5]);
    expect(Array.from(snap.hiddenCols).sort()).toEqual([7, 8]);
  });

  it('applyLayoutSnapshot restores all fields', () => {
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, freezeRows: 2, hiddenRows: new Set([1]) },
    }));
    const snap = captureLayoutSnapshot(store.getState());
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, freezeRows: 0, hiddenRows: new Set([99]) },
    }));
    applyLayoutSnapshot(store, snap);
    const layout = store.getState().layout;
    expect(layout.freezeRows).toBe(2);
    expect(Array.from(layout.hiddenRows)).toEqual([1]);
  });
});

describe('recordFormatChange / recordLayoutChange', () => {
  let store: SpreadsheetStore;
  let h: History;

  beforeEach(() => {
    store = createSpreadsheetStore();
    h = new History();
  });

  it('recordFormatChange pushes one entry that round-trips', () => {
    recordFormatChange(h, store, () => {
      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { bold: true });
    });
    expect(h.canUndo()).toBe(true);
    expect(store.getState().format.formats.get('0:0:0')?.bold).toBe(true);

    h.undo();
    expect(store.getState().format.formats.get('0:0:0')).toBeUndefined();
    h.redo();
    expect(store.getState().format.formats.get('0:0:0')?.bold).toBe(true);
  });

  it('recordLayoutChange round-trips hidden rows', () => {
    recordLayoutChange(h, store, () => {
      store.setState((s) => ({
        ...s,
        layout: { ...s.layout, hiddenRows: new Set([3, 4]) },
      }));
    });
    expect(Array.from(store.getState().layout.hiddenRows).sort()).toEqual([3, 4]);

    h.undo();
    expect(Array.from(store.getState().layout.hiddenRows)).toEqual([]);

    h.redo();
    expect(Array.from(store.getState().layout.hiddenRows).sort()).toEqual([3, 4]);
  });

  it('passes through when history is null', () => {
    recordFormatChange(null, store, () => {
      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { bold: true });
    });
    expect(store.getState().format.formats.get('0:0:0')?.bold).toBe(true);
  });

  it('does not record while replaying', () => {
    recordFormatChange(h, store, () => {
      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { bold: true });
    });
    expect(h.canUndo()).toBe(true);

    // Manually wrap the undo with another recordFormatChange — the inner call
    // must not push a competing entry while history is replaying.
    h.push({
      undo: () =>
        recordFormatChange(h, store, () => {
          mutators.setCellFormat(store, { sheet: 0, row: 1, col: 0 }, { italic: true });
        }),
      redo: () => {},
    });
    const stackSizeBefore = (h as unknown as { undoStack: unknown[] }).undoStack.length;
    h.undo();
    const stackSizeAfter = (h as unknown as { undoStack: unknown[] }).undoStack.length;
    // One pop = one less in undoStack. No extra push happened.
    expect(stackSizeAfter).toBe(stackSizeBefore - 1);
  });
});
