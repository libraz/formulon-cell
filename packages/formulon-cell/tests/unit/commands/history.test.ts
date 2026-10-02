import { beforeEach, describe, expect, it } from 'vitest';
import {
  applyChartsSnapshot,
  applyConditionalRulesSnapshot,
  applyFormatSnapshot,
  applyLayoutSnapshot,
  applyTableOverlaysSnapshot,
  captureChartsSnapshot,
  captureConditionalRulesSnapshot,
  captureFormatSnapshot,
  captureLayoutSnapshot,
  captureTableOverlaysSnapshot,
  History,
  type HistoryEntry,
  type HistoryTransaction,
  recordChartsChange,
  recordConditionalRulesChange,
  recordFormatChange,
  recordLayoutChange,
  recordMergesChange,
  recordRepeatableFormatChange,
  recordTablesChange,
} from '../../../src/commands/history.js';
import type { OperationIntent } from '../../../src/commands/interaction-policy.js';
import {
  type ConditionalRule,
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

  it('repeats only an entry with an explicit selection-aware operation', () => {
    let repeats = 0;
    h.push({ undo: () => {}, redo: () => {} });
    expect(h.repeatLast()).toBe(false);

    h.setRepeat(() => (repeats += 1));
    expect(h.repeatLast()).toBe(true);
    expect(repeats).toBe(1);

    h.push({ undo: () => {}, redo: () => {} });
    expect(h.repeatLast()).toBe(false);

    h.push({ undo: () => {}, redo: () => {}, repeat: () => (repeats += 1) });
    expect(h.repeatLast()).toBe(true);
    expect(repeats).toBe(2);
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
    it('retains outer replay authorization and callback order for a composite entry', () => {
      const callbackLog: string[] = [];
      const guardCalls: Array<{
        direction: 'undo' | 'redo';
        entry: HistoryEntry;
      }> = [];
      const intent = (commandId: string): OperationIntent => ({
        operation: 'valueEdit',
        origin: 'ribbon',
        commandId,
        effects: [{ kind: 'workbook' }],
      });
      const undoIntent = intent('composite-undo');
      const redoIntent = intent('composite-redo');
      h.setGuard((entry, direction) => {
        guardCalls.push({ direction, entry });
        return true;
      });

      const tx = h.begin({
        replayAuthorization: {
          undo: [undoIntent],
          redo: [redoIntent],
        },
      });
      h.push({
        undo: () => callbackLog.push('undo-first'),
        redo: () => callbackLog.push('redo-first'),
      });
      h.push({
        undo: () => callbackLog.push('undo-second'),
        redo: () => callbackLog.push('redo-second'),
      });
      h.end(tx);

      expect(h.undo()).toBe(true);
      expect(callbackLog).toEqual(['undo-second', 'undo-first']);
      expect(guardCalls[0]).toMatchObject({
        direction: 'undo',
        entry: { replayAuthorization: { undo: [undoIntent], redo: [redoIntent] } },
      });
      callbackLog.length = 0;

      expect(h.redo()).toBe(true);
      expect(callbackLog).toEqual(['redo-first', 'redo-second']);
      expect(guardCalls[1]).toMatchObject({
        direction: 'redo',
        entry: { replayAuthorization: { undo: [undoIntent], redo: [redoIntent] } },
      });
    });

    it('rejects metadata on nested begin before changing the outer frame', () => {
      const outer = h.begin();
      expect(() =>
        h.begin({
          replayAuthorization: { undo: [], redo: [] },
        }),
      ).toThrow('History transaction metadata is only valid on an outer transaction');

      h.push({ undo: () => {}, redo: () => {} });
      h.end(outer);
      expect(h.canUndo()).toBe(true);
    });

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

    it('aborts only the current nested frame and leaves earlier outer entries pending', () => {
      const undoLog: string[] = [];
      const outer = h.begin();
      h.push({ undo: () => undoLog.push('A'), redo: () => {} });
      const inner = h.begin();
      h.push({ undo: () => undoLog.push('B'), redo: () => {} });
      h.push({ undo: () => undoLog.push('C'), redo: () => {} });

      h.abort(inner);

      expect(undoLog).toEqual(['C', 'B']);
      expect(h.canUndo()).toBe(false);
      h.end(outer);
      expect(h.canUndo()).toBe(true);
      h.undo();
      expect(undoLog).toEqual(['C', 'B', 'A']);
    });

    it('aborts entries in reverse order and suppresses pushes during rollback', () => {
      const log: string[] = [];
      const tx = h.begin();
      h.push({
        undo: () => {
          log.push('A');
          h.push({ undo: () => log.push('nested'), redo: () => {} });
        },
        redo: () => {},
      });
      h.push({ undo: () => log.push('B'), redo: () => {} });
      h.push({ undo: () => log.push('C'), redo: () => {} });

      h.abort(tx);

      expect(log).toEqual(['C', 'B', 'A']);
      expect(h.canUndo()).toBe(false);
      expect(h.canRedo()).toBe(false);
    });

    it('preserves a pre-existing redo entry when a transaction is aborted', () => {
      let value = 0;
      h.push({
        undo: () => {
          value = 0;
        },
        redo: () => {
          value = 1;
        },
      });
      h.undo();
      const tx = h.begin();
      h.push({ undo: () => {}, redo: () => {} });

      h.abort(tx);

      expect(h.canRedo()).toBe(true);
      expect(h.redo()).toBe(true);
      expect(value).toBe(1);
    });

    it('suppresses notifications when an abort inverse clears history', () => {
      h.push({ undo: () => {}, redo: () => {} });
      h.undo();
      const tx = h.begin();
      h.push({ undo: () => h.clear(), redo: () => {} });
      let notifications = 0;
      const off = h.subscribe(() => {
        notifications += 1;
      });

      h.abort(tx);

      expect(notifications).toBe(0);
      expect(h.canUndo()).toBe(false);
      expect(h.canRedo()).toBe(true);
      off();
    });

    it('rejects stale and non-top transaction tokens without changing frames', () => {
      const outer = h.begin();
      const inner = h.begin();

      expect(() => h.abort(outer)).toThrow();
      expect(() => h.end(outer)).toThrow();
      expect(() => h.abort(Symbol('stale') as HistoryTransaction)).toThrow();

      h.abort(inner);
      h.end(outer);
      expect(h.canUndo()).toBe(false);
      expect(() => h.abort(inner)).toThrow();
    });

    it('attempts every inverse and restores bookkeeping when an inverse throws', () => {
      const log: string[] = [];
      let repeatCount = 0;
      const base = {
        undo: () => {},
        redo: () => log.push('redo-base'),
      };
      h.push(base);
      h.undo();
      const repeat = () => {
        repeatCount += 1;
      };
      h.setRepeat(repeat);
      const outer = h.begin();
      h.push({ undo: () => log.push('outer'), redo: () => {} });
      const inner = h.begin();
      h.push({
        undo: () => {
          log.push('first');
          h.setRepeat(() => {
            repeatCount += 1000;
          });
          throw new Error('first inverse failed');
        },
        redo: () => {},
      });
      h.push({
        undo: () => {
          log.push('second');
          throw new Error('second inverse failed');
        },
        redo: () => {},
      });
      h.setRepeat(() => {
        repeatCount += 100;
      });
      let notifications = 0;
      const off = h.subscribe(() => {
        notifications += 1;
      });

      expect(() => h.abort(inner)).toThrow('second inverse failed');

      expect(log).toEqual(['second', 'first']);
      expect(notifications).toBe(0);
      expect(h.canUndo()).toBe(false);
      expect(h.canRedo()).toBe(true);
      expect(h.repeatLast()).toBe(true);
      expect(repeatCount).toBe(100);
      h.end(outer);
      expect(h.canUndo()).toBe(true);
      expect(h.canRedo()).toBe(false);
      off();
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
    expect(snap.formats.get('0:0:0')?.bold).toBe(true);
    expect(snap.formats.get('0:0:0')?.italic).toBeUndefined();
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
        hiddenSheets: new Set([1]),
        sheetTabColors: new Map([[1, '#c00000']]),
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
        hiddenSheets: new Set(),
        sheetTabColors: new Map(),
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
    expect(Array.from(snap.hiddenSheets)).toEqual([1]);
    expect(Array.from(snap.sheetTabColors.entries())).toEqual([[1, '#c00000']]);
  });

  it('applyLayoutSnapshot restores all fields', () => {
    store.setState((s) => ({
      ...s,
      layout: {
        ...s.layout,
        freezeRows: 2,
        hiddenRows: new Set([1]),
        sheetTabColors: new Map([[0, '#4472c4']]),
      },
    }));
    const snap = captureLayoutSnapshot(store.getState());
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, freezeRows: 0, hiddenRows: new Set([99]), sheetTabColors: new Map() },
    }));
    applyLayoutSnapshot(store, snap);
    const layout = store.getState().layout;
    expect(layout.freezeRows).toBe(2);
    expect(Array.from(layout.hiddenRows)).toEqual([1]);
    expect(Array.from(layout.sheetTabColors.entries())).toEqual([[0, '#4472c4']]);
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

  it('repeats a current-selection format command after the selection moves', () => {
    recordRepeatableFormatChange(h, store, () => {
      mutators.setCellFormat(store, store.getState().selection.active, { italic: true });
    });
    mutators.setActive(store, { sheet: 0, row: 0, col: 1 });

    expect(h.repeatLast()).toBe(true);
    expect(store.getState().format.formats.get('0:0:1')).toEqual({ italic: true });
  });

  it('recordFormatChange skips unchanged format snapshots', () => {
    recordFormatChange(h, store, () => {
      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { bold: true });
    });
    expect(h.canUndo()).toBe(true);

    h.undo();
    expect(store.getState().format.formats.size).toBe(0);

    recordFormatChange(h, store, () => {
      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, null);
    });
    expect(h.canUndo()).toBe(false);
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

  it('recordLayoutChange skips unchanged layout snapshots', () => {
    recordLayoutChange(h, store, () => {
      store.setState((s) => ({
        ...s,
        layout: { ...s.layout, freezeRows: s.layout.freezeRows, freezeCols: s.layout.freezeCols },
      }));
    });

    expect(h.canUndo()).toBe(false);
  });

  it('recordConditionalRulesChange round-trips rule presets', () => {
    const rule: ConditionalRule = {
      kind: 'color-scale',
      range: { sheet: 0, r0: 1, c0: 1, r1: 4, c1: 2 },
      stops: ['#63be7b', '#ffeb84', '#f8696b'],
    };

    recordConditionalRulesChange(h, store, () => {
      mutators.addConditionalRule(store, rule);
    });
    expect(captureConditionalRulesSnapshot(store.getState())).toEqual([rule]);
    expect(h.canUndo()).toBe(true);

    h.undo();
    expect(store.getState().conditional.rules).toEqual([]);

    h.redo();
    expect(store.getState().conditional.rules).toEqual([rule]);

    applyConditionalRulesSnapshot(store, []);
    expect(store.getState().conditional.rules).toEqual([]);
  });

  it('recordConditionalRulesChange skips unchanged rule snapshots', () => {
    recordConditionalRulesChange(h, store, () => {
      mutators.clearConditionalRulesInRange(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    });

    expect(h.canUndo()).toBe(false);
  });

  it('recordMergesChange skips unchanged merge snapshots', () => {
    recordMergesChange(h, store, () => {
      store.setState((s) => ({
        ...s,
        merges: {
          byAnchor: new Map(s.merges.byAnchor),
          byCell: new Map(s.merges.byCell),
        },
      }));
    });

    expect(h.canUndo()).toBe(false);
  });

  it('recordTablesChange round-trips session table overlays', () => {
    const table = {
      id: 'table-0-0-0-2-1',
      source: 'session' as const,
      range: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 },
      style: 'dark' as const,
      showHeader: true,
      showTotal: false,
      banded: true,
    };

    recordTablesChange(h, store, () => {
      mutators.upsertTableOverlay(store, table);
    });
    expect(captureTableOverlaysSnapshot(store.getState())).toEqual({
      tables: [table],
      customTableStyles: [],
      customPivotTableStyles: [],
      pivotTableStyles: [],
    });

    h.undo();
    expect(store.getState().tables.tables).toEqual([]);

    h.redo();
    expect(store.getState().tables.tables).toEqual([table]);

    applyTableOverlaysSnapshot(store, []);
    expect(store.getState().tables.tables).toEqual([]);
  });

  it('recordTablesChange skips unchanged table snapshots', () => {
    recordTablesChange(h, store, () => {
      store.setState((s) => ({ ...s, tables: { tables: [...s.tables.tables] } }));
    });

    expect(h.canUndo()).toBe(false);
  });

  it('recordChartsChange round-trips session chart overlays', () => {
    const chart = {
      id: 'chart-a',
      kind: 'column' as const,
      source: { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 },
      title: 'Chart',
      x: 10,
      y: 20,
      w: 300,
      h: 180,
    };

    recordChartsChange(h, store, () => {
      mutators.upsertChart(store, chart);
    });
    expect(captureChartsSnapshot(store.getState())).toEqual([chart]);

    h.undo();
    expect(store.getState().charts.charts).toEqual([]);

    h.redo();
    expect(store.getState().charts.charts).toEqual([chart]);

    applyChartsSnapshot(store, []);
    expect(store.getState().charts.charts).toEqual([]);
  });

  it('recordChartsChange skips unchanged chart snapshots', () => {
    recordChartsChange(h, store, () => {
      store.setState((s) => ({ ...s, charts: { charts: [...s.charts.charts] } }));
    });

    expect(h.canUndo()).toBe(false);
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
