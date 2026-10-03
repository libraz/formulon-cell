import { beforeEach, describe, expect, it } from 'vitest';
import {
  History,
  type HistoryEntry,
  type HistoryTransaction,
} from '../../../src/commands/history.js';
import type { OperationIntent } from '../../../src/commands/interaction-policy.js';

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
