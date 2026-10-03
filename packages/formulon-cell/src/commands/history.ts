import type { OperationIntent } from './interaction-policy.js';

const LIMIT = 200;

/** A reversible operation. Both functions must be idempotent w.r.t. each other —
 *  calling `undo` then `redo` must leave the system in the same state. */
export interface HistoryEntry {
  undo: () => void;
  redo: () => void;
  /** Reapply this logical command to the *current* selection. Unlike `redo`,
   *  this must not restore the original before/after snapshot. */
  repeat?: () => void;
  /** The logical command that produced this entry. Restricted instances use
   * this metadata to re-authorize undo and redo against the current policy. */
  intent?: OperationIntent;
  /** The logical command represented by the inverse replay. */
  inverseIntent?: OperationIntent;
  /** The complete logical authorization for a composite replay. */
  replayAuthorization?: {
    readonly undo: readonly OperationIntent[];
    readonly redo: readonly OperationIntent[];
  };
}

/* Keep these fields in one interface so transaction metadata can be overlaid
 * onto a single child entry without changing the existing entry contract. */
type HistoryBeginMetadata = Pick<HistoryEntry, 'replayAuthorization'>;

interface TransactionFrame {
  id: HistoryTransaction;
  startIndex: number;
  metadata?: HistoryBeginMetadata;
}

/* The logical command represented by a transaction is known only by its
 * caller. Child entries may be low-level structural callbacks with no policy
 * meaning, so History deliberately does not infer or union their intents. */

export type HistoryDirection = 'undo' | 'redo';
export type HistoryGuard = (entry: HistoryEntry, direction: HistoryDirection) => boolean;
/** Opaque handle for a transaction frame returned by `History.begin()`. */
export type HistoryTransaction = symbol;

/**
 * Single source of truth for undoable mutations. Cell writes (workbook),
 * format changes, and layout changes (col widths, row heights, freeze) all
 * push entries here so one Cmd/Ctrl+Z spans the whole instance.
 *
 * Use `begin()` / `end()` to batch multiple entries into one logical step
 * (e.g. paste-special, fill drag).
 */
export class History {
  private undoStack: HistoryEntry[] = [];
  private redoStack: HistoryEntry[] = [];
  private replaying = false;
  private txnFrames: TransactionFrame[] = [];
  private txnEntries: HistoryEntry[] = [];
  private listeners = new Set<() => void>();
  private lastRepeat: (() => void) | null = null;
  private guard: HistoryGuard | null = null;
  private notificationSuppressionDepth = 0;

  push(entry: HistoryEntry): void {
    if (this.replaying) return;
    if (this.txnFrames.length > 0) {
      this.txnEntries.push(entry);
      return;
    }
    this.commit(entry);
  }

  begin(metadata?: HistoryBeginMetadata): HistoryTransaction {
    if (this.txnFrames.length > 0 && metadata !== undefined) {
      throw new Error('History transaction metadata is only valid on an outer transaction');
    }
    const frame: TransactionFrame = {
      id: Symbol('history-transaction'),
      startIndex: this.txnEntries.length,
      metadata,
    };
    this.txnFrames.push(frame);
    return frame.id;
  }

  end(token?: HistoryTransaction): void {
    const frame = this.txnFrames.at(-1);
    if (!frame) {
      if (token !== undefined) throw new Error('History transaction token is not current');
      return;
    }
    if (token !== undefined && token !== frame.id) {
      throw new Error('History transaction token is not current');
    }
    this.txnFrames.pop();
    if (this.txnFrames.length > 0) return;
    const entries = this.txnEntries;
    this.txnEntries = [];
    if (entries.length === 0) return;
    if (entries.length === 1) {
      const only = entries[0];
      if (only) this.commit(this.withMetadata(only, frame.metadata));
      return;
    }
    this.commit({
      undo: () => {
        for (let i = entries.length - 1; i >= 0; i -= 1) entries[i]?.undo();
      },
      redo: () => {
        for (const e of entries) e.redo();
      },
      ...frame.metadata,
    });
  }

  private withMetadata(entry: HistoryEntry, metadata?: HistoryBeginMetadata): HistoryEntry {
    return metadata ? { ...entry, ...metadata } : entry;
  }

  /** Roll back the current transaction frame without touching committed history. */
  abort(token: HistoryTransaction): void {
    const frame = this.txnFrames.at(-1);
    if (!frame || frame.id !== token) {
      throw new Error('History transaction token is not current');
    }

    const pending = this.txnEntries.splice(frame.startIndex);
    this.txnFrames.pop();

    // Undo callbacks are user supplied and can themselves mutate history. Keep
    // every bookkeeping field stable while replaying the discarded entries;
    // pushes are suppressed by `replaying`, and the snapshots cover stronger
    // callbacks such as clear(), undo(), or setRepeat().
    const undoStack = [...this.undoStack];
    const redoStack = [...this.redoStack];
    const txnEntries = [...this.txnEntries];
    const txnFrames = [...this.txnFrames];
    const lastRepeat = this.lastRepeat;
    const guard = this.guard;
    const wasReplaying = this.replaying;
    const notificationSuppressionDepth = this.notificationSuppressionDepth;
    let firstError: unknown;
    let failed = false;
    this.replaying = true;
    this.notificationSuppressionDepth += 1;
    try {
      for (let i = pending.length - 1; i >= 0; i -= 1) {
        try {
          pending[i]?.undo();
        } catch (error) {
          if (!failed) {
            firstError = error;
            failed = true;
          }
        }
      }
    } finally {
      this.undoStack = undoStack;
      this.redoStack = redoStack;
      this.txnEntries = txnEntries;
      this.txnFrames = txnFrames;
      this.lastRepeat = lastRepeat;
      this.guard = guard;
      this.replaying = wasReplaying;
      this.notificationSuppressionDepth = notificationSuppressionDepth;
    }
    if (failed) throw firstError;
  }

  private commit(entry: HistoryEntry): void {
    this.undoStack.push(entry);
    if (this.undoStack.length > LIMIT) this.undoStack.shift();
    this.redoStack.length = 0;
    this.lastRepeat = entry.repeat ?? null;
    this.notify();
  }

  isReplaying(): boolean {
    return this.replaying;
  }

  undo(): boolean {
    const e = this.undoStack.at(-1);
    if (!e) return false;
    if (this.guard && !this.guard(e, 'undo')) return false;
    this.replaying = true;
    try {
      e.undo();
    } finally {
      this.replaying = false;
    }
    this.undoStack.pop();
    this.redoStack.push(e);
    this.notify();
    return true;
  }

  redo(): boolean {
    const e = this.redoStack.at(-1);
    if (!e) return false;
    if (this.guard && !this.guard(e, 'redo')) return false;
    this.replaying = true;
    try {
      e.redo();
    } finally {
      this.replaying = false;
    }
    this.redoStack.pop();
    this.undoStack.push(e);
    this.notify();
    return true;
  }

  canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  /** Register a selection-aware repeat operation for an action that has no
   *  material undo snapshot (for example a pending format on a blank cell). */
  setRepeat(repeat: (() => void) | null): void {
    this.lastRepeat = repeat;
  }

  /** Install an authorization guard for undo/redo. The stack is moved only
   * after the guard and replay both succeed. */
  setGuard(guard: HistoryGuard | null): void {
    this.guard = guard;
  }

  /** Repeat the latest command only when it explicitly supplied a
   *  selection-aware replay operation. Snapshot-only entries deliberately do
   *  not fall back to `redo`, because that would write to their old range. */
  repeatLast(): boolean {
    const repeat = this.lastRepeat;
    if (!repeat || this.replaying) return false;
    repeat();
    return true;
  }

  clear(): void {
    this.undoStack.length = 0;
    this.redoStack.length = 0;
    this.txnEntries.length = 0;
    this.txnFrames.length = 0;
    this.lastRepeat = null;
    this.notify();
  }

  subscribe(fn: () => void): () => void {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  }

  private notify(): void {
    if (this.notificationSuppressionDepth > 0) return;
    // History mutations are already committed when observers run. A stale
    // renderer must not make push/undo/redo throw after the stack moved, and
    // one bad observer must not prevent the remaining observers from seeing
    // the same transition.
    for (const l of [...this.listeners]) {
      try {
        l();
      } catch {
        // Observers are advisory; the history state is authoritative.
      }
    }
  }
}

/* ---------- Public undo/redo API ---------- */

/** Pull one entry off the undo stack and apply it. */
export function undo(history: History): boolean {
  return history.undo();
}

export function redo(history: History): boolean {
  return history.redo();
}

export function canUndo(history: History): boolean {
  return history.canUndo();
}

export function canRedo(history: History): boolean {
  return history.canRedo();
}
