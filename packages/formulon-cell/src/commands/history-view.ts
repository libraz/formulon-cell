import type { History } from './history.js';

/** The undo history as the public instance presents it. */
export interface HistoryView {
  readonly canUndo: boolean;
  readonly canRedo: boolean;
  /** Whether `repeat` would run an operation. */
  readonly canRepeat: boolean;
  undo(): void;
  redo(): void;
  /** Runs the latest repeatable action again on the current selection (F4); false when there is none. */
  repeat(): boolean;
  /** Drops every undo and redo step. */
  clear(): void;
  on(name: 'change', listener: () => void): () => void;
}

/** Wraps `history` in the public history shape. */
export const createHistoryView = (history: History): HistoryView => ({
  get canUndo() {
    return history.canUndo();
  },
  get canRedo() {
    return history.canRedo();
  },
  get canRepeat() {
    return history.canRepeat();
  },
  undo() {
    history.undo();
  },
  redo() {
    history.redo();
  },
  repeat() {
    return history.repeatLast();
  },
  clear() {
    history.clear();
  },
  on: (_name, listener) => history.subscribe(listener),
});
