import type { History } from './history.js';

/** The undo history as the public instance presents it. */
export interface HistoryView {
  readonly canUndo: boolean;
  readonly canRedo: boolean;
  undo(): void;
  redo(): void;
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
  undo() {
    history.undo();
  },
  redo() {
    history.redo();
  },
  clear() {
    history.clear();
  },
  on: (_name, listener) => history.subscribe(listener),
});
