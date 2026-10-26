import type { SpreadsheetStore, State } from '../store/store.js';
import { ShapeMirror } from './shapes-sync.js';
import { SlicerMirror } from './slicer-sync.js';
import { SparklineMirror } from './sparkline-sync.js';
import type { WorkbookHandle } from './workbook-handle.js';

/** Writes the store's shapes, sparklines and slicers to the engine as they
 *  change, so a save carries them. */
export class DrawingObjectMirror {
  constructor(handle: WorkbookHandle) {
    this.mirrors = [new ShapeMirror(handle), new SparklineMirror(handle), new SlicerMirror(handle)];
  }

  private readonly mirrors: readonly (ShapeMirror | SparklineMirror | SlicerMirror)[];

  /** Adopt `state` as what the engine already carries. */
  rebase(state: State | null): void {
    for (const mirror of this.mirrors) mirror.rebase(state);
  }

  /** Returns true when a write tied the store to the engine, which notified it again. */
  sync(store: SpreadsheetStore): boolean {
    const before = store.getState();
    for (const mirror of this.mirrors) mirror.sync(store);
    return store.getState() !== before;
  }
}
