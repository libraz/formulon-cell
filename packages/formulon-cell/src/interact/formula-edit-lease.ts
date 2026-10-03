import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type {
  EditorMode,
  EditorRefHighlight,
  PendingFormat,
  SelectionSlice,
  UiSlice,
} from '../store/types.js';

/** Session that borrows a suspended edit; a locale switch or a stale session invalidates the lease. */
export interface FormulaEditLeaseContext {
  getLocale(): string;
  contextCurrent(): boolean;
}

export interface FormulaEditLeaseSnapshot {
  readonly source: 'inline' | 'formulaBar';
  readonly workbook: WorkbookHandle;
  readonly anchor: Addr;
  readonly raw: string;
  readonly baseline: string;
  readonly caret: {
    readonly start: number;
    readonly end: number;
    readonly direction: HTMLTextAreaElement['selectionDirection'];
  };
  readonly selection: SelectionSlice;
  readonly editorMode: EditorMode;
  readonly pendingFormat: PendingFormat | null;
  readonly editorRefs: readonly EditorRefHighlight[];
  readonly copy: Pick<UiSlice, 'copyRange' | 'copyRanges' | 'copyMode' | 'copyRevision'>;
  readonly r1c1: boolean;
  readonly locale: string;
}

export interface FormulaEditLeaseOptions {
  readonly isOwnerCurrent: () => boolean;
  readonly context?: FormulaEditLeaseContext;
  /** Rebuild the owner from the lease's own deep-cloned snapshot. */
  readonly restore: (snapshot: Readonly<FormulaEditLeaseSnapshot>) => HTMLElement | null;
  readonly release: () => void;
}

export interface FormulaEditLease {
  readonly snapshot: Readonly<FormulaEditLeaseSnapshot>;
  valid(): boolean;
  /** Restore the suspended owner, without focusing it. Invalid leases discard. */
  userCancel(): HTMLElement | null;
  finalize(): void;
  discard(): void;
}

const cloneValue = <T>(value: T): T => {
  if (value === null || typeof value !== 'object') return value;
  if (Array.isArray(value)) return value.map((item) => cloneValue(item)) as T;
  const result: Record<string, unknown> = {};
  for (const [key, child] of Object.entries(value as Record<string, unknown>)) {
    result[key] = cloneValue(child);
  }
  return result as T;
};

const cloneSnapshot = (snapshot: FormulaEditLeaseSnapshot): FormulaEditLeaseSnapshot => ({
  source: snapshot.source,
  // Workbook identity is deliberately retained; cloning an engine handle would
  // make a stale lease appear valid while losing the protected workbook state.
  workbook: snapshot.workbook,
  anchor: cloneValue(snapshot.anchor),
  raw: snapshot.raw,
  baseline: snapshot.baseline,
  caret: cloneValue(snapshot.caret),
  selection: cloneValue(snapshot.selection),
  editorMode: cloneValue(snapshot.editorMode),
  pendingFormat: cloneValue(snapshot.pendingFormat),
  editorRefs: cloneValue(snapshot.editorRefs),
  copy: cloneValue(snapshot.copy),
  r1c1: snapshot.r1c1,
  locale: snapshot.locale,
});

export function createFormulaEditLease(
  snapshot: FormulaEditLeaseSnapshot,
  options: FormulaEditLeaseOptions,
): FormulaEditLease {
  const captured = cloneSnapshot(snapshot);
  let terminal = false;
  let consuming = false;

  const valid = (): boolean => {
    if (terminal || consuming) return false;
    try {
      // Accessing sheetCount is the public disposed-workbook probe. The owner
      // callback additionally checks generation, workbook, sheet, and R1C1.
      if (!Number.isInteger(captured.workbook.sheetCount) || captured.workbook.sheetCount < 0)
        return false;
      if (!options.isOwnerCurrent()) return false;
      if (options.context && options.context.getLocale() !== captured.locale) return false;
      if (options.context && !options.context.contextCurrent()) return false;
      return true;
    } catch {
      return false;
    }
  };

  const release = (): void => {
    if (terminal) return;
    terminal = true;
    options.release();
  };

  const userCancel = (): HTMLElement | null => {
    if (!valid()) {
      release();
      return null;
    }
    consuming = true;
    let restored: HTMLElement | null = null;
    try {
      restored = options.restore(captured);
    } finally {
      release();
    }
    return restored;
  };

  return {
    snapshot: captured,
    valid,
    userCancel,
    finalize: release,
    discard: release,
  };
}
