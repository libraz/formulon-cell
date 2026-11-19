/** External formula-draft model: types and pure helpers shared with the formula bar controller. */
import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { FormulaEditLease } from '../interact/formula-edit-lease.js';

export interface ExternalFormulaDraftView {
  raw: string;
  caret: {
    start: number;
    end: number;
    direction: HTMLTextAreaElement['selectionDirection'];
  };
}

export type ExternalFormulaDraftOutcome = 'committed' | 'cancelled' | 'discarded';

export interface ExternalFormulaDraftHooks {
  onFinish(outcome: ExternalFormulaDraftOutcome, restoredFocusTarget?: HTMLElement | null): void;
}

export interface ExternalFormulaDraftHandle {
  readonly anchor: Addr;
  value(): string;
  snapshot(): ExternalFormulaDraftView | null;
  setValue(raw: string, caret?: number): void;
  commit(): boolean;
  cancel(): void;
  discard(): void;
  subscribe(fn: (view: ExternalFormulaDraftView) => void): () => void;
}

export type ExternalDraftStatus = 'open' | 'committing' | 'committed' | 'cancelled' | 'discarded';

export interface ExternalDraftState {
  anchor: Addr;
  workbook: WorkbookHandle;
  restoreValue: string;
  hooks: ExternalFormulaDraftHooks;
  listeners: Set<(view: ExternalFormulaDraftView) => void>;
  lastNotifiedView: ExternalFormulaDraftView | null;
  status: ExternalDraftStatus;
  cancelRequested: boolean;
  cancelMode: 'restore' | 'discard' | null;
  lease: FormulaEditLease | null;
}

export const workbookLive = (workbook: WorkbookHandle): boolean =>
  Number.isInteger(workbook.sheetCount) && workbook.sheetCount >= 0;

export const externalDraftWorkbookCurrent = (
  draft: ExternalDraftState,
  wb: () => WorkbookHandle,
): boolean => {
  try {
    const current = wb();
    return current === draft.workbook && workbookLive(current);
  } catch {
    return false;
  }
};

export const externalDraftStale = (draft: ExternalDraftState, wb: () => WorkbookHandle): boolean =>
  !externalDraftWorkbookCurrent(draft, wb) || (draft.lease !== null && !draft.lease.valid());

/** Notifies listeners of a changed view; stops once the draft is closed, replaced or superseded. */
export function publishExternalDraftView(
  draft: ExternalDraftState,
  view: ExternalFormulaDraftView,
  isCurrent: () => boolean,
): void {
  const last = draft.lastNotifiedView;
  if (
    last &&
    last.raw === view.raw &&
    last.caret.start === view.caret.start &&
    last.caret.end === view.caret.end &&
    last.caret.direction === view.caret.direction
  )
    return;
  draft.lastNotifiedView = view;
  for (const listener of [...draft.listeners]) {
    if (draft.status !== 'open' || !isCurrent() || draft.lastNotifiedView !== view) break;
    listener(view);
  }
}
