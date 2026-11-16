export type FormulaDraftMode = 'edit' | 'point';

export interface FormulaDraftSnapshot {
  readonly value: string;
  readonly start: number;
  readonly end: number;
  readonly direction: HTMLTextAreaElement['selectionDirection'];
  readonly mode: FormulaDraftMode;
}

type TransactionKind = 'input' | 'command' | 'composition' | 'group';

interface Transaction {
  before: FormulaDraftSnapshot;
  after: FormulaDraftSnapshot;
  kind: TransactionKind;
  identity?: string;
}

interface ActiveGroup {
  identity: string;
  before: FormulaDraftSnapshot;
  after: FormulaDraftSnapshot;
}

interface CompositionState {
  before: FormulaDraftSnapshot;
  after: FormulaDraftSnapshot;
}

export interface FormulaDraftHistoryState {
  readonly undo: readonly Transaction[];
  readonly redo: readonly Transaction[];
  readonly observed: FormulaDraftSnapshot;
  readonly inputBefore: FormulaDraftSnapshot | null;
  readonly commandBefore: FormulaDraftSnapshot | null;
  readonly group: ActiveGroup | null;
  readonly composition: CompositionState | null;
}

/** Bounds the memory held by one edit session's undo stack. */
const MAX_TRANSACTIONS = 100;

const cloneSnapshot = (snapshot: FormulaDraftSnapshot): FormulaDraftSnapshot => ({ ...snapshot });

const cloneGroup = (group: ActiveGroup): ActiveGroup => ({
  identity: group.identity,
  before: cloneSnapshot(group.before),
  after: cloneSnapshot(group.after),
});

const cloneComposition = (composition: CompositionState): CompositionState => ({
  before: cloneSnapshot(composition.before),
  after: cloneSnapshot(composition.after),
});

const cloneTransaction = (transaction: Transaction): Transaction => ({
  ...transaction,
  before: cloneSnapshot(transaction.before),
  after: cloneSnapshot(transaction.after),
});

const sameSnapshot = (a: FormulaDraftSnapshot, b: FormulaDraftSnapshot): boolean =>
  a.value === b.value &&
  a.start === b.start &&
  a.end === b.end &&
  a.direction === b.direction &&
  a.mode === b.mode;

/** A snapshot with the caret collapsed at `caret`. */
export const caretSnapshot = (
  value: string,
  caret: number,
  mode: FormulaDraftMode,
): FormulaDraftSnapshot => ({ value, start: caret, end: caret, direction: 'none', mode });

export type DraftHistoryAction = 'undo' | 'redo';

/** The history action an Undo/Redo key chord asks for, or null for any other key. */
export const historyActionForKey = (
  event: Pick<KeyboardEvent, 'key' | 'ctrlKey' | 'metaKey' | 'altKey' | 'shiftKey'>,
): DraftHistoryAction | null => {
  if (event.ctrlKey === event.metaKey || event.altKey) return null;
  const key = event.key.toLowerCase();
  if ((key === 'z' && event.shiftKey) || (key === 'y' && !event.shiftKey)) return 'redo';
  return key === 'z' && !event.shiftKey ? 'undo' : null;
};

/** The history action a native `historyUndo`/`historyRedo` input type asks for. */
export const historyActionForInputType = (
  inputType: string | undefined,
): DraftHistoryAction | null =>
  inputType === 'historyUndo' ? 'undo' : inputType === 'historyRedo' ? 'redo' : null;

const fallbackSnapshot = (): FormulaDraftSnapshot => ({
  value: '',
  start: 0,
  end: 0,
  direction: 'none',
  mode: 'edit',
});

/** Local history for an uncommitted formula draft. Workbook history remains
 *  untouched until the draft commits. */
export class FormulaDraftHistory {
  private undoStack: Transaction[] = [];

  private redoStack: Transaction[] = [];

  private observed: FormulaDraftSnapshot;

  private inputBefore: FormulaDraftSnapshot | null = null;

  private commandBefore: FormulaDraftSnapshot | null = null;

  private group: ActiveGroup | null = null;

  private composition: CompositionState | null = null;

  constructor(initial: FormulaDraftSnapshot = fallbackSnapshot()) {
    this.observed = cloneSnapshot(initial);
  }

  reset(initial: FormulaDraftSnapshot): void {
    this.undoStack = [];
    this.redoStack = [];
    this.inputBefore = null;
    this.commandBefore = null;
    this.group = null;
    this.composition = null;
    this.observed = cloneSnapshot(initial);
  }

  clear(): void {
    this.reset(this.observed);
  }

  current(): FormulaDraftSnapshot {
    return cloneSnapshot(this.observed);
  }

  captureState(): FormulaDraftHistoryState {
    return {
      undo: this.undoStack.map(cloneTransaction),
      redo: this.redoStack.map(cloneTransaction),
      observed: cloneSnapshot(this.observed),
      inputBefore: this.inputBefore ? cloneSnapshot(this.inputBefore) : null,
      commandBefore: this.commandBefore ? cloneSnapshot(this.commandBefore) : null,
      group: this.group ? cloneGroup(this.group) : null,
      composition: this.composition ? cloneComposition(this.composition) : null,
    };
  }

  restoreState(state: FormulaDraftHistoryState): void {
    this.undoStack = state.undo.map(cloneTransaction);
    this.redoStack = state.redo.map(cloneTransaction);
    this.observed = cloneSnapshot(state.observed);
    this.inputBefore = state.inputBefore ? cloneSnapshot(state.inputBefore) : null;
    this.commandBefore = state.commandBefore ? cloneSnapshot(state.commandBefore) : null;
    this.group = state.group ? cloneGroup(state.group) : null;
    this.composition = state.composition ? cloneComposition(state.composition) : null;
  }

  get isCommandPending(): boolean {
    return this.commandBefore !== null;
  }

  get hasActiveGroup(): boolean {
    return this.group !== null;
  }

  /** Capture the browser's pre-edit value when `beforeinput` is available. */
  beforeInput(snapshot: FormulaDraftSnapshot): void {
    this.endGroup();
    this.commandBefore = null;
    this.inputBefore = cloneSnapshot(snapshot);
    this.observed = cloneSnapshot(snapshot);
  }

  /** Finish one native input transaction, or update an in-progress IME group. */
  input(
    snapshot: FormulaDraftSnapshot,
    options: { composing?: boolean; inputType?: string } = {},
  ): void {
    const after = cloneSnapshot(snapshot);
    const before = this.inputBefore ?? this.observed;
    this.inputBefore = null;

    if (options.composing && this.composition) {
      this.composition.after = after;
      this.observed = after;
      return;
    }

    if (
      options.inputType === 'insertFromComposition' &&
      this.undoStack.at(-1)?.kind === 'composition' &&
      this.undoStack.at(-1)?.after.value === this.observed.value
    ) {
      const transaction = this.undoStack.at(-1);
      if (transaction) transaction.after = after;
      this.observed = after;
      return;
    }

    this.endGroup();
    this.push({ before, after, kind: 'input' });
    this.observed = after;
  }

  startComposition(snapshot: FormulaDraftSnapshot): void {
    this.endGroup();
    this.inputBefore = null;
    this.commandBefore = null;
    const before = cloneSnapshot(snapshot);
    this.composition = { before, after: cloneSnapshot(before) };
    this.observed = before;
  }

  endComposition(snapshot: FormulaDraftSnapshot): void {
    const after = cloneSnapshot(snapshot);
    const composition = this.composition;
    this.composition = null;
    this.inputBefore = null;
    if (composition) {
      this.push({ before: composition.before, after, kind: 'composition' });
    }
    this.observed = after;
  }

  /** Start or continue a grouped pointer/keyboard reference gesture. */
  beginGroup(before: FormulaDraftSnapshot, identity: string): void {
    this.endGroup();
    this.inputBefore = null;
    this.commandBefore = null;
    const captured = cloneSnapshot(before);
    this.group = { identity, before: captured, after: cloneSnapshot(captured) };
    this.observed = captured;
  }

  updateGroup(after: FormulaDraftSnapshot, identity: string, before?: FormulaDraftSnapshot): void {
    if (!this.group || this.group.identity !== identity) {
      this.beginGroup(before ?? this.observed, identity);
    }
    const captured = cloneSnapshot(after);
    if (this.group) this.group.after = captured;
    this.observed = captured;
  }

  endGroup(): void {
    const group = this.group;
    this.group = null;
    if (!group) return;
    this.push({
      before: group.before,
      after: group.after,
      kind: 'group',
      identity: group.identity,
    });
    this.observed = cloneSnapshot(group.after);
  }

  beginCommand(before: FormulaDraftSnapshot): void {
    this.endGroup();
    this.inputBefore = null;
    this.commandBefore = cloneSnapshot(before);
    this.observed = cloneSnapshot(before);
  }

  finishCommand(after: FormulaDraftSnapshot): void {
    const captured = cloneSnapshot(after);
    const before = this.commandBefore ?? this.observed;
    this.commandBefore = null;
    this.inputBefore = null;
    this.push({ before, after: captured, kind: 'command' });
    this.observed = captured;
  }

  recordCommand(
    before: FormulaDraftSnapshot,
    after: FormulaDraftSnapshot,
    identity?: string,
  ): void {
    this.endGroup();
    this.inputBefore = null;
    this.commandBefore = null;
    const transaction: Transaction = {
      before: cloneSnapshot(before),
      after: cloneSnapshot(after),
      kind: 'command',
      ...(identity ? { identity } : {}),
    };
    this.push(transaction);
    this.observed = cloneSnapshot(after);
  }

  /** Caret/selection changes update the next edit's baseline but add no Undo. */
  observeSelection(snapshot: FormulaDraftSnapshot): void {
    const captured = cloneSnapshot(snapshot);
    if (this.group && !sameSnapshot(captured, this.group.after)) this.endGroup();
    if (!this.inputBefore && !this.commandBefore && !this.composition) {
      this.observed = captured;
    }
  }

  undo(): FormulaDraftSnapshot | null {
    this.endGroup();
    this.inputBefore = null;
    this.commandBefore = null;
    const transaction = this.undoStack.pop();
    if (!transaction) return null;
    this.redoStack.push(transaction);
    this.observed = cloneSnapshot(transaction.before);
    return cloneSnapshot(transaction.before);
  }

  redo(): FormulaDraftSnapshot | null {
    this.endGroup();
    this.inputBefore = null;
    this.commandBefore = null;
    const transaction = this.redoStack.pop();
    if (!transaction) return null;
    this.undoStack.push(transaction);
    this.observed = cloneSnapshot(transaction.after);
    return cloneSnapshot(transaction.after);
  }

  private push(transaction: Transaction): void {
    if (transaction.before.value === transaction.after.value) return;
    this.undoStack.push(cloneTransaction(transaction));
    if (this.undoStack.length > MAX_TRANSACTIONS) this.undoStack.shift();
    this.redoStack = [];
  }
}
