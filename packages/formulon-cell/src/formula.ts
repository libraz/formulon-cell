// Formula hints: names and argument labels of functions a host describes, for the cell
// editor's autocomplete list and argument helper. Custom functions themselves are LAMBDA
// defined names in the workbook; these hints only describe how to type a call.

export interface FormulaHintMeta {
  /** Short description for host code; the spreadsheet UI does not display it. */
  description?: string;
  /** Argument labels, shown in the autocomplete list and the argument helper while typing
   *  inside the call. Mark optional args with surrounding brackets: `'[step]'`. */
  args?: readonly string[];
}

export interface RegisteredFormulaHint {
  /** Upper-cased function name. Stored case-insensitively. */
  readonly name: string;
  readonly meta: FormulaHintMeta;
}

export class FormulaRegistry {
  private readonly entries = new Map<string, RegisteredFormulaHint>();
  private readonly listeners = new Set<() => void>();

  /** Register or replace the hint for `name`. Returns a disposer that removes only this exact
   *  registration. */
  register(name: string, meta: FormulaHintMeta = {}): () => void {
    const upper = name.toUpperCase();
    if (!/^[A-Z][A-Z0-9_]*$/.test(upper)) {
      throw new Error(`formulon-cell: invalid function name "${name}"`);
    }
    const entry: RegisteredFormulaHint = { name: upper, meta };
    this.entries.set(upper, entry);
    this.notify();
    return () => {
      // Only remove if we still own the slot — protects against leaking
      // a previous disposer after a later register replaces the entry.
      if (this.entries.get(upper) === entry) {
        this.entries.delete(upper);
        this.notify();
      }
    };
  }

  unregister(name: string): boolean {
    const ok = this.entries.delete(name.toUpperCase());
    if (ok) this.notify();
    return ok;
  }

  has(name: string): boolean {
    return this.entries.has(name.toUpperCase());
  }

  get(name: string): RegisteredFormulaHint | undefined {
    return this.entries.get(name.toUpperCase());
  }

  /** Sorted, upper-cased list of registered names. Stable for autocomplete. */
  list(): string[] {
    return [...this.entries.keys()].sort();
  }

  /** Subscribe to registry changes (register / unregister). Returns
   *  unsubscribe. Used by autocomplete to keep its menu fresh. */
  subscribe(fn: () => void): () => void {
    this.listeners.add(fn);
    return () => {
      this.listeners.delete(fn);
    };
  }

  /** Drop every registered function. Mostly useful for tests. */
  clear(): void {
    if (this.entries.size === 0) return;
    this.entries.clear();
    this.notify();
  }

  private notify(): void {
    for (const fn of [...this.listeners]) {
      try {
        fn();
      } catch (err) {
        console.error('formulon-cell: formula-registry listener threw', err);
      }
    }
  }
}
