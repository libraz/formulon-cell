// Cell renderer registry — lets consumers override the displayed string
// for matching cells.
//
// Formatters do not bypass canvas rendering — the grid still paints
// text via the normal cell-paint pipeline. They just substitute the
// string that gets painted, so font / color / format come from
// `CellFormat` as usual. Use this for domain transforms (badge labels,
// currency rounding, opaque IDs) without losing the spreadsheet's
// styling story.
import type { Addr, CellValue } from './engine/types.js';
import type { CellFormat } from './store/store.js';

export interface CellRenderInput {
  readonly addr: Addr;
  readonly value: CellValue;
  readonly formula: string | null;
  readonly format: CellFormat | undefined;
}

export interface CellFormatterEntry {
  /** Stable id — used for unregister() and to dedup repeated registrations. */
  readonly id: string;
  /** Predicate. Return true to make `format` run for this cell. */
  readonly match: (input: CellRenderInput) => boolean;
  /** Returns the displayed string. Return `null` to fall through to the
   *  next formatter (or the default). */
  readonly format: (input: CellRenderInput) => string | null;
  /** Lower runs first. Default 50. */
  readonly priority?: number;
}

export class CellRegistry {
  private readonly formatters: CellFormatterEntry[] = [];
  private readonly listeners = new Set<() => void>();

  /** Register a formatter. Returns a disposer. Last-wins on duplicate id. */
  registerFormatter(entry: CellFormatterEntry): () => void {
    const existing = this.formatters.findIndex((e) => e.id === entry.id);
    if (existing >= 0) this.formatters.splice(existing, 1);
    this.formatters.push(entry);
    this.formatters.sort((a, b) => (a.priority ?? 50) - (b.priority ?? 50));
    this.notify();
    return () => {
      const i = this.formatters.indexOf(entry);
      if (i >= 0) {
        this.formatters.splice(i, 1);
        this.notify();
      }
    };
  }

  unregisterFormatter(id: string): boolean {
    const i = this.formatters.findIndex((e) => e.id === id);
    if (i < 0) return false;
    this.formatters.splice(i, 1);
    this.notify();
    return true;
  }

  /** Run formatters in priority order; first non-null result wins.
   *  Returns null when no formatter matched, signalling the renderer
   *  should fall back to the default text. */
  resolveDisplay(input: CellRenderInput): string | null {
    for (const e of this.formatters) {
      if (!e.match(input)) continue;
      const out = e.format(input);
      if (out !== null && out !== undefined) return out;
    }
    return null;
  }

  /** Snapshot of registered formatter ids. */
  formatterIds(): string[] {
    return this.formatters.map((e) => e.id);
  }

  subscribe(fn: () => void): () => void {
    this.listeners.add(fn);
    return () => {
      this.listeners.delete(fn);
    };
  }

  private notify(): void {
    for (const fn of [...this.listeners]) {
      try {
        fn();
      } catch (err) {
        console.error('formulon-cell: cell-registry listener threw', err);
      }
    }
  }
}
