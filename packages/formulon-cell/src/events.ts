// Public event surface for the spreadsheet instance.
//
// The store has its own subscribe(), and the workbook handle has its own
// subscribe() — both are still exposed for power users — but these are
// low-level and tied to internal shapes. For the typical "wire the
// spreadsheet to an outer state container" job, applications want a
// stable, event-named API. That is what this emitter provides.
//
// Adapter packages (`-react`, `-vue`) consume this through `inst.on()`.
import type { Addr, CellValue, Range } from './engine/types.js';
import type { ChangeEvent, WorkbookHandle } from './engine/workbook-handle.js';
import type { Strings } from './i18n/strings.js';
import type { State } from './store/store.js';

export interface CellChangeEvent {
  readonly addr: Addr;
  readonly value: CellValue;
  readonly formula: string | null;
}

export interface SelectionChangeEvent {
  readonly active: Addr;
  readonly anchor: Addr;
  readonly range: Range;
}

export interface WorkbookChangeEvent {
  readonly workbook: WorkbookHandle;
}

export interface LocaleChangeEvent {
  readonly locale: string;
  readonly strings: Strings;
}

export interface ThemeChangeEvent {
  readonly theme: string;
}

export interface RecalcEvent {
  /** Set of `${sheet}:${row}:${col}` keys the engine reported as dirty. */
  readonly dirty: ReadonlySet<string>;
}

/** Map of event name → payload. Used as the lookup table for `inst.on()`. */
export interface SpreadsheetEvents {
  cellChange: CellChangeEvent;
  selectionChange: SelectionChangeEvent;
  workbookChange: WorkbookChangeEvent;
  localeChange: LocaleChangeEvent;
  themeChange: ThemeChangeEvent;
  recalc: RecalcEvent;
}

export type SpreadsheetEventName = keyof SpreadsheetEvents;

export type SpreadsheetEventHandler<K extends SpreadsheetEventName> = (
  payload: SpreadsheetEvents[K],
) => void;

/** Internal emitter. Each listener-set is per-event so a missing event
 *  name returns the empty set, not undefined. */
export class SpreadsheetEmitter {
  private readonly listeners = new Map<SpreadsheetEventName, Set<(p: unknown) => void>>();

  on<K extends SpreadsheetEventName>(name: K, fn: SpreadsheetEventHandler<K>): () => void {
    let set = this.listeners.get(name);
    if (!set) {
      set = new Set();
      this.listeners.set(name, set);
    }
    set.add(fn as (p: unknown) => void);
    return () => {
      set?.delete(fn as (p: unknown) => void);
    };
  }

  off<K extends SpreadsheetEventName>(name: K, fn: SpreadsheetEventHandler<K>): void {
    this.listeners.get(name)?.delete(fn as (p: unknown) => void);
  }

  emit<K extends SpreadsheetEventName>(name: K, payload: SpreadsheetEvents[K]): void {
    const set = this.listeners.get(name);
    if (!set) return;
    // Snapshot to allow handlers to unsubscribe during iteration.
    for (const fn of [...set]) {
      try {
        (fn as SpreadsheetEventHandler<K>)(payload);
      } catch (err) {
        console.error(`formulon-cell: event "${name}" handler threw`, err);
      }
    }
  }

  dispose(): void {
    this.listeners.clear();
  }
}

/** Translate a workbook ChangeEvent into a public CellChangeEvent.
 *  Returns null for non-value events (those are routed to other channels). */
export const toCellChangeEvent = (
  e: ChangeEvent,
  formulaOf: (a: Addr) => string | null,
): CellChangeEvent | null => {
  if (e.kind !== 'value') return null;
  return { addr: e.addr, value: e.next, formula: formulaOf(e.addr) };
};

/** Selection equality check — emit only on actual change. */
export const selectionEquals = (a: State['selection'], b: State['selection']): boolean => {
  return (
    a.active.sheet === b.active.sheet &&
    a.active.row === b.active.row &&
    a.active.col === b.active.col &&
    a.anchor.sheet === b.anchor.sheet &&
    a.anchor.row === b.anchor.row &&
    a.anchor.col === b.anchor.col &&
    a.range.sheet === b.range.sheet &&
    a.range.r0 === b.range.r0 &&
    a.range.c0 === b.range.c0 &&
    a.range.r1 === b.range.r1 &&
    a.range.c1 === b.range.c1
  );
};
