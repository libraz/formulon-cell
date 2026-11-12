import { colLetter, formatA1Cell, formatA1Range, quoteSheetName } from '../engine/address.js';
import type { Range } from '../engine/types.js';
import { normalizeRange } from '../store/selection-geometry.js';
import type { SpreadsheetStore } from '../store/store.js';

/** Sheet-prefix policy for one pick. `auto` follows the per-dialog latch. */
export type RangePickPrefix = 'auto' | 'always' | 'never';

export interface RangePickFormatOptions {
  prefix?: RangePickPrefix;
  /** Format this range instead of the store selection. */
  range?: Range;
}

export interface RangePickFormatterDeps {
  store: Pick<SpreadsheetStore, 'getState' | 'subscribe'>;
  sheetName: (index: number) => string;
}

export interface RangePickFormatter {
  /** Capture the active sheet as the origin and clear the qualification latch. */
  begin(): void;
  /** Store subscription that also notices a sheet switch made while a field is armed. */
  subscribe(listener: () => void): () => void;
  /** Active cell as `$C$3`. */
  cell(opts?: RangePickFormatOptions): string;
  /** Selection as `$C$3` (single cell) or `$B$2:$D$5`. */
  range(opts?: RangePickFormatOptions): string;
  /** Selection rows as `$3:$3` / `$2:$5`. */
  rows(opts?: RangePickFormatOptions): string;
  /** Selection columns as `$C:$C` / `$B:$D`. */
  cols(opts?: RangePickFormatOptions): string;
}

/**
 * Text a dialog reference field receives from a grid pick. Picks are absolute;
 * they carry a sheet prefix only once the active sheet has differed from the
 * sheet the dialog opened on, and that qualification then holds for the rest
 * of the dialog session.
 */
export const createRangePickFormatter = (deps: RangePickFormatterDeps): RangePickFormatter => {
  const { store, sheetName } = deps;
  let origin = store.getState().data.sheetIndex;
  let latched = false;

  const observe = (pickSheet: number): void => {
    if (store.getState().data.sheetIndex !== origin || pickSheet !== origin) latched = true;
  };
  const selectionRange = (opts?: RangePickFormatOptions): Range => {
    const r = opts?.range ?? store.getState().selection.range;
    return normalizeRange(r);
  };
  const qualify = (sheet: number, body: string, opts?: RangePickFormatOptions): string => {
    const mode = opts?.prefix ?? 'auto';
    if (mode === 'never') return body;
    if (mode === 'auto') {
      observe(sheet);
      if (!latched) return body;
    }
    const name = sheetName(sheet);
    return name ? `${quoteSheetName(name)}!${body}` : body;
  };

  return {
    begin() {
      origin = store.getState().data.sheetIndex;
      latched = false;
    },
    subscribe(listener) {
      return store.subscribe(() => {
        observe(origin);
        listener();
      });
    },
    cell(opts) {
      const a = store.getState().selection.active;
      const sheet = opts?.range?.sheet ?? a.sheet;
      return qualify(sheet, formatA1Cell(a.row, a.col, true), opts);
    },
    range(opts) {
      const r = selectionRange(opts);
      return qualify(r.sheet, formatA1Range(r, { absolute: true }), opts);
    },
    rows(opts) {
      const r = selectionRange(opts);
      return qualify(r.sheet, `$${r.r0 + 1}:$${r.r1 + 1}`, opts);
    },
    cols(opts) {
      const r = selectionRange(opts);
      return qualify(r.sheet, `$${colLetter(r.c0)}:$${colLetter(r.c1)}`, opts);
    },
  };
};
