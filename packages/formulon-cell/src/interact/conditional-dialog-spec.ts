// Pure helpers and rule-kind enums for the Conditional Formatting dialog.
// The dialog DOM wiring lives in `conditional-dialog.ts`; this module
// exposes only the shape data so the parent file can focus on layout.

import { colLetter } from '../commands/print.js';
import { parseRangeRef } from '../engine/range-resolver.js';
import type { Range } from '../engine/types.js';
import type { CellFormat, ConditionalRule } from '../store/store.js';

export type RuleKind = ConditionalRule['kind'];
export type CellValueOp = '>' | '<' | '>=' | '<=' | '=' | '<>' | 'between' | 'not-between';
export type DatePeriod = Extract<ConditionalRule, { kind: 'date-occurring' }>['period'];
export type AverageMode = Extract<ConditionalRule, { kind: 'average' }>['mode'];
export type FormatPreset = 'red-fill' | 'yellow-fill' | 'green-fill' | 'red-text' | 'plain';

export const formatPresetPatch = (preset: FormatPreset): Partial<CellFormat> => {
  switch (preset) {
    case 'red-fill':
      return { color: '#9c0006', fill: '#ffc7ce' };
    case 'yellow-fill':
      return { color: '#9c6500', fill: '#ffeb9c' };
    case 'green-fill':
      return { color: '#006100', fill: '#c6efce' };
    case 'red-text':
      return { color: '#c00000' };
    case 'plain':
      return {};
  }
};

/** Render a sheet-local `Range` as A1 ("A1:B3"). */
export const formatRange = (r: Range): string =>
  `${colLetter(r.c0)}${r.r0 + 1}:${colLetter(r.c1)}${r.r1 + 1}`;

/** Parse a single-sheet A1 range. Cross-sheet refs are rejected so the dialog
 *  always operates on the active sheet. Returns `fallback` on bad input. */
export const parseRange = (raw: string, fallback: Range): Range => {
  const parsed = parseRangeRef(raw);
  if (!parsed || parsed.sheetName != null) return fallback;
  return {
    sheet: fallback.sheet,
    r0: parsed.r0,
    c0: parsed.c0,
    r1: parsed.r1,
    c1: parsed.c1,
  };
};
