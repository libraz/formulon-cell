// Pure helpers and rule-kind enums for the Conditional Formatting dialog.
// The dialog DOM wiring lives in `conditional-dialog.ts`; this module
// exposes only the shape data so the parent file can focus on layout.

import { parseRangeRef } from '../engine/range-resolver.js';
import type { Range } from '../engine/types.js';
import type { CellFormat, ConditionalRule } from '../store/store.js';

export type RuleKind = ConditionalRule['kind'];
export type CellValueOp = '>' | '<' | '>=' | '<=' | '=' | '<>' | 'between' | 'not-between';
export type DatePeriod = Extract<ConditionalRule, { kind: 'date-occurring' }>['period'];
export type AverageMode = Extract<ConditionalRule, { kind: 'average' }>['mode'];
export type FormatPreset =
  | 'red-fill'
  | 'yellow-fill'
  | 'green-fill'
  | 'light-red-fill'
  | 'red-text'
  | 'red-border'
  | 'custom';

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
    case 'light-red-fill':
      return { fill: '#ffc7ce' };
    case 'red-border':
      return {
        borders: {
          top: { style: 'thin', color: '#ff0000' },
          right: { style: 'thin', color: '#ff0000' },
          bottom: { style: 'thin', color: '#ff0000' },
          left: { style: 'thin', color: '#ff0000' },
        },
      };
    case 'custom':
      return {};
  }
};

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
