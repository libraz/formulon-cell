import { parseAddrKey } from '../engine/address.js';
import type { Addr, Range } from '../engine/types.js';
import {
  rangeArea,
  rangeContainsAddr,
  rangesIntersect,
  subtractRange,
} from './selection-geometry.js';
import type { CellValidation, FormatSlice, RangeValidation, State } from './types.js';

/** Rules covering more cells than this are stored by range, not per cell. */
export const MAX_MATERIALIZED_VALIDATION_CELLS = 10_000;

/** Validation governing `addr`: the cell's own rule, else a range-stored one. */
export function validationAt(format: FormatSlice, addr: Addr): CellValidation | undefined {
  const own = format.formats.get(`${addr.sheet}:${addr.row}:${addr.col}`)?.validation;
  if (own) return own;
  return rangeValidationAt(format, addr);
}

/** Range-stored rule covering `addr`, ignoring per-cell rules. */
export function rangeValidationAt(format: FormatSlice, addr: Addr): CellValidation | undefined {
  for (const rule of format.rangeValidations ?? []) {
    if (rule.ranges.some((range) => rangeContainsAddr(range, addr))) return rule.validation;
  }
  return undefined;
}

/** Whether the rule's ranges are too large to hold one format entry per cell. */
export const isWideValidation = (ranges: readonly Range[]): boolean =>
  ranges.reduce((sum, range) => sum + rangeArea(range), 0) > MAX_MATERIALIZED_VALIDATION_CELLS;

/** Rules with `cut` removed from their ranges; rules left empty are dropped.
 *  Returns the input array when nothing intersects. */
export function subtractFromRangeValidations(
  rules: readonly RangeValidation[] | undefined,
  cut: Range,
): RangeValidation[] | undefined {
  if (!rules?.some((rule) => rule.ranges.some((range) => rangesIntersect(range, cut)))) {
    return rules as RangeValidation[] | undefined;
  }
  return rules.flatMap((rule) => {
    const ranges = rule.ranges.flatMap((range) => subtractRange(range, cut));
    return ranges.length > 0 ? [{ ...rule, ranges }] : [];
  });
}

/** Whether a range-stored rule overlaps `range`. */
export const rangeValidationsIntersect = (format: FormatSlice, range: Range): boolean =>
  (format.rangeValidations ?? []).some((rule) =>
    rule.ranges.some((candidate) => rangesIntersect(candidate, range)),
  );

/** Populated cells governed by range-stored rules that have no rule of their
 *  own; blank cells under such a rule are not enumerated. */
export function rangeValidationCells(
  state: Pick<State, 'format' | 'data'>,
  sheet: number,
): Array<{ addr: Addr; validation: CellValidation }> {
  const rules = (state.format.rangeValidations ?? []).filter((rule) =>
    rule.ranges.some((range) => range.sheet === sheet),
  );
  if (rules.length === 0) return [];
  const out: Array<{ addr: Addr; validation: CellValidation }> = [];
  for (const key of state.data.cells.keys()) {
    const addr = parseAddrKey(key);
    if (!addr || addr.sheet !== sheet) continue;
    if (state.format.formats.get(key)?.validation) continue;
    const validation = rangeValidationAt(state.format, addr);
    if (validation) out.push({ addr, validation });
  }
  return out;
}
