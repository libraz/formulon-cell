import { PivotFilterKind } from '../engine/types.js';

/** The filter menu a pivot filter kind belongs to. */
export type PivotFilterCategory = 'label' | 'value' | 'date';

/** What a filter kind's operand is: text, a number, a date, a pair of either for a range kind,
 *  an item count for Top 10 by items, or nothing for a relative date period. */
export type PivotFilterOperand =
  | 'text'
  | 'textRange'
  | 'number'
  | 'numberRange'
  | 'count'
  | 'date'
  | 'dateRange'
  | 'none';

const LABEL_KINDS: ReadonlySet<PivotFilterKind> = new Set([
  PivotFilterKind.LabelContains,
  PivotFilterKind.LabelBeginsWith,
  PivotFilterKind.LabelEqual,
  PivotFilterKind.LabelNotEqual,
  PivotFilterKind.LabelNotBeginsWith,
  PivotFilterKind.LabelEndsWith,
  PivotFilterKind.LabelNotEndsWith,
  PivotFilterKind.LabelNotContains,
  PivotFilterKind.LabelGreaterThan,
  PivotFilterKind.LabelGreaterThanOrEqual,
  PivotFilterKind.LabelLessThan,
  PivotFilterKind.LabelLessThanOrEqual,
  PivotFilterKind.LabelBetween,
  PivotFilterKind.LabelNotBetween,
]);

const VALUE_KINDS: ReadonlySet<PivotFilterKind> = new Set([
  PivotFilterKind.ValueTopCount,
  PivotFilterKind.ValueGreaterThan,
  PivotFilterKind.ValueBetween,
  PivotFilterKind.ValueEqual,
  PivotFilterKind.ValueNotEqual,
  PivotFilterKind.ValueGreaterThanOrEqual,
  PivotFilterKind.ValueLessThan,
  PivotFilterKind.ValueLessThanOrEqual,
  PivotFilterKind.ValueNotBetween,
  PivotFilterKind.ValueBottomCount,
  PivotFilterKind.ValueTopPercent,
  PivotFilterKind.ValueBottomPercent,
  PivotFilterKind.ValueTopSum,
  PivotFilterKind.ValueBottomSum,
]);

/** The menu `kind` belongs to; every kind that is neither a label nor a value filter is a date
 *  filter, the relative periods included. */
export const pivotFilterCategory = (kind: PivotFilterKind): PivotFilterCategory =>
  LABEL_KINDS.has(kind) ? 'label' : VALUE_KINDS.has(kind) ? 'value' : 'date';

/** The operand `kind` takes. */
export const pivotFilterOperand = (kind: PivotFilterKind): PivotFilterOperand => {
  switch (kind) {
    case PivotFilterKind.LabelBetween:
    case PivotFilterKind.LabelNotBetween:
      return 'textRange';
    case PivotFilterKind.ValueBetween:
    case PivotFilterKind.ValueNotBetween:
      return 'numberRange';
    case PivotFilterKind.DateBetween:
    case PivotFilterKind.DateNotBetween:
      return 'dateRange';
    case PivotFilterKind.ValueTopCount:
    case PivotFilterKind.ValueBottomCount:
      return 'count';
    case PivotFilterKind.DateEqual:
    case PivotFilterKind.DateNotEqual:
    case PivotFilterKind.DateOlderThan:
    case PivotFilterKind.DateOlderThanOrEqual:
    case PivotFilterKind.DateNewerThan:
    case PivotFilterKind.DateNewerThanOrEqual:
      return 'date';
  }
  const category = pivotFilterCategory(kind);
  return category === 'label' ? 'text' : category === 'value' ? 'number' : 'none';
};
