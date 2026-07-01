import type { ConditionalRule } from '../store/types.js';
import type { ConditionalFormatInput } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

/** `formulon::cf::RuleType` ordinals. */
const RULE_TYPE = {
  expression: 0,
  cellIs: 1,
  top10: 5,
  aboveAverage: 6,
  containsText: 7,
  containsBlanks: 11,
  notContainsBlanks: 12,
  containsErrors: 13,
  notContainsErrors: 14,
  duplicateValues: 16,
  uniqueValues: 17,
} as const;

/** `formulon::cf::CellIsOperator` ordinals. */
const CELL_IS_OP: Record<string, number> = {
  '<': 0,
  '<=': 1,
  '=': 2,
  '<>': 3,
  '>=': 4,
  '>': 5,
  between: 6,
  'not-between': 7,
};

const sqrefOf = (rule: ConditionalRule): ConditionalFormatInput['sqref'] => [
  {
    firstRow: rule.range.r0,
    firstCol: rule.range.c0,
    lastRow: rule.range.r1,
    lastCol: rule.range.c1,
  },
];

/**
 * Translate a store conditional-format rule into the engine's
 * `addConditionalFormat` input so its predicate and range round-trip through
 * .xlsx. Returns `null` for rules the engine cannot author:
 *
 * - Visual kinds (`color-scale` / `data-bar` / `icon-set`) — the engine rejects
 *   creating their visual sub-specs (they still round-trip verbatim when read
 *   from an imported file).
 * - `date-occurring` — the `timePeriod` ordinal set isn't stable in this API.
 * - `formula` rules that carry a comparator-prefix predicate (`>10`) rather
 *   than a full `=`-expression.
 *
 * NOTE: the *applied* differential format (fill / font / border in `rule.apply`)
 * is referenced in OOXML by a `dxfId` into the differential-format table. There
 * is no TS-side dxf-creation API yet, so the translated rule persists its
 * predicate and range but not the formatting it applies.
 */
export function conditionalRuleToEngineInput(rule: ConditionalRule): ConditionalFormatInput | null {
  const sqref = sqrefOf(rule);
  switch (rule.kind) {
    case 'cell-value': {
      const op = CELL_IS_OP[rule.op];
      if (op === undefined) return null;
      const input: ConditionalFormatInput = {
        sqref,
        type: RULE_TYPE.cellIs,
        op,
        formula1: String(rule.a),
      };
      if ((rule.op === 'between' || rule.op === 'not-between') && rule.b !== undefined) {
        input.formula2 = String(rule.b);
      }
      return input;
    }
    case 'formula': {
      // Only a full `=`-expression is a valid standalone CF expression.
      if (!rule.formula.startsWith('=')) return null;
      return { sqref, type: RULE_TYPE.expression, formula1: rule.formula.slice(1) };
    }
    case 'text-contains':
      return { sqref, type: RULE_TYPE.containsText, text: rule.text };
    case 'duplicates':
      return { sqref, type: RULE_TYPE.duplicateValues };
    case 'unique':
      return { sqref, type: RULE_TYPE.uniqueValues };
    case 'blanks':
      return { sqref, type: RULE_TYPE.containsBlanks };
    case 'non-blanks':
      return { sqref, type: RULE_TYPE.notContainsBlanks };
    case 'errors':
      return { sqref, type: RULE_TYPE.containsErrors };
    case 'no-errors':
      return { sqref, type: RULE_TYPE.notContainsErrors };
    case 'top-bottom':
      return {
        sqref,
        type: RULE_TYPE.top10,
        rank: rule.n,
        percent: rule.percent === true,
        bottom: rule.mode === 'bottom',
      };
    case 'average':
      return {
        sqref,
        type: RULE_TYPE.aboveAverage,
        aboveAverage: rule.mode === 'above' || rule.mode === 'equal-or-above',
        equalAverage: rule.mode === 'equal-or-above' || rule.mode === 'equal-or-below',
      };
    default:
      // color-scale / data-bar / icon-set / date-occurring
      return null;
  }
}

/**
 * Additively write the engine-representable subset of `rules` onto `sheet` so
 * they round-trip through save. Never clears existing engine rules (the store
 * is not hydrated from engine CF on import, so a clear would drop rules authored
 * elsewhere). Returns the counts written and skipped. No-op — `{written:0,
 * skipped:0}` — when the engine can't author CF rules (stub / older builds).
 *
 * The applied differential format is not persisted (see
 * {@link conditionalRuleToEngineInput}); callers wiring this into a save
 * lifecycle should surface that limitation.
 */
export function syncConditionalRulesToEngine(
  wb: WorkbookHandle,
  rules: readonly ConditionalRule[],
  sheet: number,
): { written: number; skipped: number } {
  if (!wb.capabilities.conditionalFormatMutate) return { written: 0, skipped: 0 };
  let written = 0;
  let skipped = 0;
  for (const rule of rules) {
    if (rule.range.sheet !== sheet) continue;
    const input = conditionalRuleToEngineInput(rule);
    if (!input) {
      skipped += 1;
      continue;
    }
    if (wb.addConditionalFormat(sheet, input)) written += 1;
    else skipped += 1;
  }
  return { written, skipped };
}
