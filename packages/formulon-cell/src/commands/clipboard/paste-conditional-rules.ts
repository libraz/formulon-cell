import type { Range } from '../../engine/types.js';
import { rangesIntersect } from '../../store/selection-geometry.js';
import type { ConditionalRule, SpreadsheetStore } from '../../store/store.js';
import type { History } from '../history.js';
import { shiftFormulaRefs } from '../refs.js';
import { recordConditionalRulesChange } from '../slice-history.js';
import type { ClipboardSnapshot } from './snapshot.js';

/** The copied conditional-format rules placed on every pasted tile. */
export const translatedConditionalRules = (
  snap: ClipboardSnapshot,
  destination: Range,
): ConditionalRule[] => {
  const rules = snap.conditional ?? [];
  const added: ConditionalRule[] = [];
  for (let r0 = destination.r0; r0 <= destination.r1; r0 += snap.rows) {
    for (let c0 = destination.c0; c0 <= destination.c1; c0 += snap.cols) {
      const dRow = r0 - snap.range.r0;
      const dCol = c0 - snap.range.c0;
      for (const rule of rules) {
        const copy = structuredClone(rule);
        copy.range = {
          sheet: destination.sheet,
          r0: rule.range.r0 + dRow,
          c0: rule.range.c0 + dCol,
          r1: rule.range.r1 + dRow,
          c1: rule.range.c1 + dCol,
        };
        if (copy.kind === 'formula') copy.formula = shiftFormulaRefs(copy.formula, dRow, dCol);
        added.push(copy);
      }
    }
  }
  return added;
};

/** All, Merging Conditional Formats: the copied rules land on every pasted
 *  tile ahead of the destination's own rules, which stay (Excel lists a
 *  merged cell's rules source first). */
export const pasteMergedConditionalRules = (
  store: SpreadsheetStore,
  history: History | null,
  snap: ClipboardSnapshot,
  destination: Range,
): void => {
  const added = translatedConditionalRules(snap, destination);
  if (added.length === 0) return;
  recordConditionalRulesChange(history, store, () => {
    store.setState((s) => ({
      ...s,
      conditional: { ...s.conditional, rules: [...added, ...s.conditional.rules] },
    }));
  });
};

/** `rule` with `cut` taken out of its range: up to four rules over what is
 *  left, a formula rule re-anchored to each piece's top-left. */
const subtractFromRule = (rule: ConditionalRule, cut: Range): ConditionalRule[] => {
  const r = rule.range;
  if (r.sheet !== cut.sheet || !rangesIntersect(r, cut)) return [rule];
  const rows0 = Math.max(r.r0, cut.r0);
  const rows1 = Math.min(r.r1, cut.r1);
  const pieces: Range[] = [
    { ...r, r1: cut.r0 - 1 },
    { ...r, r0: cut.r1 + 1 },
    { ...r, r0: rows0, r1: rows1, c1: cut.c0 - 1 },
    { ...r, r0: rows0, r1: rows1, c0: cut.c1 + 1 },
  ].filter((p) => p.r0 <= p.r1 && p.c0 <= p.c1 && p.r0 >= r.r0 && p.r1 <= r.r1);
  return pieces.map((piece) => {
    const copy = structuredClone(rule);
    copy.range = piece;
    if (copy.kind === 'formula') {
      copy.formula = shiftFormulaRefs(copy.formula, piece.r0 - r.r0, piece.c0 - r.c0);
    }
    return copy;
  });
};

/** All and Formats replace the conditional formats of the pasted cells: the
 *  rules lose every `cleared` range (the destination, and a cut's source) and
 *  the copied rules take the destination (an uncovered source leaves none).
 *  Cells outside keep theirs. */
export const replaceConditionalRules = (
  store: SpreadsheetStore,
  history: History | null,
  added: readonly ConditionalRule[],
  cleared: readonly Range[],
): void => {
  const current = store.getState().conditional.rules;
  const touched = current.some((rule) =>
    cleared.some((range) => rangesIntersect(rule.range, range)),
  );
  if (added.length === 0 && !touched) return;
  recordConditionalRulesChange(history, store, () => {
    store.setState((s) => ({
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          ...cleared.reduce<ConditionalRule[]>(
            (rules, range) => rules.flatMap((rule) => subtractFromRule(rule, range)),
            s.conditional.rules,
          ),
          ...added,
        ],
      },
    }));
  });
};
