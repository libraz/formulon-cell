import type { CellValue } from '../../engine/types.js';
import { FORMULA_NUMBER_LITERAL } from './parser.js';

function compareValues(
  left: CellValue,
  op: '>' | '<' | '>=' | '<=' | '=' | '<>',
  right: CellValue,
): boolean {
  if (left.kind === 'number' && right.kind === 'number') {
    switch (op) {
      case '>':
        return left.value > right.value;
      case '<':
        return left.value < right.value;
      case '>=':
        return left.value >= right.value;
      case '<=':
        return left.value <= right.value;
      case '=':
        return left.value === right.value;
      case '<>':
        return left.value !== right.value;
    }
  }
  if (left.kind === 'error' && right.kind === 'error') {
    return op === '=' ? left.text === right.text : op === '<>' ? left.text !== right.text : false;
  }
  const leftText =
    left.kind === 'text'
      ? left.value
      : left.kind === 'bool'
        ? String(left.value).toUpperCase()
        : null;
  const rightText =
    right.kind === 'text'
      ? right.value
      : right.kind === 'bool'
        ? String(right.value).toUpperCase()
        : null;
  if (leftText === null || rightText === null) return false;
  const leftComparable = leftText.toLocaleLowerCase();
  const rightComparable = rightText.toLocaleLowerCase();
  return op === '='
    ? leftComparable === rightComparable
    : op === '<>'
      ? leftComparable !== rightComparable
      : false;
}

function countIfWildcardPattern(criteria: string): RegExp | null {
  let pattern = '^';
  let hasWildcard = false;
  for (let i = 0; i < criteria.length; i += 1) {
    const ch = criteria[i] ?? '';
    if (ch === '~') {
      const next = criteria[i + 1];
      if (next === '*' || next === '?' || next === '~') {
        pattern += escapeRegExp(next);
        hasWildcard = true;
        i += 1;
      } else {
        pattern += escapeRegExp(ch);
      }
      continue;
    }
    if (ch === '*') {
      pattern += '.*';
      hasWildcard = true;
      continue;
    }
    if (ch === '?') {
      pattern += '.';
      hasWildcard = true;
      continue;
    }
    pattern += escapeRegExp(ch);
  }
  return hasWildcard ? new RegExp(`${pattern}$`, 'iu') : null;
}

function escapeRegExp(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

const matchesCountIfCriteria = (value: CellValue, criteria: CellValue): boolean => {
  if (criteria.kind === 'text') {
    const raw = criteria.value.trim();
    const m = raw.match(/^(>=|<=|<>|>|<|=)?\s*(.*)$/);
    const op = (m?.[1] ?? '=') as '>' | '<' | '>=' | '<=' | '=' | '<>';
    const rhs = m?.[2] ?? raw;
    if (FORMULA_NUMBER_LITERAL.test(rhs)) {
      return value.kind === 'number'
        ? compareValues(value, op, { kind: 'number', value: Number(rhs) })
        : false;
    }
    if (/^true$/i.test(rhs) || /^false$/i.test(rhs)) {
      return compareValues(value, op, { kind: 'bool', value: /^true$/i.test(rhs) });
    }
    if (rhs === '') {
      const blankLike = value.kind === 'blank' || (value.kind === 'text' && value.value === '');
      return op === '=' ? blankLike : op === '<>' ? !blankLike : false;
    }
    const leftText =
      value.kind === 'text'
        ? value.value
        : value.kind === 'bool'
          ? String(value.value).toUpperCase()
          : value.kind === 'error'
            ? value.text
            : null;
    if (leftText === null) return false;
    const wildcard = op === '=' || op === '<>' ? countIfWildcardPattern(rhs) : null;
    if (wildcard) {
      const matched = wildcard.test(leftText);
      return op === '<>' ? !matched : matched;
    }
    const left = leftText.toLocaleLowerCase();
    const right = rhs.toLocaleLowerCase();
    switch (op) {
      case '=':
        return left === right;
      case '<>':
        return left !== right;
      case '>':
        return left > right;
      case '<':
        return left < right;
      case '>=':
        return left >= right;
      case '<=':
        return left <= right;
    }
  }
  if (criteria.kind === 'number') {
    return value.kind === 'number' && value.value === criteria.value;
  }
  if (criteria.kind === 'bool') {
    return value.kind === 'bool' && value.value === criteria.value;
  }
  if (criteria.kind === 'blank') {
    return value.kind === 'blank' || (value.kind === 'text' && value.value === '');
  }
  return value.kind === 'error' && value.text === criteria.text;
};
const exactMatchValues = (left: CellValue, right: CellValue, allowWildcard = true): boolean => {
  if (left.kind === 'blank' && right.kind === 'blank') return true;
  if (left.kind === 'number' && right.kind === 'number') return left.value === right.value;
  if (left.kind === 'bool' && right.kind === 'bool') return left.value === right.value;
  if (left.kind === 'text' && right.kind === 'text') {
    const wildcard = allowWildcard ? countIfWildcardPattern(left.value) : null;
    if (wildcard) return wildcard.test(right.value);
    return left.value.toLocaleLowerCase() === right.value.toLocaleLowerCase();
  }
  if (left.kind === 'error' && right.kind === 'error') return left.text === right.text;
  return false;
};
const compareApproxValues = (lookup: CellValue, candidate: CellValue): number | null => {
  if (lookup.kind === 'number' && candidate.kind === 'number') {
    if (!Number.isFinite(lookup.value) || !Number.isFinite(candidate.value)) return null;
    return candidate.value === lookup.value ? 0 : candidate.value < lookup.value ? -1 : 1;
  }
  if (lookup.kind === 'text' && candidate.kind === 'text') {
    const left = candidate.value.toLocaleLowerCase();
    const right = lookup.value.toLocaleLowerCase();
    return left === right ? 0 : left < right ? -1 : 1;
  }
  return null;
};
const approximateMatchIndex = (
  lookup: CellValue,
  values: CellValue[],
  mode: -1 | 1,
): number | null => {
  let bestIndex: number | null = null;
  let previous: CellValue | null = null;
  for (let i = 0; i < values.length; i += 1) {
    const candidate = values[i] as CellValue;
    const comparison = compareApproxValues(lookup, candidate);
    if (comparison === null) return null;
    if (previous) {
      const order = compareApproxValues(previous, candidate);
      if (order === null || (mode === 1 ? order < 0 : order > 0)) return null;
    }
    previous = candidate;
    if (mode === 1 ? comparison <= 0 : comparison >= 0) {
      bestIndex = i;
    }
  }
  return bestIndex;
};
const approximateXmatchIndex = (
  lookup: CellValue,
  values: CellValue[],
  mode: -1 | 1,
): number | null => {
  for (let i = 0; i < values.length; i += 1) {
    const candidate = values[i] as CellValue;
    if (compareApproxValues(lookup, candidate) === null) return null;
    if (i > 0) {
      const previous = values[i - 1] as CellValue;
      const order = compareApproxValues(previous, candidate);
      if (order === null || order < 0) return null;
    }
  }
  let nextSmaller: number | null = null;
  for (let i = 0; i < values.length; i += 1) {
    const candidate = values[i] as CellValue;
    const comparison = compareApproxValues(lookup, candidate) as number;
    if (comparison === 0) return i;
    if (mode === -1 && comparison < 0) nextSmaller = i;
    if (mode === 1 && comparison > 0) return i;
  }
  return mode === -1 ? nextSmaller : null;
};
const isExactLookupMode = (value: CellValue): boolean =>
  (value.kind === 'bool' && !value.value) || (value.kind === 'number' && value.value === 0);
const isApproximateLookupMode = (value: CellValue): boolean =>
  (value.kind === 'bool' && value.value) || (value.kind === 'number' && value.value !== 0);

export {
  approximateMatchIndex,
  approximateXmatchIndex,
  compareValues,
  escapeRegExp,
  exactMatchValues,
  isApproximateLookupMode,
  isExactLookupMode,
  matchesCountIfCriteria,
};
