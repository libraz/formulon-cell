import { formatNumber } from '../../commands/format.js';
import type { CellValue } from '../../engine/types.js';
import {
  booleanValue,
  nonNegativeInteger,
  positiveInteger,
  readNumber,
  textValue,
} from './coercion.js';
import { escapeRegExp } from './matching.js';
import { FORMULA_VALUE_NUMBER_LITERAL } from './parser.js';

const concatTextValue = (value: CellValue): string =>
  value.kind === 'error' ? value.text : (textValue(value) ?? '');
const searchText = (
  fn: 'SEARCH' | 'FIND',
  needleValue: CellValue,
  haystackValue: CellValue,
  startValue: CellValue | null,
): CellValue => {
  const needle = textValue(needleValue);
  const haystack = textValue(haystackValue);
  if (needle === null || haystack === null) return { kind: 'error', code: 15, text: '#VALUE!' };
  let start = 0;
  if (startValue !== null) {
    if (
      startValue.kind !== 'number' ||
      !Number.isFinite(startValue.value) ||
      startValue.value < 1
    ) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    start = Math.floor(startValue.value) - 1;
  }
  if (start > haystack.length) return { kind: 'error', code: 15, text: '#VALUE!' };
  const searchNeedle = fn === 'SEARCH' ? needle.toLocaleLowerCase() : needle;
  const searchHaystack = fn === 'SEARCH' ? haystack.toLocaleLowerCase() : haystack;
  const literalNeedle = fn === 'SEARCH' ? searchLiteralPattern(searchNeedle) : null;
  if (literalNeedle !== null) {
    const index = searchHaystack.indexOf(literalNeedle, start);
    return index >= 0
      ? { kind: 'number', value: index + 1 }
      : { kind: 'error', code: 15, text: '#VALUE!' };
  }
  const wildcard = fn === 'SEARCH' ? searchWildcardPattern(searchNeedle) : null;
  if (wildcard) {
    for (let index = start; index <= searchHaystack.length; index += 1) {
      if (wildcard.test(searchHaystack.slice(index))) return { kind: 'number', value: index + 1 };
    }
    return { kind: 'error', code: 15, text: '#VALUE!' };
  }
  const index = searchHaystack.indexOf(searchNeedle, start);
  return index >= 0
    ? { kind: 'number', value: index + 1 }
    : { kind: 'error', code: 15, text: '#VALUE!' };
};
const searchLiteralPattern = (criteria: string): string | null => {
  let out = '';
  let hasEscape = false;
  for (let i = 0; i < criteria.length; i += 1) {
    const ch = criteria[i] ?? '';
    if (ch === '~') {
      const next = criteria[i + 1];
      if (next === '*' || next === '?' || next === '~') {
        out += next;
        hasEscape = true;
        i += 1;
        continue;
      }
    }
    if (ch === '*' || ch === '?') return null;
    out += ch;
  }
  return hasEscape ? out : null;
};
const searchWildcardPattern = (criteria: string): RegExp | null => {
  let pattern = '^';
  let hasWildcard = false;
  for (let i = 0; i < criteria.length; i += 1) {
    const ch = criteria[i] ?? '';
    if (ch === '~') {
      const next = criteria[i + 1];
      if (next === '*' || next === '?' || next === '~') {
        pattern += escapeRegExp(next);
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
  return hasWildcard ? new RegExp(pattern, 'u') : null;
};
const sliceText = (
  fn: 'LEFT' | 'RIGHT' | 'MID',
  textValueCell: CellValue,
  startValue: CellValue | null,
  countValue: CellValue,
): CellValue => {
  const text = textValue(textValueCell);
  const count = nonNegativeInteger(countValue);
  if (text === null || count === null) return { kind: 'error', code: 15, text: '#VALUE!' };
  if (fn === 'LEFT') return { kind: 'text', value: text.slice(0, count) };
  if (fn === 'RIGHT') return { kind: 'text', value: count === 0 ? '' : text.slice(-count) };
  if (startValue === null) return { kind: 'error', code: 15, text: '#VALUE!' };
  const start = positiveInteger(startValue);
  if (start === null) return { kind: 'error', code: 15, text: '#VALUE!' };
  return { kind: 'text', value: text.slice(start - 1, start - 1 + count) };
};
const transformText = (
  fn: 'LOWER' | 'UPPER' | 'TRIM' | 'CLEAN' | 'PROPER' | 'ENCODEURL',
  textValueCell: CellValue,
): CellValue => {
  const text = textValue(textValueCell);
  if (text === null) return { kind: 'error', code: 15, text: '#VALUE!' };
  if (fn === 'LOWER') return { kind: 'text', value: text.toLocaleLowerCase() };
  if (fn === 'UPPER') return { kind: 'text', value: text.toLocaleUpperCase() };
  if (fn === 'TRIM') return { kind: 'text', value: text.trim().replace(/ +/g, ' ') };
  if (fn === 'ENCODEURL') return { kind: 'text', value: encodeURIComponent(text) };
  if (fn === 'CLEAN') {
    return {
      kind: 'text',
      value: [...text].filter((char) => char.charCodeAt(0) > 31).join(''),
    };
  }
  return {
    kind: 'text',
    value: text
      .toLocaleLowerCase()
      .replace(
        /(^|[^A-Za-z0-9])([A-Za-z])/g,
        (_match, prefix: string, letter: string) => `${prefix}${letter.toLocaleUpperCase()}`,
      ),
  };
};
const substituteText = (
  textValueCell: CellValue,
  oldTextValue: CellValue,
  newTextValue: CellValue,
  instanceValue: CellValue | null,
): CellValue => {
  const text = textValue(textValueCell);
  const oldText = textValue(oldTextValue);
  const newText = textValue(newTextValue);
  if (text === null || oldText === null || newText === null) {
    return { kind: 'error', code: 15, text: '#VALUE!' };
  }
  if (oldText === '') return { kind: 'text', value: text };
  if (instanceValue === null) return { kind: 'text', value: text.split(oldText).join(newText) };
  const instance = positiveInteger(instanceValue);
  if (instance === null) return { kind: 'error', code: 15, text: '#VALUE!' };
  let seen = 0;
  let offset = 0;
  for (;;) {
    const index = text.indexOf(oldText, offset);
    if (index < 0) return { kind: 'text', value: text };
    seen += 1;
    if (seen === instance) {
      return {
        kind: 'text',
        value: `${text.slice(0, index)}${newText}${text.slice(index + oldText.length)}`,
      };
    }
    offset = index + oldText.length;
  }
};
const replaceText = (
  textValueCell: CellValue,
  startValue: CellValue,
  countValue: CellValue,
  newTextValue: CellValue,
): CellValue => {
  const text = textValue(textValueCell);
  const newText = textValue(newTextValue);
  const start = positiveInteger(startValue);
  const count = nonNegativeInteger(countValue);
  if (text === null || newText === null || start === null || count === null) {
    return { kind: 'error', code: 15, text: '#VALUE!' };
  }
  const index = start - 1;
  return { kind: 'text', value: `${text.slice(0, index)}${newText}${text.slice(index + count)}` };
};
const repeatText = (textValueCell: CellValue, countValue: CellValue): CellValue => {
  const text = textValue(textValueCell);
  const count = nonNegativeInteger(countValue);
  if (text === null || count === null || text.length * count > 32767) {
    return { kind: 'error', code: 15, text: '#VALUE!' };
  }
  return { kind: 'text', value: text.repeat(count) };
};
const beforeAfterText = (
  fn: 'TEXTBEFORE' | 'TEXTAFTER',
  textValueCell: CellValue,
  delimiterValue: CellValue,
  instanceValue: CellValue | null,
  matchModeValue: CellValue | null,
  matchEndValue: CellValue | null,
  ifNotFoundValue: CellValue | null,
): CellValue => {
  const text = textValue(textValueCell);
  const delimiter = textValue(delimiterValue);
  if (text === null || delimiter === null || delimiter === '') {
    return { kind: 'error', code: 15, text: '#VALUE!' };
  }
  const instance = instanceValue === null ? 1 : readNumber(instanceValue);
  const matchMode = matchModeValue === null ? 0 : readNumber(matchModeValue);
  const matchEnd = matchEndValue === null ? 0 : readNumber(matchEndValue);
  if (instance === null || matchMode === null || matchEnd === null || Math.trunc(instance) === 0) {
    return { kind: 'error', code: 15, text: '#VALUE!' };
  }
  const nth = Math.trunc(instance);
  const ignoreCase = Math.trunc(matchMode) === 1;
  if (Math.trunc(matchMode) !== 0 && !ignoreCase) {
    return { kind: 'error', code: 15, text: '#VALUE!' };
  }
  if (Math.trunc(matchEnd) !== 0 && Math.trunc(matchEnd) !== 1) {
    return { kind: 'error', code: 15, text: '#VALUE!' };
  }
  const haystack = ignoreCase ? text.toLocaleLowerCase() : text;
  const needle = ignoreCase ? delimiter.toLocaleLowerCase() : delimiter;
  const matches: number[] = [];
  let offset = 0;
  for (;;) {
    const index = haystack.indexOf(needle, offset);
    if (index < 0) break;
    matches.push(index);
    offset = index + needle.length;
  }
  if (Math.trunc(matchEnd) === 1) {
    if (fn === 'TEXTBEFORE' && nth > 0) matches.push(text.length);
    if (fn === 'TEXTAFTER' && nth < 0) matches.unshift(-delimiter.length);
  }
  const index = nth > 0 ? matches[nth - 1] : matches[matches.length + nth];
  if (index === undefined) {
    return ifNotFoundValue ?? { kind: 'error', code: 6, text: '#N/A' };
  }
  return {
    kind: 'text',
    value:
      fn === 'TEXTBEFORE'
        ? text.slice(0, index)
        : text.slice(Math.max(0, index + delimiter.length)),
  };
};
const exactText = (leftValue: CellValue, rightValue: CellValue): CellValue => {
  const left = textValue(leftValue);
  const right = textValue(rightValue);
  if (left === null || right === null) return { kind: 'error', code: 15, text: '#VALUE!' };
  return { kind: 'bool', value: left === right };
};
const formatText = (valueCell: CellValue, patternCell: CellValue): CellValue => {
  const value = readNumber(valueCell);
  const pattern = textValue(patternCell);
  if (value === null || pattern === null || pattern === '') {
    return { kind: 'error', code: 15, text: '#VALUE!' };
  }
  return { kind: 'text', value: formatNumber(value, { kind: 'custom', pattern }) };
};
const fixedFormatText = (
  fn: 'DOLLAR' | 'FIXED',
  valueCell: CellValue,
  decimalsCell: CellValue | null,
  noCommasCell: CellValue | null,
): CellValue => {
  const value = readNumber(valueCell);
  const decimalsValue = decimalsCell === null ? 2 : readNumber(decimalsCell);
  const noCommas = noCommasCell === null ? false : booleanValue(noCommasCell);
  if (value === null || decimalsValue === null || noCommas === null) {
    return { kind: 'error', code: 15, text: '#VALUE!' };
  }
  const decimals = Math.trunc(decimalsValue);
  const visibleDecimals = Math.max(0, decimals);
  const roundedValue =
    decimals >= 0 ? value : Math.round(value / 10 ** -decimals) * 10 ** -decimals;
  return {
    kind: 'text',
    value: formatNumber(
      roundedValue,
      fn === 'DOLLAR'
        ? { kind: 'currency', decimals: visibleDecimals, symbol: '$' }
        : { kind: 'fixed', decimals: visibleDecimals, thousands: !noCommas },
    ),
  };
};
const parseNumberText = (text: string, decimalSeparator = '.', groupSeparator = ','): CellValue => {
  if (
    decimalSeparator.length !== 1 ||
    groupSeparator.length !== 1 ||
    decimalSeparator === groupSeparator
  ) {
    return { kind: 'error', code: 15, text: '#VALUE!' };
  }
  const isPercent = text.endsWith('%');
  const body = isPercent ? text.slice(0, -1) : text;
  const decimal = escapeRegExp(decimalSeparator);
  const group = escapeRegExp(groupSeparator);
  const pattern = new RegExp(
    `^[+-]?(?:(?:\\d{1,3}(?:${group}\\d{3})+|\\d+)(?:${decimal}\\d*)?|${decimal}\\d+)(?:[eE][+-]?\\d+)?$`,
    'u',
  );
  if (!pattern.test(body)) return { kind: 'error', code: 15, text: '#VALUE!' };
  const normalized = body.split(groupSeparator).join('').replace(decimalSeparator, '.');
  const number = Number(normalized);
  if (!Number.isFinite(number)) return { kind: 'error', code: 15, text: '#VALUE!' };
  return { kind: 'number', value: isPercent ? number / 100 : number };
};
const valueText = (value: CellValue): CellValue => {
  if (value.kind === 'number') return value;
  const text = value.kind === 'text' ? value.value.trim() : value.kind === 'blank' ? '' : null;
  if (text === null || text === '') return { kind: 'error', code: 15, text: '#VALUE!' };
  if (!FORMULA_VALUE_NUMBER_LITERAL.test(text)) {
    return { kind: 'error', code: 15, text: '#VALUE!' };
  }
  return parseNumberText(text);
};
const numberValueText = (
  value: CellValue,
  decimalSeparatorValue: CellValue | null,
  groupSeparatorValue: CellValue | null,
): CellValue => {
  const text = value.kind === 'number' ? String(value.value) : (textValue(value)?.trim() ?? null);
  const decimalSeparator = decimalSeparatorValue === null ? '.' : textValue(decimalSeparatorValue);
  const groupSeparator = groupSeparatorValue === null ? ',' : textValue(groupSeparatorValue);
  if (text === null || text === '' || !decimalSeparator || !groupSeparator) {
    return { kind: 'error', code: 15, text: '#VALUE!' };
  }
  return parseNumberText(text, decimalSeparator, groupSeparator);
};
const valueToText = (value: CellValue, formatValue: CellValue | null): CellValue => {
  const format = formatValue === null ? 0 : readNumber(formatValue);
  if (format === null) return { kind: 'error', code: 15, text: '#VALUE!' };
  const mode = Math.trunc(format);
  if (mode !== 0 && mode !== 1) return { kind: 'error', code: 15, text: '#VALUE!' };
  if (value.kind === 'error') return { kind: 'text', value: value.text };
  if (value.kind === 'blank') return { kind: 'text', value: '' };
  const text = textValue(value);
  if (text === null) return { kind: 'error', code: 15, text: '#VALUE!' };
  return {
    kind: 'text',
    value: mode === 1 && value.kind === 'text' ? `"${text.replace(/"/g, '""')}"` : text,
  };
};
const coerceScalar = (fn: 'N' | 'T', value: CellValue): CellValue => {
  if (value.kind === 'error') return value;
  if (fn === 'N') {
    if (value.kind === 'number') return value;
    if (value.kind === 'bool') return { kind: 'number', value: value.value ? 1 : 0 };
    return { kind: 'number', value: 0 };
  }
  return value.kind === 'text' ? value : { kind: 'text', value: '' };
};

export {
  beforeAfterText,
  coerceScalar,
  concatTextValue,
  exactText,
  fixedFormatText,
  formatText,
  numberValueText,
  repeatText,
  replaceText,
  searchText,
  sliceText,
  substituteText,
  transformText,
  valueText,
  valueToText,
};
