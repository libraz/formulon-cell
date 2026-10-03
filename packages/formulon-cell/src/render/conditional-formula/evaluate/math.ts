import { colLetter, MAX_COL, MAX_ROW } from '../../../engine/address.js';
import type { CellValue } from '../../../engine/types.js';
import { numericResult, readLogical, readNumber, textValue } from '../coercion.js';
import {
  combination,
  doubleFactorial,
  erf,
  factorial,
  gamma,
  logGamma,
  standardNormalCdf,
  standardNormalPdf,
} from '../distributions.js';
import {
  baseDigits,
  bitOperand,
  engineeringBaseText,
  engineeringBaseValue,
  maxBitValue,
  romanText,
  romanValue,
} from '../numerals.js';
import type { FormulaOperand } from '../types.js';
import type { FormulaReaderContext, NumericFunctionName } from './context.js';

/** Exactly the reader members this family touches. */
export type MathEvaluatorContext = Pick<FormulaReaderContext, 'readOperand'>;

/** Scalar numeric functions: arithmetic, rounding, trigonometry, combinatorics,
 *  base/bit conversions, plus ADDRESS, CODE, TYPE and ERROR.TYPE. Every remaining
 *  argument is coerced to a number up front, after the raw-operand functions. */
export function createMathEvaluator(ctx: MathEvaluatorContext) {
  const { readOperand } = ctx;
  const roundAwayFromZero = (value: number): number =>
    Math.sign(value) * Math.round(Math.abs(value));
  const roundUpAwayFromZero = (value: number): number =>
    Math.sign(value) * Math.ceil(Math.abs(value));
  const gcdPair = (a: number, b: number): number => {
    let x = Math.abs(a);
    let y = Math.abs(b);
    while (y !== 0) {
      const next = x % y;
      x = y;
      y = next;
    }
    return x;
  };
  const mathFunction = (
    fn: NumericFunctionName,
    args: FormulaOperand[],
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const quoteAddressSheet = (sheet: string): string =>
      /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet) ? sheet : `'${sheet.replace(/'/g, "''")}'`;
    if (fn === 'ADDRESS') {
      const [rowOperand, colOperand, absOperand, a1Operand, sheetOperand] = args;
      if (!rowOperand || !colOperand) return { kind: 'error', code: 15, text: '#VALUE!' };
      const rowValue = readNumber(readOperand(rowOperand, rowOffset, colOffset));
      const colValue = readNumber(readOperand(colOperand, rowOffset, colOffset));
      const absValue = absOperand ? readNumber(readOperand(absOperand, rowOffset, colOffset)) : 1;
      const a1Value = a1Operand ? readLogical(readOperand(a1Operand, rowOffset, colOffset)) : true;
      const sheetValue = sheetOperand
        ? textValue(readOperand(sheetOperand, rowOffset, colOffset))
        : '';
      if (
        rowValue === null ||
        colValue === null ||
        absValue === null ||
        a1Value === null ||
        sheetValue === null
      ) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const row = Math.trunc(rowValue);
      const col = Math.trunc(colValue);
      const abs = Math.trunc(absValue);
      if (row < 1 || row > MAX_ROW + 1 || col < 1 || col > MAX_COL + 1 || abs < 1 || abs > 4) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const absoluteCol = abs === 1 || abs === 3;
      const absoluteRow = abs === 1 || abs === 2;
      const ref = a1Value
        ? `${absoluteCol ? '$' : ''}${colLetter(col - 1)}${absoluteRow ? '$' : ''}${row}`
        : `${absoluteRow ? `R${row}` : `R[${row}]`}${absoluteCol ? `C${col}` : `C[${col}]`}`;
      return {
        kind: 'text',
        value: sheetValue === '' ? ref : `${quoteAddressSheet(sheetValue)}!${ref}`,
      };
    }
    if (fn === 'DECIMAL') {
      const [textOperand, radixOperand] = args;
      if (!textOperand || !radixOperand) return { kind: 'error', code: 15, text: '#VALUE!' };
      const text = textValue(readOperand(textOperand, rowOffset, colOffset));
      const radix = readNumber(readOperand(radixOperand, rowOffset, colOffset));
      if (text === null || radix === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      const base = Math.trunc(radix);
      if (base < 2 || base > 36) return { kind: 'error', code: 6, text: '#NUM!' };
      const normalized = text.trim().toUpperCase();
      if (normalized === '') return { kind: 'error', code: 6, text: '#NUM!' };
      let result = 0;
      for (const char of normalized) {
        const digit = baseDigits.indexOf(char);
        if (digit < 0 || digit >= base) return { kind: 'error', code: 6, text: '#NUM!' };
        result = result * base + digit;
      }
      return numericResult(result);
    }
    if (fn === 'BIN2DEC') {
      const [textOperand] = args;
      if (!textOperand) return { kind: 'error', code: 15, text: '#VALUE!' };
      const text = textValue(readOperand(textOperand, rowOffset, colOffset));
      if (text === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      const value = engineeringBaseValue(text, 2);
      return value === null ? { kind: 'error', code: 6, text: '#NUM!' } : numericResult(value);
    }
    if (fn === 'HEX2DEC' || fn === 'OCT2DEC') {
      const [textOperand] = args;
      if (!textOperand) return { kind: 'error', code: 15, text: '#VALUE!' };
      const text = textValue(readOperand(textOperand, rowOffset, colOffset));
      if (text === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      const base = fn === 'HEX2DEC' ? 16 : 8;
      const value = engineeringBaseValue(text, base);
      return value === null ? { kind: 'error', code: 6, text: '#NUM!' } : numericResult(value);
    }
    if (
      fn === 'BIN2HEX' ||
      fn === 'HEX2BIN' ||
      fn === 'BIN2OCT' ||
      fn === 'OCT2BIN' ||
      fn === 'HEX2OCT' ||
      fn === 'OCT2HEX'
    ) {
      const [textOperand, placesOperand] = args;
      if (!textOperand) return { kind: 'error', code: 15, text: '#VALUE!' };
      const text = textValue(readOperand(textOperand, rowOffset, colOffset));
      const rawPlaces =
        placesOperand === undefined
          ? undefined
          : readNumber(readOperand(placesOperand, rowOffset, colOffset));
      if (text === null || rawPlaces === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      const fromBase = fn.startsWith('BIN') ? 2 : fn.startsWith('HEX') ? 16 : 8;
      const toBase = fn.endsWith('BIN') ? 2 : fn.endsWith('HEX') ? 16 : 8;
      const value = engineeringBaseValue(text, fromBase);
      const output =
        value === null
          ? null
          : engineeringBaseText(
              value,
              toBase,
              rawPlaces === undefined ? null : Math.trunc(rawPlaces),
            );
      return output === null
        ? { kind: 'error', code: 6, text: '#NUM!' }
        : { kind: 'text', value: output };
    }
    if (fn === 'ARABIC') {
      const [textOperand] = args;
      if (!textOperand) return { kind: 'error', code: 15, text: '#VALUE!' };
      const text = textValue(readOperand(textOperand, rowOffset, colOffset));
      if (text === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      const value = romanValue(text);
      return value === null
        ? { kind: 'error', code: 15, text: '#VALUE!' }
        : { kind: 'number', value };
    }
    if (fn === 'CODE' || fn === 'UNICODE') {
      const [textOperand] = args;
      if (!textOperand) return { kind: 'error', code: 15, text: '#VALUE!' };
      const text = textValue(readOperand(textOperand, rowOffset, colOffset));
      const value = text?.codePointAt(0);
      return value === undefined
        ? { kind: 'error', code: 15, text: '#VALUE!' }
        : { kind: 'number', value };
    }
    if (fn === 'TYPE') {
      const [valueOperand] = args;
      if (!valueOperand) return { kind: 'error', code: 15, text: '#VALUE!' };
      const value = readOperand(valueOperand, rowOffset, colOffset);
      const type =
        value.kind === 'text' ? 2 : value.kind === 'bool' ? 4 : value.kind === 'error' ? 16 : 1;
      return { kind: 'number', value: type };
    }
    if (fn === 'ERROR.TYPE') {
      const [valueOperand] = args;
      if (!valueOperand) return { kind: 'error', code: 15, text: '#VALUE!' };
      const value = readOperand(valueOperand, rowOffset, colOffset);
      if (value.kind !== 'error') return { kind: 'error', code: 6, text: '#N/A' };
      const errorTypes: Record<string, number> = {
        '#NULL!': 1,
        '#DIV/0!': 2,
        '#VALUE!': 3,
        '#REF!': 4,
        '#NAME?': 5,
        '#NUM!': 6,
        '#N/A': 7,
        '#GETTING_DATA': 8,
        '#SPILL!': 9,
        '#CALC!': 14,
      };
      const type = errorTypes[value.text];
      return type === undefined
        ? { kind: 'error', code: 6, text: '#N/A' }
        : { kind: 'number', value: type };
    }
    const values = args.map((arg) => readNumber(readOperand(arg, rowOffset, colOffset)));
    if (values.some((value) => value === null)) return { kind: 'error', code: 15, text: '#VALUE!' };
    const first = values[0] as number;
    if (fn === 'PI') return { kind: 'number', value: Math.PI };
    if (fn === 'CHAR' || fn === 'UNICHAR') {
      const codePoint = Math.trunc(first);
      const max = fn === 'CHAR' ? 255 : 0x10ffff;
      if (
        codePoint < 1 ||
        codePoint > max ||
        (fn === 'UNICHAR' && codePoint >= 0xd800 && codePoint <= 0xdfff)
      ) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      return { kind: 'text', value: String.fromCodePoint(codePoint) };
    }
    if (fn === 'ABS') return { kind: 'number', value: Math.abs(first) };
    if (fn === 'RADIANS') return { kind: 'number', value: (first * Math.PI) / 180 };
    if (fn === 'DEGREES') return { kind: 'number', value: (first * 180) / Math.PI };
    if (fn === 'SIN') return { kind: 'number', value: Math.sin(first) };
    if (fn === 'COS') return { kind: 'number', value: Math.cos(first) };
    if (fn === 'TAN') return numericResult(Math.tan(first));
    if (fn === 'SEC') return numericResult(1 / Math.cos(first));
    if (fn === 'CSC') {
      const sine = Math.sin(first);
      if (sine === 0) return { kind: 'error', code: 1, text: '#DIV/0!' };
      return numericResult(1 / sine);
    }
    if (fn === 'COT') {
      const tangent = Math.tan(first);
      if (tangent === 0) return { kind: 'error', code: 1, text: '#DIV/0!' };
      return numericResult(1 / tangent);
    }
    if (fn === 'ASIN' || fn === 'ACOS') {
      if (first < -1 || first > 1) return { kind: 'error', code: 6, text: '#NUM!' };
      return { kind: 'number', value: fn === 'ASIN' ? Math.asin(first) : Math.acos(first) };
    }
    if (fn === 'ATAN') return { kind: 'number', value: Math.atan(first) };
    if (fn === 'ACOT') {
      if (first === 0) return { kind: 'number', value: Math.PI / 2 };
      const value = Math.atan(1 / first);
      return { kind: 'number', value: value < 0 ? value + Math.PI : value };
    }
    if (fn === 'ATAN2') {
      const second = values[1] as number;
      if (first === 0 && second === 0) return { kind: 'error', code: 1, text: '#DIV/0!' };
      return { kind: 'number', value: Math.atan2(second, first) };
    }
    if (fn === 'SINH') return numericResult(Math.sinh(first));
    if (fn === 'COSH') return numericResult(Math.cosh(first));
    if (fn === 'TANH') return { kind: 'number', value: Math.tanh(first) };
    if (fn === 'COTH') {
      const tangent = Math.tanh(first);
      if (tangent === 0) return { kind: 'error', code: 1, text: '#DIV/0!' };
      return numericResult(1 / tangent);
    }
    if (fn === 'SECH') return numericResult(1 / Math.cosh(first));
    if (fn === 'CSCH') {
      const sine = Math.sinh(first);
      if (sine === 0) return { kind: 'error', code: 1, text: '#DIV/0!' };
      return numericResult(1 / sine);
    }
    if (fn === 'ASINH') return numericResult(Math.asinh(first));
    if (fn === 'ACOSH') {
      if (first < 1) return { kind: 'error', code: 6, text: '#NUM!' };
      return numericResult(Math.acosh(first));
    }
    if (fn === 'ATANH') {
      if (first <= -1 || first >= 1) return { kind: 'error', code: 6, text: '#NUM!' };
      return numericResult(Math.atanh(first));
    }
    if (fn === 'ACOTH') {
      if (Math.abs(first) <= 1) return { kind: 'error', code: 6, text: '#NUM!' };
      return numericResult(0.5 * Math.log((first + 1) / (first - 1)));
    }
    if (fn === 'EXP') return numericResult(Math.exp(first));
    if (fn === 'LN' || fn === 'LOG10') {
      if (first <= 0) return { kind: 'error', code: 6, text: '#NUM!' };
      return { kind: 'number', value: fn === 'LN' ? Math.log(first) : Math.log10(first) };
    }
    if (fn === 'FISHER') {
      if (first <= -1 || first >= 1) return { kind: 'error', code: 6, text: '#NUM!' };
      return numericResult(0.5 * Math.log((1 + first) / (1 - first)));
    }
    if (fn === 'FISHERINV') {
      const exponent = Math.exp(2 * first);
      return numericResult((exponent - 1) / (exponent + 1));
    }
    if (fn === 'ERF') {
      const upper = values[1] as number | undefined;
      return numericResult(upper === undefined ? erf(first) : erf(upper) - erf(first));
    }
    if (fn === 'ERF.PRECISE') return numericResult(erf(first));
    if (fn === 'ERFC' || fn === 'ERFC.PRECISE') return numericResult(1 - erf(first));
    if (fn === 'GAUSS') return numericResult(standardNormalCdf(first) - 0.5);
    if (fn === 'BASE') {
      const number = Math.trunc(first);
      const radix = Math.trunc(values[1] as number);
      const minLength = Math.trunc((values[2] as number | undefined) ?? 0);
      if (number < 0 || radix < 2 || radix > 36 || minLength < 0) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      let remaining = number;
      let text = '';
      do {
        text = `${baseDigits[remaining % radix] ?? ''}${text}`;
        remaining = Math.trunc(remaining / radix);
      } while (remaining > 0);
      return { kind: 'text', value: text.padStart(minLength, '0') };
    }
    if (fn === 'DEC2BIN') {
      const number = Math.trunc(first);
      const places = values[1] === undefined ? null : Math.trunc(values[1] as number);
      if (number < -512 || number > 511 || (places !== null && places < 0)) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      if (number < 0) return { kind: 'text', value: (number + 1024).toString(2) };
      const text = number.toString(2);
      if (places !== null && text.length > places) return { kind: 'error', code: 6, text: '#NUM!' };
      return { kind: 'text', value: places === null ? text : text.padStart(places, '0') };
    }
    if (fn === 'DEC2HEX' || fn === 'DEC2OCT') {
      const number = Math.trunc(first);
      const places = values[1] === undefined ? null : Math.trunc(values[1] as number);
      const base = fn === 'DEC2HEX' ? 16 : 8;
      const negativeLimit = fn === 'DEC2HEX' ? -(16 ** 9) : -(8 ** 9);
      const positiveLimit = fn === 'DEC2HEX' ? 16 ** 9 - 1 : 8 ** 9 - 1;
      if (number < negativeLimit || number > positiveLimit || (places !== null && places < 0)) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      if (number < 0)
        return { kind: 'text', value: (number + base ** 10).toString(base).toUpperCase() };
      const text = number.toString(base).toUpperCase();
      if (places !== null && text.length > places) return { kind: 'error', code: 6, text: '#NUM!' };
      return { kind: 'text', value: places === null ? text : text.padStart(places, '0') };
    }
    if (fn === 'ROMAN') {
      const number = Math.trunc(first);
      const form = Math.trunc((values[1] as number | undefined) ?? 0);
      if (number < 1 || number > 3999 || form < 0 || form > 4) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      if (form !== 0) return { kind: 'error', code: 15, text: '#VALUE!' };
      return { kind: 'text', value: romanText(number) };
    }
    if (fn === 'DELTA' || fn === 'GESTEP') {
      const second = (values[1] as number | undefined) ?? 0;
      return {
        kind: 'number',
        value: fn === 'DELTA' ? (first === second ? 1 : 0) : first >= second ? 1 : 0,
      };
    }
    if (
      fn === 'BITAND' ||
      fn === 'BITOR' ||
      fn === 'BITXOR' ||
      fn === 'BITLSHIFT' ||
      fn === 'BITRSHIFT'
    ) {
      const left = bitOperand(first);
      const second = values[1] as number;
      if (left === null) return { kind: 'error', code: 6, text: '#NUM!' };
      let result: bigint;
      if (fn === 'BITAND' || fn === 'BITOR' || fn === 'BITXOR') {
        const right = bitOperand(second);
        if (right === null) return { kind: 'error', code: 6, text: '#NUM!' };
        result = fn === 'BITAND' ? left & right : fn === 'BITOR' ? left | right : left ^ right;
      } else {
        const shift = Math.trunc(second);
        if (Math.abs(shift) > 53) return { kind: 'error', code: 6, text: '#NUM!' };
        const amount = BigInt(Math.abs(shift));
        const shiftLeft = fn === 'BITLSHIFT' ? shift >= 0 : shift < 0;
        result = shiftLeft ? left << amount : left >> amount;
      }
      const number = Number(result);
      return number < 0 || number > maxBitValue || !Number.isFinite(number)
        ? { kind: 'error', code: 6, text: '#NUM!' }
        : { kind: 'number', value: number };
    }
    if (fn === 'SQRTPI') {
      if (first < 0) return { kind: 'error', code: 6, text: '#NUM!' };
      return { kind: 'number', value: Math.sqrt(first * Math.PI) };
    }
    if (fn === 'SUMSQ') {
      return {
        kind: 'number',
        value: values.reduce<number>((sum, value) => sum + (value as number) ** 2, 0),
      };
    }
    if (fn === 'INT') return { kind: 'number', value: Math.floor(first) };
    if (fn === 'SIGN') return { kind: 'number', value: Math.sign(first) };
    if (fn === 'GAMMA') {
      const result = gamma(first);
      return result === null ? { kind: 'error', code: 6, text: '#NUM!' } : numericResult(result);
    }
    if (fn === 'GAMMALN' || fn === 'GAMMALN.PRECISE') {
      if (first <= 0) return { kind: 'error', code: 6, text: '#NUM!' };
      return numericResult(logGamma(first));
    }
    if (fn === 'FACT' || fn === 'FACTDOUBLE') {
      const integer = Math.trunc(first);
      if (integer < 0) return { kind: 'error', code: 6, text: '#NUM!' };
      return numericResult(fn === 'FACT' ? factorial(integer) : doubleFactorial(integer));
    }
    if (fn === 'STANDARDIZE') {
      const mean = values[1] as number;
      const standardDeviation = values[2] as number;
      if (standardDeviation <= 0) return { kind: 'error', code: 6, text: '#NUM!' };
      return { kind: 'number', value: (first - mean) / standardDeviation };
    }
    if (fn === 'PHI') return numericResult(standardNormalPdf(first));
    if (fn === 'COMBIN' || fn === 'COMBINA' || fn === 'PERMUT' || fn === 'PERMUTATIONA') {
      const n = Math.trunc(first);
      const k = Math.trunc(values[1] as number);
      if (n < 0 || k < 0) return { kind: 'error', code: 6, text: '#NUM!' };
      if (fn === 'COMBINA') {
        if (n === 0 && k > 0) return { kind: 'error', code: 6, text: '#NUM!' };
        return numericResult(k === 0 ? 1 : combination(n + k - 1, k));
      }
      if (fn === 'PERMUTATIONA') return numericResult(n ** k);
      if (k > n) return { kind: 'error', code: 6, text: '#NUM!' };
      if (fn === 'COMBIN') return numericResult(combination(n, k));
      return numericResult(factorial(n) / factorial(n - k));
    }
    if (fn === 'MULTINOMIAL') {
      const integers = values.map((value) => Math.trunc(value as number));
      if (integers.some((value) => value < 0)) return { kind: 'error', code: 6, text: '#NUM!' };
      const total = integers.reduce((sum, value) => sum + value, 0);
      return numericResult(
        factorial(total) / integers.reduce((product, value) => product * factorial(value), 1),
      );
    }
    if (fn === 'GCD' || fn === 'LCM') {
      const integers = values.map((value) => Math.trunc(value as number));
      if (integers.some((value) => value < 0)) return { kind: 'error', code: 6, text: '#NUM!' };
      if (fn === 'GCD') {
        return { kind: 'number', value: integers.reduce((acc, value) => gcdPair(acc, value), 0) };
      }
      if (integers.some((value) => value === 0)) return { kind: 'number', value: 0 };
      return {
        kind: 'number',
        value: integers.reduce((acc, value) => Math.abs(acc * value) / gcdPair(acc, value), 1),
      };
    }
    if (fn === 'EVEN' || fn === 'ODD') {
      const magnitude = Math.ceil(Math.abs(first));
      const isEven = magnitude % 2 === 0;
      const rounded =
        fn === 'EVEN' ? (isEven ? magnitude : magnitude + 1) : isEven ? magnitude + 1 : magnitude;
      return { kind: 'number', value: Math.sign(first) * rounded };
    }
    if (
      fn === 'CEILING.MATH' ||
      fn === 'FLOOR.MATH' ||
      fn === 'CEILING.PRECISE' ||
      fn === 'FLOOR.PRECISE' ||
      fn === 'ISO.CEILING'
    ) {
      const significance = Math.abs((values[1] as number | undefined) ?? 1);
      const mode =
        fn === 'CEILING.PRECISE' || fn === 'FLOOR.PRECISE' || fn === 'ISO.CEILING'
          ? 0
          : Math.trunc((values[2] as number | undefined) ?? 0);
      if (significance === 0) return { kind: 'number', value: 0 };
      const scaled = Math.abs(first) / significance;
      const rounded =
        fn === 'CEILING.MATH' || fn === 'CEILING.PRECISE' || fn === 'ISO.CEILING'
          ? first < 0 && mode === 0
            ? Math.floor(scaled)
            : Math.ceil(scaled)
          : first < 0 && mode === 0
            ? Math.ceil(scaled)
            : Math.floor(scaled);
      return { kind: 'number', value: Math.sign(first) * rounded * significance };
    }
    if (fn === 'CEILING' || fn === 'FLOOR') {
      const significance = values[1] as number;
      if (significance === 0) return { kind: 'number', value: 0 };
      if (Math.sign(first) !== Math.sign(significance)) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      const scaled = first / significance;
      const rounded = fn === 'CEILING' ? Math.ceil(scaled) : Math.floor(scaled);
      return { kind: 'number', value: rounded * significance };
    }
    if (fn === 'MROUND') {
      const multiple = values[1] as number;
      if (multiple === 0) return { kind: 'number', value: 0 };
      if (Math.sign(first) !== Math.sign(multiple)) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      return {
        kind: 'number',
        value:
          Math.sign(first) * Math.round(Math.abs(first) / Math.abs(multiple)) * Math.abs(multiple),
      };
    }
    if (fn === 'SQRT') {
      return first < 0
        ? { kind: 'error', code: 6, text: '#NUM!' }
        : { kind: 'number', value: Math.sqrt(first) };
    }
    if (fn === 'TRUNC') {
      const digits = Math.trunc((values[1] as number | undefined) ?? 0);
      if (digits >= 0) {
        const factor = 10 ** digits;
        return { kind: 'number', value: Math.trunc(first * factor) / factor };
      }
      const factor = 10 ** -digits;
      return { kind: 'number', value: Math.trunc(first / factor) * factor };
    }
    const second = values[1] as number;
    if (fn === 'QUOTIENT') {
      if (second === 0) return { kind: 'error', code: 1, text: '#DIV/0!' };
      return { kind: 'number', value: Math.trunc(first / second) };
    }
    if (fn === 'MOD') {
      if (second === 0) return { kind: 'error', code: 1, text: '#DIV/0!' };
      return { kind: 'number', value: first - second * Math.floor(first / second) };
    }
    if (fn === 'POWER') {
      const value = first ** second;
      return Number.isFinite(value)
        ? { kind: 'number', value }
        : { kind: 'error', code: 6, text: '#NUM!' };
    }
    if (fn === 'LOG') {
      const base = second ?? 10;
      if (first <= 0 || base <= 0 || base === 1) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      return { kind: 'number', value: Math.log(first) / Math.log(base) };
    }
    const digits = Math.trunc(second);
    if (fn === 'ROUNDDOWN') {
      if (digits >= 0) {
        const factor = 10 ** digits;
        return { kind: 'number', value: Math.trunc(first * factor) / factor };
      }
      const factor = 10 ** -digits;
      return { kind: 'number', value: Math.trunc(first / factor) * factor };
    }
    if (fn === 'ROUNDUP') {
      if (digits >= 0) {
        const factor = 10 ** digits;
        return { kind: 'number', value: roundUpAwayFromZero(first * factor) / factor };
      }
      const factor = 10 ** -digits;
      return { kind: 'number', value: roundUpAwayFromZero(first / factor) * factor };
    }
    if (digits >= 0) {
      const factor = 10 ** digits;
      return { kind: 'number', value: roundAwayFromZero(first * factor) / factor };
    }
    const factor = 10 ** -digits;
    return { kind: 'number', value: roundAwayFromZero(first / factor) * factor };
  };
  const numericPredicate = (
    fn: 'ISEVEN' | 'ISODD',
    value: FormulaOperand,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const number = readNumber(readOperand(value, rowOffset, colOffset));
    if (number === null) return { kind: 'error', code: 15, text: '#VALUE!' };
    const even = Math.abs(Math.trunc(number)) % 2 === 0;
    return { kind: 'bool', value: fn === 'ISEVEN' ? even : !even };
  };
  return { mathFunction, numericPredicate };
}
