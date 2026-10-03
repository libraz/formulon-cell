import type { CellValue } from '../../../engine/types.js';
import {
  datedif,
  dateFromSerial,
  dateSerialFromParts,
  dateValueText,
  dayOfYear,
  days360,
  defaultWeekendDays,
  isoWeekNumber,
  networkDays,
  serialTimeFraction,
  timeValueText,
  todaySerial,
  weekdayValue,
  weekendDaysFromValue,
  weekStartForReturnType,
  workday,
  yearFrac,
} from '../calendar.js';
import { readLogical, readNumber } from '../coercion.js';
import type { FormulaDateArg, FormulaOperand } from '../types.js';
import type { FormulaReaderContext } from './context.js';
import type { RangeReader } from './ranges.js';

/** Exactly the reader members this family touches. */
export type DateEvaluatorContext = Pick<FormulaReaderContext, 'readOperand'> &
  Pick<RangeReader, 'numericValuesInFormulaRangeArg'>;

/** Date and time functions: construction and parts, week numbering, EDATE/EOMONTH,
 *  day counts, DATEDIF/YEARFRAC, text parsing, and workday arithmetic. */
export function createDateEvaluator(ctx: DateEvaluatorContext) {
  const { readOperand, numericValuesInFormulaRangeArg } = ctx;
  const dateFunction = (
    fn:
      | 'DATE'
      | 'YEAR'
      | 'MONTH'
      | 'DAY'
      | 'WEEKDAY'
      | 'WEEKNUM'
      | 'ISOWEEKNUM'
      | 'TODAY'
      | 'NOW'
      | 'TIME'
      | 'EDATE'
      | 'EOMONTH'
      | 'DAYS'
      | 'DAYS360'
      | 'DATEDIF'
      | 'YEARFRAC'
      | 'DATEVALUE'
      | 'TIMEVALUE'
      | 'NETWORKDAYS'
      | 'NETWORKDAYS.INTL'
      | 'WORKDAY'
      | 'WORKDAY.INTL'
      | 'HOUR'
      | 'MINUTE'
      | 'SECOND',
    args: FormulaDateArg[],
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    if (fn === 'TODAY') return { kind: 'number', value: todaySerial() };
    if (fn === 'NOW') return { kind: 'number', value: Date.now() / 86_400_000 + 25569 };
    const asOperand = (arg: FormulaDateArg): FormulaOperand | null =>
      arg.kind === 'range' || arg.kind === 'dynamic-range' ? null : arg;
    if (fn === 'DATEVALUE' || fn === 'TIMEVALUE') {
      const [arg] = args;
      if (!arg) return { kind: 'error', code: 15, text: '#VALUE!' };
      const operand = asOperand(arg);
      if (!operand) return { kind: 'error', code: 15, text: '#VALUE!' };
      const value = readOperand(operand, rowOffset, colOffset);
      return fn === 'DATEVALUE' ? dateValueText(value) : timeValueText(value);
    }
    if (fn === 'DATEDIF') {
      const [startArg, endArg, unitArg] = args;
      if (!startArg || !endArg || !unitArg) return { kind: 'error', code: 15, text: '#VALUE!' };
      const startOperand = asOperand(startArg);
      const endOperand = asOperand(endArg);
      const unitOperand = asOperand(unitArg);
      if (!startOperand || !endOperand || !unitOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const start = readNumber(readOperand(startOperand, rowOffset, colOffset));
      const end = readNumber(readOperand(endOperand, rowOffset, colOffset));
      if (start === null || end === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      return datedif(start, end, readOperand(unitOperand, rowOffset, colOffset));
    }
    if (fn === 'DAYS360') {
      const [startArg, endArg, methodArg] = args;
      if (!startArg || !endArg) return { kind: 'error', code: 15, text: '#VALUE!' };
      const startOperand = asOperand(startArg);
      const endOperand = asOperand(endArg);
      const methodOperand = methodArg ? asOperand(methodArg) : undefined;
      if (!startOperand || !endOperand || methodOperand === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const start = readNumber(readOperand(startOperand, rowOffset, colOffset));
      const end = readNumber(readOperand(endOperand, rowOffset, colOffset));
      if (start === null || end === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      const method =
        methodOperand === undefined
          ? false
          : readLogical(readOperand(methodOperand, rowOffset, colOffset));
      if (method === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      return days360(start, end, method);
    }
    if (fn === 'YEARFRAC') {
      const [startArg, endArg, basisArg] = args;
      if (!startArg || !endArg) return { kind: 'error', code: 15, text: '#VALUE!' };
      const startOperand = asOperand(startArg);
      const endOperand = asOperand(endArg);
      const basisOperand = basisArg ? asOperand(basisArg) : undefined;
      if (!startOperand || !endOperand || basisOperand === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const start = readNumber(readOperand(startOperand, rowOffset, colOffset));
      const end = readNumber(readOperand(endOperand, rowOffset, colOffset));
      if (start === null || end === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      const basis =
        basisOperand === undefined
          ? 0
          : readNumber(readOperand(basisOperand, rowOffset, colOffset));
      if (basis === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      return yearFrac(start, end, basis);
    }
    if (
      fn === 'NETWORKDAYS' ||
      fn === 'WORKDAY' ||
      fn === 'NETWORKDAYS.INTL' ||
      fn === 'WORKDAY.INTL'
    ) {
      const [startArg, endOrDaysArg, thirdArg, fourthArg] = args;
      const startOperand = startArg ? asOperand(startArg) : null;
      const endOrDaysOperand = endOrDaysArg ? asOperand(endOrDaysArg) : null;
      if (!startOperand || !endOrDaysOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const start = readNumber(readOperand(startOperand, rowOffset, colOffset));
      const endOrDays = readNumber(readOperand(endOrDaysOperand, rowOffset, colOffset));
      if (start === null || endOrDays === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      const isIntl = fn === 'NETWORKDAYS.INTL' || fn === 'WORKDAY.INTL';
      let weekends = defaultWeekendDays;
      if (isIntl && thirdArg) {
        const weekendOperand = asOperand(thirdArg);
        if (!weekendOperand) return { kind: 'error', code: 15, text: '#VALUE!' };
        const parsedWeekends = weekendDaysFromValue(
          readOperand(weekendOperand, rowOffset, colOffset),
        );
        if (parsedWeekends === null) return { kind: 'error', code: 6, text: '#NUM!' };
        weekends = parsedWeekends;
      }
      const holidaysArg = isIntl ? fourthArg : thirdArg;
      const holidays = new Set<number>();
      if (holidaysArg) {
        if (holidaysArg.kind === 'range' || holidaysArg.kind === 'dynamic-range') {
          const values = numericValuesInFormulaRangeArg(holidaysArg, rowOffset, colOffset);
          if (values === null) return { kind: 'error', code: 15, text: '#VALUE!' };
          for (const value of values) holidays.add(Math.trunc(value));
        } else {
          const value = readNumber(readOperand(holidaysArg, rowOffset, colOffset));
          if (value === null) return { kind: 'error', code: 15, text: '#VALUE!' };
          holidays.add(Math.trunc(value));
        }
      }
      return {
        kind: 'number',
        value:
          fn === 'NETWORKDAYS' || fn === 'NETWORKDAYS.INTL'
            ? networkDays(start, endOrDays, holidays, weekends)
            : workday(start, endOrDays, holidays, weekends),
      };
    }
    const operandArgs = args as FormulaOperand[];
    const values = operandArgs.map((arg) => readNumber(readOperand(arg, rowOffset, colOffset)));
    if (values.some((value) => value === null)) return { kind: 'error', code: 15, text: '#VALUE!' };
    if (fn === 'TIME') {
      const [hour, minute, second] = values.map((value) => Math.trunc(value as number));
      if ((hour as number) < 0 || (minute as number) < 0 || (second as number) < 0) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      const totalSeconds = (hour as number) * 3600 + (minute as number) * 60 + (second as number);
      return { kind: 'number', value: (totalSeconds % 86_400) / 86_400 };
    }
    if (fn === 'DATE') {
      const [year, month, day] = values.map((value) => Math.trunc(value as number));
      return {
        kind: 'number',
        value: dateSerialFromParts(year as number, month as number, day as number),
      };
    }
    if (fn === 'HOUR' || fn === 'MINUTE' || fn === 'SECOND') {
      const totalSeconds = Math.round(serialTimeFraction(values[0] as number) * 86_400) % 86_400;
      if (fn === 'HOUR') return { kind: 'number', value: Math.floor(totalSeconds / 3600) };
      if (fn === 'MINUTE') return { kind: 'number', value: Math.floor(totalSeconds / 60) % 60 };
      return { kind: 'number', value: totalSeconds % 60 };
    }
    if (fn === 'DAYS') {
      return { kind: 'number', value: (values[0] as number) - (values[1] as number) };
    }
    const date = dateFromSerial(values[0] as number);
    if (date === null) return { kind: 'error', code: 15, text: '#VALUE!' };
    if (fn === 'EDATE' || fn === 'EOMONTH') {
      const monthOffset = Math.trunc(values[1] as number);
      const year = date.getUTCFullYear();
      const month = date.getUTCMonth() + monthOffset;
      const lastDay = new Date(Date.UTC(year, month + 1, 0)).getUTCDate();
      const day = fn === 'EOMONTH' ? lastDay : Math.min(date.getUTCDate(), lastDay);
      return { kind: 'number', value: dateSerialFromParts(year, month + 1, day) };
    }
    if (fn === 'YEAR') return { kind: 'number', value: date.getUTCFullYear() };
    if (fn === 'MONTH') return { kind: 'number', value: date.getUTCMonth() + 1 };
    if (fn === 'DAY') return { kind: 'number', value: date.getUTCDate() };
    if (fn === 'ISOWEEKNUM') return { kind: 'number', value: isoWeekNumber(date) };
    const returnType = values.length === 2 ? Math.trunc(values[1] as number) : 1;
    if (fn === 'WEEKNUM') {
      if (returnType === 21) return { kind: 'number', value: isoWeekNumber(date) };
      const firstDay = weekStartForReturnType(returnType);
      if (firstDay === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      const janFirst = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
      const offset = (janFirst.getUTCDay() - firstDay + 7) % 7;
      return { kind: 'number', value: Math.floor((dayOfYear(date) + offset - 1) / 7) + 1 };
    }
    const weekday = weekdayValue(date, returnType);
    return weekday === null
      ? { kind: 'error', code: 15, text: '#VALUE!' }
      : { kind: 'number', value: weekday };
  };
  return { dateFunction };
}
