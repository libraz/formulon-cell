import type { CellValue } from '../../engine/types.js';
import { textValue } from './coercion.js';
import { FORMULA_NUMBER_LITERAL } from './parser.js';

const dateSerialFromParts = (year: number, month: number, day: number): number => {
  const normalizedYear = year >= 0 && year < 1900 ? year + 1900 : year;
  const date = new Date(Date.UTC(normalizedYear, month - 1, day));
  if (normalizedYear >= 0 && normalizedYear < 100) date.setUTCFullYear(normalizedYear);
  return date.getTime() / 86_400_000 + 25569;
};
const validatedDateSerial = (year: number, month: number, day: number): number | null => {
  const normalizedYear = year >= 0 && year < 1900 ? year + 1900 : year;
  const date = new Date(Date.UTC(normalizedYear, month - 1, day));
  if (normalizedYear >= 0 && normalizedYear < 100) date.setUTCFullYear(normalizedYear);
  if (
    date.getUTCFullYear() !== normalizedYear ||
    date.getUTCMonth() !== month - 1 ||
    date.getUTCDate() !== day
  ) {
    return null;
  }
  return dateSerialFromParts(year, month, day);
};
const dateValueText = (value: CellValue): CellValue => {
  const text = textValue(value)?.trim();
  if (!text) return { kind: 'error', code: 15, text: '#VALUE!' };
  let match = /^(\d{4})-(\d{1,2})-(\d{1,2})$/u.exec(text);
  if (match) {
    const serial = validatedDateSerial(Number(match[1]), Number(match[2]), Number(match[3]));
    return serial === null
      ? { kind: 'error', code: 15, text: '#VALUE!' }
      : { kind: 'number', value: serial };
  }
  match = /^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/u.exec(text);
  if (match) {
    const rawYear = Number(match[3]);
    const year = rawYear < 100 ? (rawYear < 30 ? rawYear + 2000 : rawYear + 1900) : rawYear;
    const serial = validatedDateSerial(year, Number(match[1]), Number(match[2]));
    return serial === null
      ? { kind: 'error', code: 15, text: '#VALUE!' }
      : { kind: 'number', value: serial };
  }
  return { kind: 'error', code: 15, text: '#VALUE!' };
};
const timeValueText = (value: CellValue): CellValue => {
  const text = textValue(value)?.trim();
  if (!text) return { kind: 'error', code: 15, text: '#VALUE!' };
  const match = /^(\d{1,2})(?::(\d{1,2}))(?::(\d{1,2}))?\s*(AM|PM)?$/iu.exec(text);
  if (!match) return { kind: 'error', code: 15, text: '#VALUE!' };
  let hour = Number(match[1]);
  const minute = Number(match[2]);
  const second = match[3] === undefined ? 0 : Number(match[3]);
  const meridiem = match[4]?.toUpperCase();
  if (meridiem) {
    if (hour < 1 || hour > 12) return { kind: 'error', code: 15, text: '#VALUE!' };
    hour = (hour % 12) + (meridiem === 'PM' ? 12 : 0);
  }
  if (hour > 23 || minute > 59 || second > 59) {
    return { kind: 'error', code: 15, text: '#VALUE!' };
  }
  return { kind: 'number', value: (hour * 3600 + minute * 60 + second) / 86_400 };
};
const todaySerial = (): number => {
  const now = new Date(Date.now());
  return dateSerialFromParts(now.getUTCFullYear(), now.getUTCMonth() + 1, now.getUTCDate());
};
const serialTimeFraction = (serial: number): number => ((serial % 1) + 1) % 1;
const dateFromSerial = (serial: number): Date | null => {
  if (!Number.isFinite(serial)) return null;
  return new Date(Math.trunc(serial) * 86_400_000 - 25569 * 86_400_000);
};
const serialDateParts = (serial: number): { year: number; month: number; day: number } | null => {
  const date = dateFromSerial(serial);
  if (date === null) return null;
  return {
    year: date.getUTCFullYear(),
    month: date.getUTCMonth() + 1,
    day: date.getUTCDate(),
  };
};
const isLastDayOfFebruary = (year: number, month: number, day: number): boolean =>
  month === 2 && day === new Date(Date.UTC(year, 2, 0)).getUTCDate();
const nextMonthStart = (year: number, month: number): { year: number; month: number; day: 1 } =>
  month === 12 ? { year: year + 1, month: 1, day: 1 } : { year, month: month + 1, day: 1 };
const previousMonthParts = (year: number, month: number): { year: number; month: number } =>
  month === 1 ? { year: year - 1, month: 12 } : { year, month: month - 1 };
const daysInMonth = (year: number, month: number): number =>
  new Date(Date.UTC(year, month, 0)).getUTCDate();
const clampedDateSerial = (year: number, month: number, day: number): number =>
  dateSerialFromParts(year, month, Math.min(day, daysInMonth(year, month)));
const defaultWeekendDays = new Set<number>([0, 6]);
const isWeekendSerial = (serial: number, weekends = defaultWeekendDays): boolean => {
  const date = dateFromSerial(serial);
  if (date === null) return false;
  return weekends.has(date.getUTCDay());
};
const weekendDaysFromCode = (code: number): Set<number> | null => {
  const normalized = Math.trunc(code);
  if (normalized >= 1 && normalized <= 7) {
    const first = (normalized + 5) % 7;
    return new Set<number>([first, (first + 1) % 7]);
  }
  if (normalized >= 11 && normalized <= 17) {
    return new Set<number>([normalized - 11]);
  }
  return null;
};
const weekendDaysFromValue = (value: CellValue): Set<number> | null => {
  if (value.kind === 'number') return weekendDaysFromCode(value.value);
  const text = textValue(value)?.trim();
  if (!text) return null;
  if (/^[01]{7}$/.test(text)) {
    const days = new Set<number>();
    for (let index = 0; index < text.length; index += 1) {
      if (text[index] === '1') days.add((index + 1) % 7);
    }
    return days.size === 7 ? null : days;
  }
  if (FORMULA_NUMBER_LITERAL.test(text)) return weekendDaysFromCode(Number(text));
  return null;
};
const isBusinessDay = (
  serial: number,
  holidays: Set<number>,
  weekends = defaultWeekendDays,
): boolean => !isWeekendSerial(serial, weekends) && !holidays.has(Math.trunc(serial));
const networkDays = (
  start: number,
  end: number,
  holidays = new Set<number>(),
  weekends = defaultWeekendDays,
): number => {
  const first = Math.trunc(start);
  const last = Math.trunc(end);
  const direction = first <= last ? 1 : -1;
  let count = 0;
  for (let serial = first; direction > 0 ? serial <= last : serial >= last; serial += direction) {
    if (isBusinessDay(serial, holidays, weekends)) count += direction;
  }
  return count;
};
const workday = (
  start: number,
  days: number,
  holidays = new Set<number>(),
  weekends = defaultWeekendDays,
): number => {
  let remaining = Math.trunc(days);
  let serial = Math.trunc(start);
  const direction = remaining >= 0 ? 1 : -1;
  while (remaining !== 0) {
    serial += direction;
    if (!isBusinessDay(serial, holidays, weekends)) continue;
    remaining -= direction;
  }
  return serial;
};
const days360 = (start: number, end: number, european: boolean): CellValue => {
  const startParts = serialDateParts(start);
  const endParts = serialDateParts(end);
  if (startParts === null || endParts === null) {
    return { kind: 'error', code: 15, text: '#VALUE!' };
  }
  let { year: y1, month: m1, day: d1 } = startParts;
  let { year: y2, month: m2, day: d2 } = endParts;
  if (european) {
    if (d1 === 31) d1 = 30;
    if (d2 === 31) d2 = 30;
  } else {
    if (d1 === 31 || isLastDayOfFebruary(y1, m1, d1)) d1 = 30;
    if (isLastDayOfFebruary(y2, m2, d2)) {
      if (d1 < 30) {
        ({ year: y2, month: m2, day: d2 } = nextMonthStart(y2, m2));
      } else {
        d2 = 30;
      }
    } else if (d2 === 31) {
      if (d1 < 30) {
        ({ year: y2, month: m2, day: d2 } = nextMonthStart(y2, m2));
      } else {
        d2 = 30;
      }
    }
  }
  return { kind: 'number', value: (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1) };
};
const isLeapYear = (year: number): boolean =>
  (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0;
const daysInYear = (year: number): number => (isLeapYear(year) ? 366 : 365);
const actualActualYearFrac = (start: number, end: number): CellValue => {
  const startDate = dateFromSerial(start);
  const endDate = dateFromSerial(end);
  if (startDate === null || endDate === null) {
    return { kind: 'error', code: 15, text: '#VALUE!' };
  }
  const first = Math.trunc(start);
  const last = Math.trunc(end);
  if (first === last) return { kind: 'number', value: 0 };
  if (first > last) {
    const value = actualActualYearFrac(end, start);
    return value.kind === 'number' ? { kind: 'number', value: -value.value } : value;
  }
  const startYear = startDate.getUTCFullYear();
  const endYear = endDate.getUTCFullYear();
  if (startYear === endYear) {
    return { kind: 'number', value: (last - first) / daysInYear(startYear) };
  }
  const nextYearStart = dateSerialFromParts(startYear + 1, 1, 1);
  const endYearStart = dateSerialFromParts(endYear, 1, 1);
  let value = (nextYearStart - first) / daysInYear(startYear);
  for (let year = startYear + 1; year < endYear; year += 1) {
    value += 1;
  }
  value += (last - endYearStart) / daysInYear(endYear);
  return { kind: 'number', value };
};
const yearFrac = (start: number, end: number, basis: number): CellValue => {
  const normalizedBasis = Math.trunc(basis);
  if (normalizedBasis < 0 || normalizedBasis > 4) {
    return { kind: 'error', code: 6, text: '#NUM!' };
  }
  if (normalizedBasis === 0 || normalizedBasis === 4) {
    const value = days360(start, end, normalizedBasis === 4);
    return value.kind === 'number' ? { kind: 'number', value: value.value / 360 } : value;
  }
  const days = Math.trunc(end) - Math.trunc(start);
  if (normalizedBasis === 1) return actualActualYearFrac(start, end);
  return { kind: 'number', value: days / (normalizedBasis === 2 ? 360 : 365) };
};
const datedif = (start: number, end: number, unitValue: CellValue): CellValue => {
  const startParts = serialDateParts(start);
  const endParts = serialDateParts(end);
  const unit = textValue(unitValue)?.trim().toUpperCase();
  if (startParts === null || endParts === null || !unit) {
    return { kind: 'error', code: 15, text: '#VALUE!' };
  }
  const first = Math.trunc(start);
  const last = Math.trunc(end);
  if (first > last) return { kind: 'error', code: 6, text: '#NUM!' };
  const { year: y1, month: m1, day: d1 } = startParts;
  const { year: y2, month: m2, day: d2 } = endParts;
  const anniversaryThisYear = clampedDateSerial(y2, m1, d1);
  const fullYears = y2 - y1 - (anniversaryThisYear > last ? 1 : 0);
  const fullMonths = (y2 - y1) * 12 + (m2 - m1) - (d2 < d1 ? 1 : 0);
  if (unit === 'D') return { kind: 'number', value: last - first };
  if (unit === 'Y') return { kind: 'number', value: fullYears };
  if (unit === 'M') return { kind: 'number', value: fullMonths };
  if (unit === 'YM') return { kind: 'number', value: ((fullMonths % 12) + 12) % 12 };
  if (unit === 'YD') {
    const anniversary =
      anniversaryThisYear <= last ? anniversaryThisYear : clampedDateSerial(y2 - 1, m1, d1);
    return { kind: 'number', value: last - anniversary };
  }
  if (unit === 'MD') {
    if (d2 >= d1) return { kind: 'number', value: d2 - d1 };
    const previous = previousMonthParts(y2, m2);
    return { kind: 'number', value: last - clampedDateSerial(previous.year, previous.month, d1) };
  }
  return { kind: 'error', code: 6, text: '#NUM!' };
};
const dayOfYear = (date: Date): number => {
  const start = Date.UTC(date.getUTCFullYear(), 0, 1);
  return Math.floor((date.getTime() - start) / 86_400_000) + 1;
};
const isoWeekNumber = (date: Date): number => {
  const normalized = new Date(
    Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()),
  );
  const day = normalized.getUTCDay() || 7;
  normalized.setUTCDate(normalized.getUTCDate() + 4 - day);
  const yearStart = new Date(Date.UTC(normalized.getUTCFullYear(), 0, 1));
  return Math.ceil(((normalized.getTime() - yearStart.getTime()) / 86_400_000 + 1) / 7);
};
const weekStartForReturnType = (returnType: number): number | null => {
  if (returnType === 1 || returnType === 17) return 0;
  if (returnType === 2 || returnType === 11) return 1;
  if (returnType >= 12 && returnType <= 16) return returnType - 10;
  return null;
};
const weekdayValue = (date: Date, returnType: number): number | null => {
  const day = date.getUTCDay();
  if (returnType === 1) return day + 1;
  if (returnType === 2) return ((day + 6) % 7) + 1;
  if (returnType === 3) return (day + 6) % 7;
  if (returnType >= 11 && returnType <= 17) {
    const firstDay = returnType === 17 ? 0 : returnType - 10;
    return ((day - firstDay + 7) % 7) + 1;
  }
  return null;
};

export {
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
};
