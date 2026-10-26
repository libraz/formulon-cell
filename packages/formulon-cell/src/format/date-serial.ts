/**
 * Excel serial <-> civil date conversion. The default 1900 system keeps its
 * deliberate serial-60 gap; the 1904 system counts from 1904-01-01 = 0, has no
 * placeholder day and cannot represent earlier dates.
 */

export const MS_PER_DAY = 86_400_000;

const SERIAL_BASE_UTC_MS = Date.UTC(1899, 11, 30);
const SERIAL_BASE_1904_UTC_MS = Date.UTC(1904, 0, 1);
const FIRST_EXCEL_DATE_UTC_MS = Date.UTC(1900, 0, 1);
const MARCH_1_1900_UTC_MS = Date.UTC(1900, 2, 1);

export interface ExcelSerialDateParts {
  readonly year: number;
  readonly month: number;
  readonly day: number;
  /** Serial 60 is Excel's non-Gregorian 1900-02-29 placeholder. */
  readonly fictitious: boolean;
}

export interface ExcelSerialFromPartsOptions {
  /** Permit the explicit Excel-only 1900-02-29 placeholder. */
  readonly allowFictitious?: boolean;
  /** Interpret the serial in the 1904 date system. */
  readonly date1904?: boolean;
}

/** Construct a UTC date without JavaScript's special handling of years 0–99. */
function utcDateFromParts(year: number, month: number, day: number): Date | null {
  if (!Number.isInteger(year) || !Number.isInteger(month) || !Number.isInteger(day)) return null;
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;
  const date = new Date(0);
  date.setUTCHours(0, 0, 0, 0);
  date.setUTCFullYear(year, month - 1, day);
  if (
    date.getUTCFullYear() !== year ||
    date.getUTCMonth() !== month - 1 ||
    date.getUTCDate() !== day
  ) {
    return null;
  }
  return date;
}

/** Convert a Gregorian civil date to an Excel serial (null if unrepresentable). */
export function excelSerialFromParts(
  year: number,
  month: number,
  day: number,
  options: ExcelSerialFromPartsOptions = {},
): number | null {
  const date1904 = options.date1904 === true;
  if (!date1904 && year === 1900 && month === 2 && day === 29) {
    return options.allowFictitious === true ? 60 : null;
  }
  const date = utcDateFromParts(year, month, day);
  if (date === null) return null;
  if (date1904) {
    return date.getTime() >= SERIAL_BASE_1904_UTC_MS
      ? (date.getTime() - SERIAL_BASE_1904_UTC_MS) / MS_PER_DAY
      : null;
  }
  const rawSerial = (date.getTime() - SERIAL_BASE_UTC_MS) / MS_PER_DAY;
  // Excel inserts serial 60 between 1900-02-28 and 1900-03-01. Dates before
  // 1900-01-01 keep the ordinary serial projection used by existing workbooks.
  return date.getTime() >= FIRST_EXCEL_DATE_UTC_MS && date.getTime() < MARCH_1_1900_UTC_MS
    ? rawSerial - 1
    : rawSerial;
}

/** Convert a UTC Date to an Excel serial, retaining its time part. */
export function excelSerialFromDate(date: Date, date1904 = false): number | null {
  if (!(date instanceof Date) || !Number.isFinite(date.getTime())) return null;
  const day = excelSerialFromParts(
    date.getUTCFullYear(),
    date.getUTCMonth() + 1,
    date.getUTCDate(),
    { date1904 },
  );
  if (day === null) return null;
  const midnight = Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate());
  return day + (date.getTime() - midnight) / MS_PER_DAY;
}

/** Convert a serial to civil date parts, exposing serial 60 explicitly. */
export function excelSerialToDateParts(
  serial: number,
  date1904 = false,
): ExcelSerialDateParts | null {
  if (!Number.isFinite(serial)) return null;
  const whole = Math.floor(serial);
  if (date1904) {
    if (whole < 0) return null;
    const d = new Date(SERIAL_BASE_1904_UTC_MS + whole * MS_PER_DAY);
    if (!Number.isFinite(d.getTime())) return null;
    return {
      year: d.getUTCFullYear(),
      month: d.getUTCMonth() + 1,
      day: d.getUTCDate(),
      fictitious: false,
    };
  }
  if (whole === 60) return { year: 1900, month: 2, day: 29, fictitious: true };
  const adjusted = whole > 60 ? whole - 1 : whole;
  const date = new Date(SERIAL_BASE_UTC_MS + (adjusted + (whole > 0 ? 1 : 0)) * MS_PER_DAY);
  if (!Number.isFinite(date.getTime())) return null;
  return {
    year: date.getUTCFullYear(),
    month: date.getUTCMonth() + 1,
    day: date.getUTCDate(),
    fictitious: false,
  };
}

/**
 * Convert a serial to a UTC Date for arithmetic. Serial 60 has no Gregorian
 * Date, so it uses 1900-02-28 as an arithmetic sentinel; callers that expose
 * a date to users must use excelSerialToDateParts instead.
 */
export function excelSerialToDate(serial: number, date1904 = false): Date | null {
  if (!Number.isFinite(serial)) return null;
  const whole = Math.floor(serial);
  if (date1904) {
    if (whole < 0) return null;
    const d = new Date(SERIAL_BASE_1904_UTC_MS + serial * MS_PER_DAY);
    return Number.isFinite(d.getTime()) ? d : null;
  }
  const adjusted = whole === 60 ? serial - 1 : whole > 60 ? serial - 1 : serial;
  const date = new Date(SERIAL_BASE_UTC_MS + (adjusted + (whole > 0 ? 1 : 0)) * MS_PER_DAY);
  return Number.isFinite(date.getTime()) ? date : null;
}

/** Return the native date-input representation, or null for serial 60. */
export function excelSerialToIsoDate(serial: number, date1904 = false): string | null {
  const parts = excelSerialToDateParts(serial, date1904);
  if (parts === null || parts.fictitious) return null;
  return (
    String(parts.year).padStart(4, '0') +
    '-' +
    String(parts.month).padStart(2, '0') +
    '-' +
    String(parts.day).padStart(2, '0')
  );
}
