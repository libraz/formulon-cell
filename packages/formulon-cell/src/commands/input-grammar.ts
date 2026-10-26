/**
 * @file Typed-value grammar for cell input: numbers, percents, currency,
 * fractions, times, dates and date-times, per input locale.
 *
 * Every rule here reproduces what desktop Excel stores when the same text is
 * typed into a General cell; `tests/fixtures/excel-input-coercion.json` holds
 * the capture the rules are checked against. A parse returns the value plus the
 * number-format code Excel attaches, written exactly as Excel saves it, so the
 * caller converts it the same way a loaded workbook's format is converted.
 */

import { excelSerialFromParts } from '../format/date-serial.js';

export interface TypedValue {
  readonly value: number;
  /** Format code Excel attaches to the cell; absent when it stays General. */
  readonly formatCode?: string;
}

type DateOrder = 'ymd' | 'mdy' | 'dmy';

interface CurrencyRule {
  readonly position: 'prefix' | 'suffix';
  /** Format code for whole and for fractional amounts. */
  readonly code: readonly [string, string];
}

interface MonthNames {
  /** Full names, lower case, January first. */
  readonly names: readonly string[];
  /** `prefix`: any leading run of 3+ letters; `exact`: full name or first 3 letters. */
  readonly match: 'prefix' | 'exact';
  /** Extra abbreviations that are not a prefix of the full name. */
  readonly aliases?: Readonly<Record<string, number>>;
}

interface InputProfile {
  readonly decimal: '.' | ',';
  readonly groups: readonly string[];
  /** Full-width digits and punctuation are read as their ASCII forms. */
  readonly wideForms: boolean;
  readonly order: DateOrder;
  /** `.` separates date parts as well as `/` and `-`. */
  readonly dotDates: boolean;
  /** `Jan 5` reads the number as the day (`day`) or as a 2-digit year (`year`). */
  readonly monthFirstNumber: 'day' | 'year';
  /** `Jan 5, 2024` is a date. */
  readonly monthDayCommaYear: boolean;
  /** `5. Jan 2024` is a date. */
  readonly dottedDayMonthYear: boolean;
  /** Month names typed input recognises. */
  readonly monthNames: MonthNames;
  /** Calendar prefix Excel puts on month-name formats. */
  readonly monthNamePrefix: '' | 'B1';
  /** Fractional seconds (`12:30:45.5`) are a time. */
  readonly fractionalSeconds: boolean;
  /** Format for a month and day typed as numbers (`1/5`). */
  readonly monthDayCode: string;
  /** Format for a 4-digit year and month typed as numbers (`2024/1`); null = text. */
  readonly yearMonthCode: string | null;
  readonly currencies: Readonly<Record<string, CurrencyRule>>;
  /** Symbols dropped on either side without attaching a currency format. */
  readonly ignoredCurrency: readonly string[];
  readonly cjkDates?: {
    readonly ymd: string;
    readonly ym: string;
    readonly md: string;
  };
  readonly koreanDate?: string;
  readonly japaneseEra?: boolean;
}

const usd = (prefix: string): CurrencyRule => ({
  position: 'prefix',
  code: [
    `${prefix}#,##0_);[Red]\\(${prefix}#,##0\\)`,
    `${prefix}#,##0.00_);[Red]\\(${prefix}#,##0.00\\)`,
  ],
});

const ESCAPED_DOLLAR = usd('\\$');

const ENGLISH_MONTHS = [
  'january',
  'february',
  'march',
  'april',
  'may',
  'june',
  'july',
  'august',
  'september',
  'october',
  'november',
  'december',
];

const COMMON = {
  decimal: '.',
  groups: [','],
  wideForms: false,
  dotDates: false,
  monthFirstNumber: 'day',
  monthDayCommaYear: true,
  dottedDayMonthYear: false,
  monthNames: { names: ENGLISH_MONTHS, match: 'prefix' },
  monthNamePrefix: '',
  fractionalSeconds: true,
  monthDayCode: 'd-mmm',
  yearMonthCode: null,
  ignoredCurrency: ['€'],
} as const;

const CJK_MONTH_DAY = 'm"月"d"日"';
const CJK_DATES = { ymd: 'yyyy"年"m"月"d"日"', ym: 'yyyy"年"m"月"', md: CJK_MONTH_DAY };

const PROFILES: Readonly<Record<string, InputProfile>> = {
  en: { ...COMMON, order: 'mdy', currencies: { $: usd('"$"') } },
  ja: {
    ...COMMON,
    wideForms: true,
    order: 'ymd',
    monthDayCode: CJK_MONTH_DAY,
    yearMonthCode: 'mmm-yy',
    currencies: {
      $: ESCAPED_DOLLAR,
      '¥': {
        position: 'prefix',
        code: ['"¥"#,##0;[Red]"¥"\\-#,##0', '"¥"#,##0.00;[Red]"¥"\\-#,##0.00'],
      },
    },
    cjkDates: CJK_DATES,
    japaneseEra: true,
  },
  zh: {
    ...COMMON,
    wideForms: true,
    order: 'ymd',
    monthDayCode: CJK_MONTH_DAY,
    yearMonthCode: 'mmm-yy',
    currencies: { $: ESCAPED_DOLLAR, '¥': usd('"¥"') },
    cjkDates: CJK_DATES,
  },
  ko: {
    ...COMMON,
    wideForms: true,
    order: 'ymd',
    dotDates: true,
    dottedDayMonthYear: true,
    fractionalSeconds: false,
    monthDayCode: 'm"월"\\ d"일"',
    yearMonthCode: 'mmm-yy',
    currencies: { $: ESCAPED_DOLLAR, '₩': usd('"₩"') },
    koreanDate: 'yyyy"년" m"월" d"일"',
  },
  th: {
    ...COMMON,
    order: 'dmy',
    monthFirstNumber: 'year',
    monthDayCommaYear: false,
    monthNamePrefix: 'B1',
    monthNames: { names: ENGLISH_MONTHS, match: 'exact' },
    currencies: { '฿': usd('"฿"') },
  },
  de: {
    ...COMMON,
    decimal: ',',
    groups: ['.'],
    order: 'dmy',
    dotDates: true,
    monthFirstNumber: 'year',
    monthDayCommaYear: false,
    dottedDayMonthYear: true,
    fractionalSeconds: false,
    monthNames: {
      names: [
        'januar',
        'februar',
        'märz',
        'april',
        'mai',
        'juni',
        'juli',
        'august',
        'september',
        'oktober',
        'november',
        'dezember',
      ],
      match: 'prefix',
      aliases: { mrz: 3 },
    },
    currencies: {
      '€': {
        position: 'suffix',
        code: ['#,##0\\ "€";[Red]\\-#,##0\\ "€"', '#,##0.00\\ "€";[Red]\\-#,##0.00\\ "€"'],
      },
    },
  },
  fr: {
    ...COMMON,
    decimal: ',',
    groups: [' '],
    order: 'dmy',
    monthFirstNumber: 'year',
    monthDayCommaYear: false,
    fractionalSeconds: false,
    monthNames: {
      names: [
        'janvier',
        'février',
        'mars',
        'avril',
        'mai',
        'juin',
        'juillet',
        'août',
        'septembre',
        'octobre',
        'novembre',
        'décembre',
      ],
      match: 'prefix',
    },
    currencies: {
      '€': {
        position: 'suffix',
        code: [
          '#,##0\\ "€"_);[Red]\\(#,##0\\ "€"\\)',
          '#,##0.00\\ "€"_);[Red]\\(#,##0.00\\ "€"\\)',
        ],
      },
    },
  },
};

/** The input grammar's language key for a BCP47 tag (`en` when unknown). */
export const inputLanguage = (locale: string | undefined): string => {
  const language = (locale?.replace('_', '-').split('-', 1)[0] ?? 'en').toLowerCase();
  return PROFILES[language] ? language : 'en';
};

const profileFor = (locale: string | undefined): InputProfile =>
  PROFILES[inputLanguage(locale)] ?? (PROFILES.en as InputProfile);

/** Fold full-width forms to ASCII where the locale accepts them. The yen
 *  sign keeps its full-width form outside Japanese, where Excel rejects it. */
const foldWideForms = (raw: string, locale: string | undefined): string => {
  if (!profileFor(locale).wideForms) return raw;
  const japanese = inputLanguage(locale) === 'ja';
  return [...raw]
    .map((ch) => {
      const cp = ch.codePointAt(0) ?? 0;
      if (cp === 0xffe5) return japanese ? '¥' : ch;
      if (cp >= 0xff01 && cp <= 0xff5e) return String.fromCodePoint(cp - 0xfee0);
      if (cp === 0x3000) return ' ';
      return ch;
    })
    .join('');
};

/** Japanese and Korean keyboards type the backslash key as the yen / won sign. */
const BACKSLASH_CURRENCY: Readonly<Record<string, string>> = { ja: '¥', ko: '₩' };

const DIGITS = /^\d+$/;

interface PlainNumber {
  readonly value: number;
  readonly grouped: boolean;
  readonly fractional: boolean;
  readonly exponent: boolean;
}

/** Parse a signed decimal with the locale's grouping, rejecting malformed
 *  groups instead of deleting punctuation. */
function parsePlainNumber(raw: string, profile: InputProfile): PlainNumber | null {
  const exponentMatch = raw.match(/^(.+?)[eE]([+-]?\d+)$/);
  let unsigned = exponentMatch?.[1] ?? raw;
  let sign = 1;
  if (unsigned.startsWith('+') || unsigned.startsWith('-')) {
    if (unsigned.startsWith('-')) sign = -1;
    unsigned = unsigned.slice(1);
  }
  if (unsigned === '') return null;
  const decimalAt = unsigned.indexOf(profile.decimal);
  if (decimalAt !== unsigned.lastIndexOf(profile.decimal)) return null;
  const intPart = decimalAt < 0 ? unsigned : unsigned.slice(0, decimalAt);
  const fracPart = decimalAt < 0 ? '' : unsigned.slice(decimalAt + 1);
  if (fracPart !== '' && !DIGITS.test(fracPart)) return null;
  if (intPart === '' && fracPart === '') return null;
  // Comma-decimal locales do not accept a bare leading or trailing mark.
  if (profile.decimal === ',' && decimalAt >= 0 && (intPart === '' || fracPart === '')) {
    return null;
  }
  let digits = intPart;
  let grouped = false;
  if (intPart !== '' && !DIGITS.test(intPart)) {
    const present = profile.groups.filter((g) => intPart.includes(g));
    if (present.length !== 1) return null;
    const groups = intPart.split(present[0] ?? '');
    if (
      groups.length < 2 ||
      !DIGITS.test(groups[0] ?? '') ||
      (groups[0]?.length ?? 0) > 3 ||
      groups.slice(1).some((g) => g.length !== 3 || !DIGITS.test(g))
    ) {
      return null;
    }
    digits = groups.join('');
    grouped = true;
  }
  const exponent = exponentMatch?.[2] ?? '';
  const value =
    sign *
    Number(`${digits || '0'}${fracPart ? `.${fracPart}` : ''}${exponent ? `e${exponent}` : ''}`);
  if (!Number.isFinite(value)) return null;
  return { value, grouped, fractional: fracPart !== '', exponent: exponent !== '' };
}

const SCIENTIFIC = '0.00E+00';

/** Numbers, percents and currency amounts. */
export function parseNumberInput(text: string, locale: string | undefined): TypedValue | null {
  const profile = profileFor(locale);
  let body = foldWideForms(text.trim(), locale);
  let negative = false;
  if (body.startsWith('(') && body.endsWith(')')) {
    negative = true;
    body = body.slice(1, -1).trim();
  }
  let percent = false;
  if (body.endsWith('%')) {
    percent = true;
    body = body.slice(0, -1).replace(/ +$/, '');
  } else if (body.startsWith('%')) {
    percent = true;
    body = body.slice(1).replace(/^ +/, '');
  }

  // Only a plain space may separate a currency symbol from the amount.
  const takePrefix = (symbol: string): boolean => {
    for (const lead of ['', '-']) {
      if (!body.startsWith(`${lead}${symbol}`)) continue;
      body = `${lead}${body.slice(lead.length + symbol.length).replace(/^ +/, '')}`;
      return true;
    }
    return false;
  };
  const takeSuffix = (symbol: string): boolean => {
    if (!body.endsWith(symbol)) return false;
    body = body.slice(0, -symbol.length).replace(/ +$/, '');
    return true;
  };
  let currency: CurrencyRule | null = null;
  const backslash = BACKSLASH_CURRENCY[inputLanguage(locale)];
  for (const [symbol, rule] of Object.entries(profile.currencies)) {
    const typed =
      rule.position === 'prefix' ? [symbol, ...(backslash === symbol ? ['\\'] : [])] : [];
    if (typed.some(takePrefix) || (rule.position === 'suffix' && takeSuffix(symbol))) {
      currency = rule;
      break;
    }
  }
  // A recognised but foreign symbol is dropped and leaves the cell General.
  const ignored =
    currency === null &&
    (profile.ignoredCurrency.some(takePrefix) || profile.ignoredCurrency.some(takeSuffix));
  if ((currency || ignored) && percent) return null;

  const parsed = parsePlainNumber(body, profile);
  if (!parsed) return null;
  let value = negative ? -Math.abs(parsed.value) : parsed.value;
  if (percent) value /= 100;
  if (ignored) return { value };
  if (parsed.exponent) return { value, formatCode: SCIENTIFIC };
  if (percent) return { value, formatCode: parsed.fractional ? '0.00%' : '0%' };
  if (currency) return { value, formatCode: currency.code[parsed.fractional ? 1 : 0] };
  if (parsed.grouped) return { value, formatCode: parsed.fractional ? '#,##0.00' : '#,##0' };
  return { value };
}

/** Mixed fractions such as `1 1/2`. A bare `n/d` is a date, not a fraction. */
function parseFraction(text: string): TypedValue | null {
  const m = text.match(/^([+-]?)(\d+)\s+(\d+)\/(\d+)$/);
  if (!m) return null;
  const denominator = Number(m[4]);
  if (denominator === 0) return null;
  const magnitude = Number(m[2]) + Number(m[3]) / denominator;
  const width = '?'.repeat(Math.min(3, (m[4] ?? '').length));
  return { value: m[1] === '-' ? -magnitude : magnitude, formatCode: `# ${width}/${width}` };
}

const SECONDS_PER_DAY = 86_400;

/** Clock times, elapsed times past 24h, AM/PM times and fractional seconds. */
function parseTime(text: string, profile: InputProfile): TypedValue | null {
  const meridiemMatch = text.match(/^(.*\d)\s+(AM|PM)$/i);
  const clock = meridiemMatch ? (meridiemMatch[1] ?? '') : text;
  const meridiem = meridiemMatch?.[2]?.toUpperCase();

  const fractional = clock.match(/^(\d+):(\d{1,2})(?::(\d{1,2}))?\.(\d+)$/);
  if (fractional) {
    if (!profile.fractionalSeconds || meridiem) return null;
    // `m:ss.0` without hours reads the first field as minutes.
    const hasHours = fractional[3] !== undefined;
    const hours = hasHours ? Number(fractional[1]) : 0;
    const minutes = Number(hasHours ? fractional[2] : fractional[1]);
    const seconds = Number(`${hasHours ? fractional[3] : fractional[2]}.${fractional[4]}`);
    return {
      value: (hours * 3600 + minutes * 60 + seconds) / SECONDS_PER_DAY,
      formatCode: 'mmss.0',
    };
  }

  const m = clock.match(/^(\d+)(?::(\d{1,2}))?(?::(\d{1,2}))?$/);
  if (!m) return null;
  const hasMinutes = m[2] !== undefined;
  const hasSeconds = m[3] !== undefined;
  // An hour on its own is a time only with an AM/PM marker (`1 PM`).
  if (!hasMinutes && !meridiem) return null;
  let hours = Number(m[1]);
  const minutes = hasMinutes ? Number(m[2]) : 0;
  const seconds = hasSeconds ? Number(m[3]) : 0;
  if (meridiem) {
    if (hours > 12) return null;
    hours = meridiem === 'AM' ? hours % 12 : (hours % 12) + 12;
  }
  const value = (hours * 3600 + minutes * 60 + seconds) / SECONDS_PER_DAY;
  // An overflowing minute or second field still yields a time value, but the
  // cell stays on General.
  if (minutes > 59 || seconds > 59) return { value };
  if (!meridiem && hours >= 24) return { value, formatCode: '[h]:mm:ss' };
  const base = hasSeconds ? 'h:mm:ss' : 'h:mm';
  return { value, formatCode: meridiem ? `${base} AM/PM` : base };
}

/** The two-digit-year window: 00–29 → 2000s, 30–99 → 1900s. */
const windowYear = (year: string): number | null => {
  const n = Number(year);
  if (year.length <= 2) return n < 30 ? 2000 + n : 1900 + n;
  return year.length === 4 ? n : null;
};

/** A calendar date as a serial, inside the 1900–9999 range dates may take. */
const dateSerial = (
  year: number | null,
  month: number,
  day: number,
  date1904: boolean,
): number | null => {
  if (year === null || year < 1900 || year > 9999) return null;
  return excelSerialFromParts(year, month, day, { allowFictitious: true, date1904 });
};

const thisYear = (): number => new Date().getFullYear();

/** Month number for a typed month name, case-insensitive. */
const monthFromName = (name: string, months: MonthNames): number | null => {
  const lower = name.toLowerCase();
  const alias = months.aliases?.[lower];
  if (alias !== undefined) return alias;
  const index = months.names.findIndex((full) =>
    months.match === 'exact'
      ? full === lower || full.slice(0, 3) === lower
      : lower.length >= 3 && full.startsWith(lower),
  );
  return index < 0 ? null : index + 1;
};

/** Dates written with numbers and `/`, `-` (and `.` where the locale allows). */
function parseNumericDate(
  text: string,
  profile: InputProfile,
  date1904: boolean,
): TypedValue | null {
  const separator = text.match(/[/.-]/)?.[0];
  if (!separator) return null;
  if (separator === '.' && !profile.dotDates) return null;
  const parts = text.split(separator);
  if (parts.some((part) => !DIGITS.test(part) || part.length > 4)) return null;

  if (parts.length === 3) {
    const [a = '', b = '', c = ''] = parts;
    let serial: number | null;
    if (a.length === 4) serial = dateSerial(Number(a), Number(b), Number(c), date1904);
    else if (profile.order === 'ymd')
      serial = dateSerial(windowYear(a), Number(b), Number(c), date1904);
    else if (profile.order === 'mdy')
      serial = dateSerial(windowYear(c), Number(a), Number(b), date1904);
    else serial = dateSerial(windowYear(c), Number(b), Number(a), date1904);
    return serial === null ? null : { value: serial, formatCode: 'mm-dd-yy' };
  }
  if (parts.length !== 2) return null;
  const [a = '', b = ''] = parts;
  if (a.length === 4) {
    if (!profile.yearMonthCode) return null;
    const serial = dateSerial(Number(a), Number(b), 1, date1904);
    return serial === null ? null : { value: serial, formatCode: profile.yearMonthCode };
  }
  if (a.length > 2 || b.length > 2) return null;
  const first = Number(a);
  const second = Number(b);
  const month = profile.order === 'dmy' ? second : first;
  const day = profile.order === 'dmy' ? first : second;
  const monthDay = dateSerial(thisYear(), month, day, date1904);
  if (monthDay !== null) return { value: monthDay, formatCode: profile.monthDayCode };
  // A second field that cannot be the day (m/d) or the month (d/m) is read as
  // a two-digit year: `1/0` is January 2000 in en-US, `12.30` December 1930
  // in de-DE.
  const yearFallback = profile.order === 'mdy' || (profile.order === 'dmy' && second > 12);
  if (!yearFallback) return null;
  const serial = dateSerial(windowYear(b), first, 1, date1904);
  return serial === null ? null : { value: serial, formatCode: 'mmm-yy' };
}

/** Dates that spell the month out. */
function parseMonthNameDate(
  text: string,
  profile: InputProfile,
  date1904: boolean,
): TypedValue | null {
  const prefix = profile.monthNamePrefix;
  const code = (plain: string, prefixed: string): string =>
    prefix ? `${prefix}${prefixed}` : plain;

  let m = text.match(/^(\d{1,2})[-\s](\p{L}+)[-\s](\d{2}|\d{4})$/u);
  if (!m && profile.dottedDayMonthYear) m = text.match(/^(\d{1,2})\.\s(\p{L}+)\s(\d{2}|\d{4})$/u);
  if (m) {
    const month = monthFromName(m[2] ?? '', profile.monthNames);
    const serial =
      month === null ? null : dateSerial(windowYear(m[3] ?? ''), month, Number(m[1]), date1904);
    return serial === null ? null : { value: serial, formatCode: code('d-mmm-yy', 'd\\-mmm\\-yy') };
  }
  if (profile.monthDayCommaYear) {
    m = text.match(/^(\p{L}+)\s(\d{1,2}),\s?(\d{2}|\d{4})$/u);
    if (m) {
      const month = monthFromName(m[1] ?? '', profile.monthNames);
      const serial =
        month === null ? null : dateSerial(windowYear(m[3] ?? ''), month, Number(m[2]), date1904);
      return serial === null ? null : { value: serial, formatCode: 'd-mmm-yy' };
    }
  }
  m = text.match(/^(\d{1,2})[-\s](\p{L}+)$/u);
  if (m) {
    const month = monthFromName(m[2] ?? '', profile.monthNames);
    const serial = month === null ? null : dateSerial(thisYear(), month, Number(m[1]), date1904);
    return serial === null ? null : { value: serial, formatCode: code('d-mmm', 'd\\-mmm') };
  }
  m = text.match(/^(\p{L}+)[-\s](\d{1,2}|\d{4})$/u);
  if (!m) return null;
  const month = monthFromName(m[1] ?? '', profile.monthNames);
  if (month === null) return null;
  if (profile.monthFirstNumber === 'day') {
    if ((m[2] ?? '').length > 2) return null;
    const serial = dateSerial(thisYear(), month, Number(m[2]), date1904);
    return serial === null ? null : { value: serial, formatCode: 'd-mmm' };
  }
  const serial = dateSerial(windowYear(m[2] ?? ''), month, 1, date1904);
  return serial === null ? null : { value: serial, formatCode: code('mmm-yy', 'mmm\\-yy') };
}

const ERA_START: Readonly<Record<string, number>> = {
  明治: 1868,
  大正: 1912,
  昭和: 1926,
  平成: 1989,
  令和: 2019,
  M: 1868,
  T: 1912,
  S: 1926,
  H: 1989,
  R: 2019,
};

const eraSerial = (
  era: string,
  eraYear: number,
  month: number,
  day: number,
  date1904: boolean,
): number | null =>
  eraYear < 1 ? null : dateSerial((ERA_START[era] ?? 0) + eraYear - 1, month, day, date1904);

/** Dates written with CJK year/month/day markers or a Japanese era. */
function parseCjkDate(text: string, profile: InputProfile, date1904: boolean): TypedValue | null {
  const cjk = profile.cjkDates;
  if (cjk) {
    let m = text.match(/^(\d{1,4})年(\d{1,2})月(\d{1,2})日$/);
    if (m) {
      const serial = dateSerial(windowYear(m[1] ?? ''), Number(m[2]), Number(m[3]), date1904);
      return serial === null ? null : { value: serial, formatCode: cjk.ymd };
    }
    m = text.match(/^(\d{1,4})年(\d{1,2})月$/);
    if (m) {
      const serial = dateSerial(windowYear(m[1] ?? ''), Number(m[2]), 1, date1904);
      return serial === null ? null : { value: serial, formatCode: cjk.ym };
    }
    m = text.match(/^(\d{1,2})月(\d{1,2})日$/);
    if (m) {
      const serial = dateSerial(thisYear(), Number(m[1]), Number(m[2]), date1904);
      return serial === null ? null : { value: serial, formatCode: cjk.md };
    }
  }
  if (profile.koreanDate) {
    const m = text.match(/^(\d{4})년\s(\d{1,2})월\s(\d{1,2})일$/);
    if (m) {
      const serial = dateSerial(Number(m[1]), Number(m[2]), Number(m[3]), date1904);
      return serial === null ? null : { value: serial, formatCode: profile.koreanDate };
    }
  }
  if (profile.japaneseEra) {
    let m = text.match(/^(明治|大正|昭和|平成|令和)(\d{1,2}|元)年(\d{1,2})月(\d{1,2})日$/);
    if (m) {
      const eraYear = m[2] === '元' ? 1 : Number(m[2]);
      const serial = eraSerial(m[1] ?? '', eraYear, Number(m[3]), Number(m[4]), date1904);
      return serial === null ? null : { value: serial, formatCode: '[$-411]ggge"年"m"月"d"日"' };
    }
    m = text.match(/^([MTSHR])(\d{1,2})([/.-])(\d{1,2})\3(\d{1,2})$/);
    if (m) {
      const serial = eraSerial(m[1] ?? '', Number(m[2]), Number(m[4]), Number(m[5]), date1904);
      return serial === null ? null : { value: serial, formatCode: '[$-411]ge.m.d' };
    }
  }
  return null;
}

/** Times, dates and numeric date + time pairs. */
export function parseDateTimeInput(
  text: string,
  locale: string | undefined,
  date1904 = false,
): TypedValue | null {
  const profile = profileFor(locale);
  const body = foldWideForms(text.trim(), locale);
  const time = parseTime(body, profile);
  if (time) return time;
  const date =
    parseNumericDate(body, profile, date1904) ??
    parseMonthNameDate(body, profile, date1904) ??
    parseCjkDate(body, profile, date1904);
  if (date) return date;
  const pair = body.match(/^(\S+)\s+(\d.*)$/);
  if (!pair) return null;
  const datePart = parseNumericDate(pair[1] ?? '', profile, date1904);
  const timePart = parseTime(pair[2] ?? '', profile);
  if (!datePart || !timePart?.formatCode || timePart.formatCode === 'mmss.0') return null;
  return { value: datePart.value + timePart.value, formatCode: 'm/d/yy h:mm' };
}

/** Every typed-value form, in the order overlapping spellings resolve. */
export function parseTypedValue(
  text: string,
  locale: string | undefined,
  date1904 = false,
): TypedValue | null {
  return (
    parseNumberInput(text, locale) ??
    parseFraction(foldWideForms(text.trim(), locale)) ??
    parseDateTimeInput(text, locale, date1904)
  );
}
