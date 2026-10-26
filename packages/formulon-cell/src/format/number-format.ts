import { formatGeneralNumber, formatScientificNumber } from '../engine/value.js';
import type { NumFmt } from '../store/store.js';
import { excelSerialToDate, excelSerialToDateParts } from './date-serial.js';
import { normalizeFormatLocale } from './locale.js';

export function formatNumber(value: number, fmt: NumFmt | undefined, locale = 'en-US'): string {
  if (!Number.isFinite(value)) return String(value);
  if (!fmt || fmt.kind === 'general') return formatGeneralNumber(value, locale);
  if (fmt.kind === 'text') return String(value);
  if (fmt.kind === 'fixed') {
    const negStyle = fmt.negativeStyle ?? 'minus';
    const body = new Intl.NumberFormat(normalizeFormatLocale(locale), {
      style: 'decimal',
      minimumFractionDigits: fmt.decimals,
      maximumFractionDigits: fmt.decimals,
      useGrouping: !!fmt.thousands,
    }).format(Math.abs(value));
    return applyNegative(value, body, '', negStyle);
  }
  if (fmt.kind === 'currency') {
    const symbol = fmt.symbol ?? '$';
    const negStyle = fmt.negativeStyle ?? 'minus';
    const body = new Intl.NumberFormat(normalizeFormatLocale(locale), {
      style: 'decimal',
      minimumFractionDigits: fmt.decimals,
      maximumFractionDigits: fmt.decimals,
      useGrouping: true,
    }).format(Math.abs(value));
    return applyNegative(value, placeCurrencySymbol(body, symbol, locale), '', negStyle);
  }
  if (fmt.kind === 'percent') {
    return new Intl.NumberFormat(normalizeFormatLocale(locale), {
      style: 'percent',
      minimumFractionDigits: fmt.decimals,
      maximumFractionDigits: fmt.decimals,
    }).format(value);
  }
  if (fmt.kind === 'scientific') {
    // Excel's built-in Scientific format is `0.00E+00` — the exponent is
    // zero-padded to at least two digits. `toExponential` emits a bare `e+4`,
    // so re-pad to match.
    return formatScientificNumber(value, fmt.decimals, locale);
  }
  if (fmt.kind === 'accounting') {
    const symbol = fmt.symbol ?? '$';
    const body = new Intl.NumberFormat(normalizeFormatLocale(locale), {
      style: 'decimal',
      minimumFractionDigits: fmt.decimals,
      maximumFractionDigits: fmt.decimals,
      useGrouping: true,
    }).format(Math.abs(value));
    if (value === 0) {
      const placed = placeCurrencySymbol('-', symbol, locale);
      return placed.startsWith(symbol) ? `${symbol} -` : placed;
    }
    const placed = placeCurrencySymbol(body, symbol, locale);
    return value < 0 ? `(${placed})` : `${placed} `;
  }
  if (fmt.kind === 'date') {
    return renderDateTimePattern(value, fmt.pattern, locale);
  }
  if (fmt.kind === 'time') {
    return renderDateTimePattern(value, fmt.pattern, locale);
  }
  if (fmt.kind === 'datetime') {
    return renderDateTimePattern(value, fmt.pattern, locale);
  }
  if (fmt.kind === 'special') {
    return formatSpecialPattern(value, fmt.pattern);
  }
  if (fmt.kind === 'custom') {
    return formatCustomPattern(value, fmt.pattern, locale);
  }
  return String(value);
}

/** Use the locale's currency position and spacing with the cell's chosen symbol. */
function placeCurrencySymbol(body: string, symbol: string, locale: string): string {
  const parts = new Intl.NumberFormat(normalizeFormatLocale(locale), {
    style: 'currency',
    currency: 'USD',
  }).formatToParts(1);
  const first = parts.findIndex((part) => part.type === 'integer');
  let last = first;
  for (let i = first; i < parts.length; i += 1) {
    if (['integer', 'group', 'decimal', 'fraction'].includes(parts[i]?.type ?? '')) last = i;
  }
  const affix = (items: Intl.NumberFormatPart[]): string =>
    items.map((part) => (part.type === 'currency' ? symbol : part.value)).join('');
  return `${affix(parts.slice(0, first))}${body}${affix(parts.slice(last + 1))}`;
}

function applyNegative(
  value: number,
  body: string,
  symbol: string,
  style: 'minus' | 'parens' | 'red' | 'red-parens',
): string {
  const positive = `${symbol}${body}`;
  if (value >= 0) return positive;
  switch (style) {
    case 'parens':
      return `(${symbol}${body})`;
    case 'red':
      return `-${symbol}${body}`; // color is applied at paint time
    case 'red-parens':
      return `(${symbol}${body})`;
    default:
      return `-${symbol}${body}`;
  }
}

const pad2 = (n: number): string => (n < 10 ? `0${n}` : `${n}`);

/** Spreadsheet-style custom format mini-language. Supports section splitting
 *  (pos;neg;zero;text), `0`/`#`/`?` digit placeholders, `.` decimal, `,`
 *  thousands & scaling, `%`, `\\X` escape, `"text"` literals, `[Red]`-style
 *  color tags (stripped — color is applied at paint time), and date tokens
 *  `yyyy`/`yy`/`mmmmm`/`mmmm`/`mmm`/`mm`/`m`/`dddd`/`ddd`/`dd`/`d`/`hh`/`h`/
 *  `ss`/`s` plus `am/pm`. Not exhaustive but covers the patterns spreadsheets
 *  ship in its built-in format codes. */
function formatCustomPattern(value: number, pattern: string, locale: string): string {
  const selected = selectNumberSection(pattern, value);
  // Strip style-only directives. Quoted literals must stay quoted until the
  // numeric/date renderer decides whether a token is active or literal.
  const active = stripFormatDirectives(selected.section);

  // If the section contains date/time tokens, render as date.
  if (/y|m|d|h|s/i.test(stripLiterals(active))) {
    return renderDateTimePattern(value, active, locale);
  }

  return renderNumericPattern(selected.useAbs ? Math.abs(value) : value, active, locale);
}

/** Pick the section of `pattern` that renders `value`. Spreadsheets allow
 *  positive;negative;zero;text; sections carrying a [>n]/[<n]/[=n] condition
 *  are tried first, and the sign-based default applies only when none match.
 *  `useAbs` is set when the chosen section supplies its own sign. */
function selectNumberSection(pattern: string, value: number): { section: string; useAbs: boolean } {
  const sections = splitSections(pattern);

  let active: string | null = null;
  let useAbs = false;
  // First pass: try condition-bearing sections.
  for (const sec of sections) {
    const cond = parseCondition(sec);
    if (!cond) continue;
    if (cond.test(value)) {
      active = cond.body;
      // Condition matched → caller already wrote the comparator into the
      //  literal so we should NOT show a leading minus from `value` itself.
      useAbs = value < 0;
      break;
    }
  }
  const conditional = sections.some((s) => parseCondition(s));
  if (active === null && conditional) {
    // With conditions present, the first unconditioned section takes every
    //  value no condition matched.
    active = sections.find((s) => !parseCondition(s)) ?? sections[0] ?? '';
  }
  if (active === null) {
    // Second pass: classic sign-based selection over sections that carry
    //  no explicit condition.
    const plain = sections.map((s) => (parseCondition(s) ? null : s));
    const pos = plain[0] ?? null;
    const neg = plain[1] ?? null;
    const zero = plain[2] ?? null;
    if (value < 0 && neg) {
      active = neg;
      useAbs = true;
    } else if (value === 0 && zero) {
      active = zero;
    } else {
      active = pos ?? sections[0] ?? '';
    }
  }
  return { section: active, useAbs };
}

/** The desktop spreadsheet's default 56-colour palette, indexed by `[ColorN]`. */
// biome-ignore format: eight palette entries per line
const FORMAT_PALETTE = [
  '#000000', '#ffffff', '#ff0000', '#00ff00', '#0000ff', '#ffff00', '#ff00ff', '#00ffff',
  '#800000', '#008000', '#000080', '#808000', '#800080', '#008080', '#c0c0c0', '#808080',
  '#9999ff', '#993366', '#ffffcc', '#ccffff', '#660066', '#ff8080', '#0066cc', '#ccccff',
  '#000080', '#ff00ff', '#ffff00', '#00ffff', '#800080', '#800000', '#008080', '#0000ff',
  '#00ccff', '#ccffff', '#ccffcc', '#ffff99', '#99ccff', '#ff99cc', '#cc99ff', '#ffcc99',
  '#3366ff', '#33cccc', '#99cc00', '#ffcc00', '#ff9900', '#ff6600', '#666699', '#969696',
  '#003366', '#339966', '#003300', '#333300', '#993300', '#993366', '#333399', '#333333',
];

const FORMAT_COLOR_NAMES: Record<string, string> = {
  black: '#000000',
  blue: '#0000ff',
  cyan: '#00ffff',
  green: '#00ff00',
  magenta: '#ff00ff',
  red: '#ff0000',
  white: '#ffffff',
  yellow: '#ffff00',
};

/** Paint-time parts of a format code that the display text cannot carry. */
export interface NumberFormatPaint {
  /** Colour from a `[Red]` / `[ColorN]` tag in the section that renders the value. */
  color?: string;
  /** `*x` repeat: `char` fills the cell width between the literal text on
   *  one side of it (`lead` before, or `trail` after) and the rest. */
  fill?: { char: string; lead?: string; trail?: string };
}

/** Read the colour tag and `*x` fill of the section that renders `value`. */
export function numberFormatPaint(code: string, value: number): NumberFormatPaint {
  const { section } = selectNumberSection(code, value);
  const out: NumberFormatPaint = {};
  for (const tag of section.matchAll(/\[([^\]]*)\]/g)) {
    const name = (tag[1] ?? '').toLowerCase();
    const indexed = name.match(/^color(\d+)$/);
    const color =
      FORMAT_COLOR_NAMES[name] ?? (indexed ? FORMAT_PALETTE[Number(indexed[1]) - 1] : undefined);
    if (color) {
      out.color = color;
      break;
    }
  }
  const fill = fillSplit(section.replace(/\[[^\]]*\]/g, ''));
  if (fill) out.fill = fill;
  return out;
}

/** Split a section at its unquoted `*x` when the text on one side of it is
 *  pure literal (quotes, escapes, `_x` spacers, plain symbols). */
function fillSplit(section: string): NumberFormatPaint['fill'] | null {
  for (let i = 0; i < section.length; i += 1) {
    const ch = section[i];
    if (ch === '"') {
      i = section.indexOf('"', i + 1);
      if (i < 0) return null;
    } else if (ch === '\\' || ch === '_') {
      i += 1;
    } else if (ch === '*') {
      const char = section[i + 1];
      if (!char) return null;
      const lead = literalText(section.slice(0, i));
      if (lead !== null) return { char, lead };
      const trail = literalText(section.slice(i + 2));
      return trail !== null ? { char, trail } : null;
    }
  }
  return null;
}

/** The text a literal-only fragment of a format section displays, or null
 *  when it holds a placeholder or token. */
function literalText(fragment: string): string | null {
  let out = '';
  for (let i = 0; i < fragment.length; i += 1) {
    const ch = fragment[i] ?? '';
    if (ch === '"') {
      const end = fragment.indexOf('"', i + 1);
      if (end < 0) return null;
      out += fragment.slice(i + 1, end);
      i = end;
    } else if (ch === '\\') {
      out += fragment[i + 1] ?? '';
      i += 1;
    } else if (ch === '_') {
      out += ' ';
      i += 1;
    } else if (/[0#?.,%@/:*a-z]/i.test(ch)) {
      return null;
    } else {
      out += ch;
    }
  }
  return out;
}

function formatSpecialPattern(value: number, pattern: string): string {
  const sections = splitSections(pattern);
  const active =
    sections.find((sec) => {
      const cond = parseCondition(sec);
      return cond ? cond.test(value) : false;
    }) ??
    (value < 0
      ? (sections[1] ?? sections[0])
      : value === 0
        ? (sections[2] ?? sections[0])
        : sections[0]) ??
    '';
  const withoutCondition = parseCondition(active)?.body ?? active;
  const body = normalizeFormatSection(withoutCondition);
  const digits = String(Math.trunc(Math.abs(value)));
  let cursor = digits.length - 1;
  let out = '';
  for (let i = body.length - 1; i >= 0; i -= 1) {
    const ch = body[i] ?? '';
    if (ch === '0' || ch === '#' || ch === '?') {
      if (cursor >= 0) {
        out = digits[cursor] + out;
        cursor -= 1;
      } else if (ch === '0') {
        out = `0${out}`;
      } else if (ch === '?') {
        out = ` ${out}`;
      }
    } else {
      out = ch + out;
    }
  }
  if (cursor >= 0) out = digits.slice(0, cursor + 1) + out;
  return value < 0 ? `-${out}` : out;
}

function normalizeFormatSection(section: string): string {
  return stripFormatDirectives(section).replace(/"([^"]*)"/g, '$1');
}

function stripFormatDirectives(section: string): string {
  return (
    section
      // Locale/currency tags: [$¥-411]#,##0 → ¥#,##0; [$-ja-JP] is locale-only.
      .replace(/\[\$([^\]-]+)(?:-[^\]]+)?\]/g, '$1')
      .replace(/\[\$-[^\]]+\]/g, '')
      // Color tags are a style concern; the formatter returns text only.
      .replace(/\[(?:Red|Green|Blue|Black|White|Yellow|Magenta|Cyan|Color\d+)\]/gi, '')
      // Alignment/fill directives. `_x` reserves one char width; `*x`
      // repeats a fill char. Canvas text output should not show either.
      .replace(/_.|\\ /g, '')
      .replace(/\*./g, '')
  );
}

/** Parse a leading condition tag like `[>100]"big"0` into its predicate and
 *  the remaining body. Returns null when the section has no condition. */
function parseCondition(section: string): { test: (n: number) => boolean; body: string } | null {
  const m = section.match(/^\s*\[(>=|<=|<>|=|>|<)\s*(-?\d+(?:\.\d+)?)\s*\](.*)$/s);
  if (!m) return null;
  const op = m[1] ?? '=';
  const target = Number.parseFloat(m[2] ?? '0');
  const body = m[3] ?? '';
  const test = (n: number): boolean => {
    switch (op) {
      case '>':
        return n > target;
      case '<':
        return n < target;
      case '>=':
        return n >= target;
      case '<=':
        return n <= target;
      case '<>':
        return n !== target;
      default:
        return n === target;
    }
  };
  return { test, body };
}

/** Split a format string on `;` that's not inside `"..."` or `[...]`. */
function splitSections(s: string): string[] {
  const out: string[] = [];
  let buf = '';
  let inQuote = false;
  let inBracket = false;
  for (let i = 0; i < s.length; i += 1) {
    const ch = s[i];
    if (ch === '\\' && i + 1 < s.length) {
      buf += ch + s[i + 1];
      i += 1;
      continue;
    }
    if (!inBracket && ch === '"') {
      inQuote = !inQuote;
      buf += ch;
      continue;
    }
    if (!inQuote && ch === '[') {
      inBracket = true;
      buf += ch;
      continue;
    }
    if (inBracket && ch === ']') {
      inBracket = false;
      buf += ch;
      continue;
    }
    if (!inQuote && !inBracket && ch === ';') {
      out.push(buf);
      buf = '';
      continue;
    }
    buf += ch;
  }
  out.push(buf);
  return out;
}

/** Strip quoted literals and escapes so the remaining string can be probed
 *  for unescaped tokens (e.g. detecting `m` as a month token). */
function stripLiterals(s: string): string {
  return s.replace(/"[^"]*"/g, '').replace(/\\./g, '');
}

function hasUnquotedPercent(s: string): boolean {
  let inQuote = false;
  let inBracket = false;
  for (let i = 0; i < s.length; i += 1) {
    const ch = s[i];
    if (ch === '\\' && i + 1 < s.length) {
      i += 1;
      continue;
    }
    if (!inBracket && ch === '"') {
      inQuote = !inQuote;
      continue;
    }
    if (!inQuote && ch === '[') {
      inBracket = true;
      continue;
    }
    if (inBracket && ch === ']') {
      inBracket = false;
      continue;
    }
    if (!inQuote && !inBracket && ch === '%') return true;
  }
  return false;
}

function renderDateTimePattern(serial: number, pattern: string, locale: string): string {
  const d = excelSerialToDate(serial);
  const parts = excelSerialToDateParts(serial);
  if (!d || !parts) return '';
  const yyyy = parts.year;
  const mm = parts.month;
  const dd = parts.day;
  const dow = d.getUTCDay();
  const hh = d.getUTCHours();
  const mi = d.getUTCMinutes();
  const ss = d.getUTCSeconds();
  // Elapsed-time tokens ([h]/[m]/[s]) accumulate past the 24h/60m wrap and are
  // measured from the serial value directly rather than the wall-clock fields.
  const totalSeconds = Math.round(serial * 86_400);
  const elapsedHours = Math.floor(totalSeconds / 3600);
  const elapsedMinutes = Math.floor(totalSeconds / 60);
  const elapsedSeconds = totalSeconds;
  const has12h = /a\/?p|am\/pm/i.test(pattern);
  const hh12 = ((hh + 11) % 12) + 1;
  const ampm = hh < 12 ? 'AM' : 'PM';
  const DAYS_LONG = Array.from({ length: 7 }, (_, i) =>
    new Intl.DateTimeFormat(normalizeFormatLocale(locale), {
      weekday: 'long',
      timeZone: 'UTC',
    }).format(new Date(Date.UTC(2023, 0, i + 1))),
  );
  const DAYS_SHORT = Array.from({ length: 7 }, (_, i) =>
    new Intl.DateTimeFormat(normalizeFormatLocale(locale), {
      weekday: 'short',
      timeZone: 'UTC',
    }).format(new Date(Date.UTC(2023, 0, i + 1))),
  );
  const MONTHS_LONG = Array.from({ length: 12 }, (_, i) =>
    new Intl.DateTimeFormat(normalizeFormatLocale(locale), {
      month: 'long',
      timeZone: 'UTC',
    }).format(new Date(Date.UTC(2023, i, 1))),
  );
  const MONTHS_SHORT = Array.from({ length: 12 }, (_, i) =>
    new Intl.DateTimeFormat(normalizeFormatLocale(locale), {
      month: 'short',
      timeZone: 'UTC',
    }).format(new Date(Date.UTC(2023, i, 1))),
  );
  let out = '';
  let prevWasH = false;
  for (let i = 0; i < pattern.length; ) {
    const rest = pattern.slice(i);
    // Quoted literal — emit without the quotes.
    if (pattern[i] === '"') {
      const end = pattern.indexOf('"', i + 1);
      if (end < 0) {
        out += pattern.slice(i + 1);
        break;
      }
      out += pattern.slice(i + 1, end);
      i = end + 1;
      continue;
    }
    if (pattern[i] === '\\' && i + 1 < pattern.length) {
      out += pattern[i + 1];
      i += 2;
      continue;
    }
    // Elapsed-time tokens in square brackets: [h]/[hh] total hours, [m]/[mm]
    // total minutes, [s]/[ss] total seconds. They do not wrap at 24h/60.
    if (pattern[i] === '[') {
      const close = pattern.indexOf(']', i + 1);
      if (close > i) {
        const inner = pattern.slice(i + 1, close);
        if (/^h+$/i.test(inner)) {
          out += String(elapsedHours).padStart(inner.length, '0');
          i = close + 1;
          prevWasH = true;
          continue;
        }
        if (/^m+$/i.test(inner)) {
          out += String(elapsedMinutes).padStart(inner.length, '0');
          i = close + 1;
          prevWasH = false;
          continue;
        }
        if (/^s+$/i.test(inner)) {
          out += String(elapsedSeconds).padStart(inner.length, '0');
          i = close + 1;
          prevWasH = false;
          continue;
        }
      }
    }
    // Token matching, longest first. After hours, an "m"/"mm" token is
    //  minutes, not month — track prevWasH to disambiguate.
    let tok = '';
    const lowerRest = rest.toLowerCase();
    if (lowerRest.startsWith('yyyy')) tok = 'yyyy';
    else if (lowerRest.startsWith('yy')) tok = 'yy';
    else if (lowerRest.startsWith('mmmmm')) tok = 'mmmmm';
    else if (lowerRest.startsWith('mmmm')) tok = 'mmmm';
    else if (lowerRest.startsWith('mmm')) tok = 'mmm';
    else if (lowerRest.startsWith('mm')) tok = 'mm';
    else if (lowerRest.startsWith('dddd')) tok = 'dddd';
    else if (lowerRest.startsWith('ddd')) tok = 'ddd';
    else if (lowerRest.startsWith('dd')) tok = 'dd';
    else if (lowerRest.startsWith('hh')) tok = 'hh';
    else if (lowerRest.startsWith('ss')) tok = 'ss';
    else if (/^am\/pm/i.test(rest)) tok = 'am/pm';
    else if (/^a\/p/i.test(rest)) tok = 'a/p';
    else if (lowerRest[0] === 'm') tok = 'm';
    else if (lowerRest[0] === 'd') tok = 'd';
    else if (lowerRest[0] === 'h') tok = 'h';
    else if (lowerRest[0] === 's') tok = 's';

    if (!tok) {
      // Preserve the hour context across time separators (`:`, spaces): in
      // `hh:mm` the `mm` is minutes even though a literal sits between them.
      out += pattern[i];
      i += 1;
      continue;
    }
    switch (tok) {
      case 'yyyy':
        out += String(yyyy);
        break;
      case 'yy':
        out += String(yyyy).slice(-2);
        break;
      case 'mmmmm':
        out += Array.from(MONTHS_LONG[mm - 1] ?? '')[0] ?? '';
        break;
      case 'mmmm':
        out += MONTHS_LONG[mm - 1] ?? '';
        break;
      case 'mmm':
        out += MONTHS_SHORT[mm - 1] ?? '';
        break;
      case 'mm':
        out += prevWasH ? pad2(mi) : pad2(mm);
        break;
      case 'm':
        out += prevWasH ? String(mi) : String(mm);
        break;
      case 'dddd':
        out += DAYS_LONG[dow] ?? '';
        break;
      case 'ddd':
        out += DAYS_SHORT[dow] ?? '';
        break;
      case 'dd':
        out += pad2(dd);
        break;
      case 'd':
        out += String(dd);
        break;
      case 'hh':
        out += pad2(has12h ? hh12 : hh);
        break;
      case 'h':
        out += String(has12h ? hh12 : hh);
        break;
      case 'ss':
        out += pad2(ss);
        break;
      case 's':
        out += String(ss);
        break;
      case 'am/pm':
        out +=
          pattern.slice(i, i + 5) === pattern.slice(i, i + 5).toUpperCase()
            ? ampm
            : ampm.toLowerCase();
        break;
      case 'a/p':
        out +=
          pattern.slice(i, i + 3) === pattern.slice(i, i + 3).toUpperCase()
            ? hh < 12
              ? 'A'
              : 'P'
            : hh < 12
              ? 'a'
              : 'p';
        break;
    }
    prevWasH = tok === 'h' || tok === 'hh';
    i += tok.length;
  }
  return out;
}

/** Best rational approximation of `x` in [0,1) with denominator ≤ maxDen. */
function bestFraction(x: number, maxDen: number): [number, number] {
  let bestNum = 0;
  let bestDen = 1;
  let bestErr = Number.POSITIVE_INFINITY;
  for (let d = 1; d <= maxDen; d += 1) {
    const n = Math.round(x * d);
    const err = Math.abs(x - n / d);
    if (err < bestErr - 1e-12) {
      bestErr = err;
      bestNum = n;
      bestDen = d;
    }
    if (err < 1e-12) break;
  }
  return [bestNum, bestDen];
}

/** Render Excel fraction formats (`# ?/?`, `??/??`, `# ?/8`). Returns null when
 *  `pattern` is not a fraction format. */
function renderFractionPattern(value: number, pattern: string): string | null {
  const m = pattern.match(/^(.*?)([#0?]+)?\s*([#0?]+)\/([#0?]+|\d+)(.*?)$/);
  if (!m?.[3] || !m[4]) return null;
  // Reject when there's no `?`/`#`/`0` around the slash beyond a bare literal.
  const [, prefix, intPh, , denSpec, suffix] = m;
  const hasInt = !!intPh && intPh.length > 0;
  const sign = value < 0 ? '-' : '';
  let x = Math.abs(value);
  let whole = 0;
  if (hasInt) {
    whole = Math.floor(x);
    x -= whole;
  }
  let numer: number;
  let denom: number;
  if (/^\d+$/.test(denSpec)) {
    denom = Number.parseInt(denSpec, 10) || 1;
    numer = Math.round(x * denom);
  } else {
    const maxDen = 10 ** denSpec.length - 1;
    [numer, denom] = bestFraction(x, maxDen);
  }
  // Rounded up to a whole unit.
  if (numer === denom) {
    whole += 1;
    numer = 0;
  }
  const pre = normalizeFormatSection(prefix ?? '');
  const suf = normalizeFormatSection(suffix ?? '');
  if (numer === 0) {
    const wholeTxt = hasInt ? String(whole) : '0';
    return `${sign}${pre}${wholeTxt}${suf}`;
  }
  if (!hasInt) numer += whole * denom; // improper fraction
  const wholeTxt = hasInt && whole > 0 ? `${whole} ` : '';
  return `${sign}${pre}${wholeTxt}${numer}/${denom}${suf}`;
}

/** Render scientific / engineering notation (`0.00E+00`, `##0.0E+0`). The number
 *  of integer placeholders in the mantissa sets the exponent step, so `##0.0E+0`
 *  yields engineering notation (exponent stepped in multiples of 3). Returns null
 *  when `pattern` is not a scientific format. */
function renderScientificPattern(value: number, pattern: string, locale: string): string | null {
  const m = pattern.match(/^(.*?)([0#?][0#?,]*(?:\.[0#?]+)?)[eE]([+-]?)([0#?]+)(.*)$/);
  if (!m) return null;
  const prefix = normalizeFormatSection(m[1] ?? '');
  const mantissaPat = m[2] ?? '';
  const expSignSpec = m[3] ?? '';
  const expPat = m[4] ?? '';
  const suffix = normalizeFormatSection(m[5] ?? '');
  const dot = mantissaPat.indexOf('.');
  const intPat = dot >= 0 ? mantissaPat.slice(0, dot) : mantissaPat;
  const fracPat = dot >= 0 ? mantissaPat.slice(dot + 1) : '';
  const step = Math.max(1, (intPat.match(/[0#?]/g) ?? []).length);
  const minIntDigits = (intPat.match(/0/g) ?? []).length;
  const minFracDigits = (fracPat.match(/0/g) ?? []).length;
  const maxFracDigits = (fracPat.match(/[0#?]/g) ?? []).length;
  const sign = value < 0 ? '-' : '';
  const av = Math.abs(value);
  let exp = av === 0 ? 0 : Math.floor(Math.floor(Math.log10(av)) / step) * step;
  let mantissa = av === 0 ? 0 : av / 10 ** exp;
  // Rounding at the requested precision can push the mantissa past the step
  // boundary (e.g. 999.6 → 1000 for a 3-digit step); bump the exponent so the
  // integer part stays within `step` digits.
  if (av !== 0 && Number(mantissa.toFixed(maxFracDigits)) >= 10 ** step) {
    exp += step;
    mantissa = av / 10 ** exp;
  }
  const mantissaText = new Intl.NumberFormat(normalizeFormatLocale(locale), {
    minimumIntegerDigits: Math.max(1, minIntDigits),
    minimumFractionDigits: minFracDigits,
    maximumFractionDigits: maxFracDigits,
    useGrouping: false,
  }).format(mantissa);
  const expDigits = String(Math.abs(exp)).padStart(expPat.length, '0');
  const expSign = exp < 0 ? '-' : expSignSpec === '+' ? '+' : '';
  return `${sign}${prefix}${mantissaText}E${expSign}${expDigits}${suffix}`;
}

function renderNumericPattern(value: number, pattern: string, locale: string): string {
  // Scientific / engineering notation (`0.00E+00`, `##0.0E+0`) is solved into a
  // mantissa and exponent rather than printing the `E+0` placeholders verbatim.
  if (/[eE][+-]?[0#?]/.test(pattern)) {
    const sci = renderScientificPattern(value, pattern, locale);
    if (sci !== null) return sci;
  }
  // Fraction formats (`# ?/?`) are solved into integer + numerator/denominator
  // rather than printing the placeholders verbatim.
  if (pattern.includes('/')) {
    const frac = renderFractionPattern(value, pattern);
    if (frac !== null) return frac;
  }
  // Detect trailing thousand-scaling commas (e.g. "0,," divides by 1e6).
  let scale = 1;
  let body = pattern;
  // Remove trailing commas after the last digit placeholder for scaling.
  const trailingCommas = body.match(/[0#?](,+)\s*[^0#?]*$/);
  if (trailingCommas) {
    const commas = trailingCommas[1] ?? '';
    scale = 10 ** (3 * commas.length);
    // Remove just those commas from the body.
    const i = body.lastIndexOf(commas);
    if (i >= 0) body = body.slice(0, i) + body.slice(i + commas.length);
  }
  const isPercent = hasUnquotedPercent(body);
  let scaled = value / scale;
  if (isPercent) scaled *= 100;

  // Multi-run integer patterns (phone `000-000-0000`, SSN `000-00-0000`) spread
  // the digits across each placeholder run right-to-left. The single-block path
  // below only fills the first run, so delegate whole-number distribution to the
  // special renderer.
  if (!isPercent && !/\.[#0?]/.test(body)) {
    const runs = body.match(/[#0?][#0?,]*/g) ?? [];
    if (runs.length > 1 && !runs.some((r) => r.includes(','))) {
      return formatSpecialPattern(scaled, body);
    }
  }

  // Find the digit-placeholder block surrounding (and including) the decimal.
  const placeholderMatch = body.match(/[#0?][#0?,]*(?:\.[#0?]+)?|\.[#0?]+/);
  if (!placeholderMatch) return body;

  const block = placeholderMatch[0];
  const dotIndex = block.indexOf('.');
  const intPart = dotIndex >= 0 ? block.slice(0, dotIndex) : block;
  const fracPart = dotIndex >= 0 ? block.slice(dotIndex + 1) : '';
  const grouping = intPart.includes(',');
  const minIntDigits = (intPart.match(/0/g) ?? []).length;
  const minFracDigits = (fracPart.match(/0/g) ?? []).length;
  const maxFracDigits = (fracPart.match(/[0#?]/g) ?? []).length;

  const formatted = new Intl.NumberFormat(normalizeFormatLocale(locale), {
    minimumIntegerDigits: Math.max(1, minIntDigits),
    minimumFractionDigits: minFracDigits,
    maximumFractionDigits: maxFracDigits,
    useGrouping: grouping,
  }).format(scaled);

  return normalizeFormatSection(body.replace(block, formatted));
}
