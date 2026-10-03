import { addrKey } from '../engine/address.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { rangeArea, rangesIntersect } from '../store/selection-geometry.js';
import { type CellFormat, mutators, type SpreadsheetStore, type State } from '../store/store.js';
import { applyFlashFill, inferFlashFillPattern } from './flash-fill.js';
import type { History } from './history.js';
import { isCellWritable } from './protection.js';
import { shiftFormulaRefs } from './refs.js';
import { recordFormatChange } from './slice-history.js';

/**
 * Fill direction inferred from the relationship between source and dest.
 * Diagonal extensions (drag bottom-right corner past both axes) split into
 * two passes: row direction first, then column.
 */
type FillDir = 'down' | 'up' | 'right' | 'left' | 'copy';

interface SourceCell {
  row: number;
  col: number;
  formula: string | null;
  numeric: number | null;
  text: string | null;
  bool: boolean | null;
  errorCode: number | null;
  blank: boolean;
}

const MAX_FILL_RANGE_CELLS = 100_000;

const FULLWIDTH_ZERO = '０'.codePointAt(0) ?? 0xff10;

const toAsciiDigits = (s: string): string =>
  [...s]
    .map((ch) => {
      const cp = ch.codePointAt(0) ?? 0;
      if (cp >= 0xff10 && cp <= 0xff19) return String(cp - FULLWIDTH_ZERO);
      return ch;
    })
    .join('');

const toFullwidthDigits = (s: string): string =>
  [...s]
    .map((ch) => (ch >= '0' && ch <= '9' ? String.fromCodePoint(FULLWIDTH_ZERO + Number(ch)) : ch))
    .join('');

const formatTrailingNumber = (
  n: number,
  width: number,
  fullwidth: boolean,
  minus: '-' | '－',
): string => {
  const sign = n < 0 ? '-' : '';
  const abs = Math.abs(Math.trunc(n));
  const ascii = `${sign}${String(abs).padStart(width, '0')}`;
  const digits = fullwidth ? toFullwidthDigits(ascii) : ascii;
  return minus === '－' ? digits.replace('-', '－') : digits;
};

const numericFromText = (
  s: string,
): { prefix: string; n: number; width: number; fullwidth: boolean; minus: '-' | '－' } | null => {
  const m = s.match(/^(.*?)([-－]?[0-9０-９]+)$/);
  if (!m) return null;
  const prefix = m[1] ?? '';
  const rawDigits = m[2] ?? '';
  const minus = rawDigits.startsWith('－') ? '－' : '-';
  const fullwidth = /[０-９]/.test(rawDigits) && !/[0-9]/.test(rawDigits);
  const normalized = toAsciiDigits(rawDigits.replace('－', '-'));
  const n = Number(normalized);
  if (!Number.isFinite(n)) return null;
  return {
    prefix,
    n,
    width: normalized.startsWith('-') ? normalized.length - 1 : normalized.length,
    fullwidth,
    minus,
  };
};

/** Custom-list series — spreadsheets ship these by default. Each list represents
 *  a closed cycle that auto-fill steps through. Casing of the source is
 *  preserved by matching against the lowercased list. */
export const CUSTOM_LISTS: readonly string[][] = [
  ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],
  ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
  ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
  [
    'January',
    'February',
    'March',
    'April',
    'May',
    'June',
    'July',
    'August',
    'September',
    'October',
    'November',
    'December',
  ],
  ['日', '月', '火', '水', '木', '金', '土'],
  ['(日)', '(月)', '(火)', '(水)', '(木)', '(金)', '(土)'],
  ['日曜', '月曜', '火曜', '水曜', '木曜', '金曜', '土曜'],
  ['日曜日', '月曜日', '火曜日', '水曜日', '木曜日', '金曜日', '土曜日'],
  ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月'],
  [
    '１月',
    '２月',
    '３月',
    '４月',
    '５月',
    '６月',
    '７月',
    '８月',
    '９月',
    '１０月',
    '１１月',
    '１２月',
  ],
  ['Q1', 'Q2', 'Q3', 'Q4'],
  ['Q１', 'Q２', 'Q３', 'Q４'],
  ['QI', 'QII', 'QIII', 'QIV'],
  ['第1四半期', '第2四半期', '第3四半期', '第4四半期'],
  ['第１四半期', '第２四半期', '第３四半期', '第４四半期'],
];

/** Find the custom list (if any) that contains every source value. Returns
 *  the list and the list-indices of each source cell, or null when no single
 *  list matches all of them. */
function matchCustomList(values: string[]): { list: readonly string[]; indices: number[] } | null {
  for (const list of CUSTOM_LISTS) {
    const indices: number[] = [];
    let ok = true;
    for (const v of values) {
      const idx = list.findIndex((item) => item.toLowerCase() === v.toLowerCase());
      if (idx < 0) {
        ok = false;
        break;
      }
      indices.push(idx);
    }
    if (ok) return { list, indices };
  }
  return null;
}

const alphaChars = (s: string): string[] => [...s].filter((ch) => /[A-Za-z]/.test(ch));

const applyAlphaCase = (value: string, source: string): string => {
  const letters = alphaChars(source);
  if (letters.length === 0) return value;
  if (letters.every((ch) => ch === ch.toUpperCase())) return value.toUpperCase();
  if (letters.every((ch) => ch === ch.toLowerCase())) return value.toLowerCase();
  return value;
};

function readSource(state: State, src: Range): SourceCell[][] {
  const sheet = src.sheet;
  const out: SourceCell[][] = [];
  for (let r = src.r0; r <= src.r1; r += 1) {
    const row: SourceCell[] = [];
    for (let c = src.c0; c <= src.c1; c += 1) {
      const cell = state.data.cells.get(addrKey({ sheet, row: r, col: c }));
      if (!cell) {
        row.push({
          row: r,
          col: c,
          formula: null,
          numeric: null,
          text: null,
          bool: null,
          errorCode: null,
          blank: true,
        });
        continue;
      }
      const v = cell.value;
      row.push({
        row: r,
        col: c,
        formula: cell.formula,
        numeric: v.kind === 'number' ? v.value : null,
        text: v.kind === 'text' ? v.value : null,
        bool: v.kind === 'bool' ? v.value : null,
        errorCode: v.kind === 'error' ? v.code : null,
        blank: v.kind === 'blank' && !cell.formula,
      });
    }
    out.push(row);
  }
  return out;
}

function detectDirection(src: Range, dest: Range): FillDir {
  if (dest.r1 > src.r1 && dest.r0 === src.r0 && dest.c0 === src.c0 && dest.c1 === src.c1) {
    return 'down';
  }
  if (dest.r0 < src.r0 && dest.r1 === src.r1 && dest.c0 === src.c0 && dest.c1 === src.c1) {
    return 'up';
  }
  if (dest.c1 > src.c1 && dest.c0 === src.c0 && dest.r0 === src.r0 && dest.r1 === src.r1) {
    return 'right';
  }
  if (dest.c0 < src.c0 && dest.c1 === src.c1 && dest.r0 === src.r0 && dest.r1 === src.r1) {
    return 'left';
  }
  // Diagonal or arbitrary — treat as 2D copy/cycle.
  return 'copy';
}

interface SeriesProjection {
  /** Project the value at extension index `i` (1-based; 1 = first cell beyond
   *  source in the fill direction). Returns null when no value should be written. */
  at(i: number): ProjectedCell | null;
}

type ProjectedCell =
  | { kind: 'number'; value: number }
  | { kind: 'text'; value: string }
  | { kind: 'bool'; value: boolean }
  | { kind: 'error'; code: number }
  | { kind: 'formula'; formula: string; source: SourceCell }
  | { kind: 'blank' };

/**
 * Inspect a 1D source line and produce a series projector. Excel's default
 * AutoFill uses a least-squares linear trend for a numeric source, rather than
 * interpreting a stable ratio as a growth series. Formula slots in a mixed
 * source are tiled with shifted references; numeric slots advance by one per
 * repeated cycle and all other typed slots are copied as-is.
 */
type DateFillUnit = 'days' | 'weekdays' | 'months' | 'years';

const MS_PER_DAY = 86_400_000;
const SERIAL_UNIX_EPOCH = 25569;

const serialToDate = (serial: number): Date =>
  new Date((Math.floor(serial) - SERIAL_UNIX_EPOCH) * MS_PER_DAY);

const dateToSerial = (date: Date, fraction: number): number =>
  Math.floor(date.getTime() / MS_PER_DAY + SERIAL_UNIX_EPOCH) + fraction;

const daysInMonth = (year: number, month: number): number =>
  new Date(Date.UTC(year, month + 1, 0)).getUTCDate();

const addMonthsClamped = (date: Date, months: number): Date => {
  const year = date.getUTCFullYear();
  const month = date.getUTCMonth() + months;
  const day = date.getUTCDate();
  const first = new Date(Date.UTC(year, month, 1));
  const maxDay = daysInMonth(first.getUTCFullYear(), first.getUTCMonth());
  return new Date(Date.UTC(first.getUTCFullYear(), first.getUTCMonth(), Math.min(day, maxDay)));
};

const addWeekdays = (date: Date, weekdays: number): Date => {
  const next = new Date(date.getTime());
  let remaining = weekdays;
  const step = remaining < 0 ? -1 : 1;
  while (remaining !== 0) {
    next.setUTCDate(next.getUTCDate() + step);
    const dow = next.getUTCDay();
    if (dow !== 0 && dow !== 6) remaining -= step;
  }
  return next;
};

function buildDateProjection(line: SourceCell[], unit: DateFillUnit): SeriesProjection | null {
  if (line.length === 0 || !line.every((c) => c.numeric !== null && c.formula === null)) {
    return null;
  }
  const last = line[line.length - 1]?.numeric ?? 0;
  const fraction = last - Math.floor(last);
  const lastDate = serialToDate(last);
  return {
    at: (i) => {
      const date =
        unit === 'days'
          ? new Date(lastDate.getTime() + i * MS_PER_DAY)
          : unit === 'weekdays'
            ? addWeekdays(lastDate, i)
            : unit === 'months'
              ? addMonthsClamped(lastDate, i)
              : addMonthsClamped(lastDate, i * 12);
      return { kind: 'number', value: dateToSerial(date, fraction) };
    },
  };
}

function buildProjection(
  line: SourceCell[],
  dateUnit?: DateFillUnit,
  numericCycleSign: 1 | -1 = 1,
): SeriesProjection {
  if (line.length === 0) {
    return { at: () => null };
  }
  const dateProjection = dateUnit ? buildDateProjection(line, dateUnit) : null;
  if (dateProjection) return dateProjection;
  // All numeric?
  if (line.every((c) => c.numeric !== null && c.formula === null)) {
    if (line.length === 1) {
      const v = line[0]?.numeric ?? 0;
      return { at: () => ({ kind: 'number', value: v }) };
    }
    // Fit y = intercept + slope*x for x = 0, …, n-1. This gives Excel's
    // default linear-trend projection for irregular samples such as 1, 2, 4.
    const n = line.length;
    const meanX = (n - 1) / 2;
    const meanY = line.reduce((sum, cell) => sum + (cell.numeric ?? 0), 0) / n;
    let covariance = 0;
    let variance = 0;
    for (let x = 0; x < n; x += 1) {
      const dx = x - meanX;
      covariance += dx * ((line[x]?.numeric ?? 0) - meanY);
      variance += dx * dx;
    }
    const slope = variance === 0 ? 0 : covariance / variance;
    const intercept = meanY - slope * meanX;
    return {
      at: (i) => ({ kind: 'number', value: intercept + slope * (n - 1 + i) }),
    };
  }

  // Custom list match (e.g. Mon/Tue/Wed, Jan/Feb...) — preserves the casing
  // of the *list*, not the source.
  if (line.every((c) => c.text !== null && c.formula === null)) {
    const values = line.map((c) => c.text ?? '');
    const lm = matchCustomList(values);
    if (lm) {
      const lastIdx = lm.indices[lm.indices.length - 1] ?? 0;
      const step =
        lm.indices.length >= 2
          ? (lastIdx - (lm.indices[0] ?? 0)) / Math.max(1, lm.indices.length - 1)
          : 1;
      const stepInt = lm.indices.length >= 2 ? Math.round(step) : 1;
      return {
        at: (i) => {
          const target = lastIdx + stepInt * i;
          const len = lm.list.length;
          const idx = ((target % len) + len) % len;
          const value = lm.list[idx] ?? '';
          return { kind: 'text', value: applyAlphaCase(value, values[values.length - 1] ?? value) };
        },
      };
    }
  }

  // Trailing-integer text? Need every cell to have the same prefix and a numeric tail.
  if (line.every((c) => c.text !== null && c.formula === null)) {
    const parsed = line.map((c) => numericFromText(c.text ?? ''));
    if (parsed.every((p) => p !== null)) {
      const first = parsed[0];
      if (first && parsed.every((p) => p?.prefix === first.prefix)) {
        if (line.length === 1) {
          // Single cell: increment by 1 each step.
          return {
            at: (i) => ({
              kind: 'text',
              value: `${first.prefix}${formatTrailingNumber(
                first.n + i,
                first.width,
                first.fullwidth,
                first.minus,
              )}`,
            }),
          };
        }
        let stepSum = 0;
        for (let i = 1; i < parsed.length; i += 1) {
          stepSum += (parsed[i]?.n ?? 0) - (parsed[i - 1]?.n ?? 0);
        }
        const step = stepSum / (parsed.length - 1);
        const last = parsed[parsed.length - 1]?.n ?? 0;
        const lastParsed = parsed[parsed.length - 1];
        return {
          at: (i) => ({
            kind: 'text',
            value: `${first.prefix}${formatTrailingNumber(
              Math.round(last + step * i),
              lastParsed?.width ?? first.width,
              lastParsed?.fullwidth ?? first.fullwidth,
              lastParsed?.minus ?? first.minus,
            )}`,
          }),
        };
      }
    }
  }

  // A mixed formula/value source repeats its slots. Formula refs shift from
  // their original source address, while numeric slots advance one per cycle.
  // This branch also handles formula-only lines, preserving formula errors as
  // formulas rather than converting their current error result to a literal.
  if (line.some((c) => c.formula !== null)) {
    return {
      at: (i) => {
        const cycle = Math.floor((i - 1) / line.length) + 1;
        const idx = (((i - 1) % line.length) + line.length) % line.length;
        const c = line[idx];
        if (!c) return null;
        if (c.formula !== null) return { kind: 'formula', formula: c.formula, source: c };
        if (c.numeric !== null)
          return { kind: 'number', value: c.numeric + numericCycleSign * cycle };
        if (c.text !== null) return { kind: 'text', value: c.text };
        if (c.bool !== null) return { kind: 'bool', value: c.bool };
        if (c.errorCode !== null) return { kind: 'error', code: c.errorCode };
        return { kind: 'blank' };
      },
    };
  }

  // Cycle/copy for non-formula typed values.
  return {
    at: (i) => {
      const idx = (((i - 1) % line.length) + line.length) % line.length;
      const c = line[idx];
      if (!c) return null;
      if (c.numeric !== null) return { kind: 'number', value: c.numeric };
      if (c.text !== null) return { kind: 'text', value: c.text };
      if (c.bool !== null) return { kind: 'bool', value: c.bool };
      if (c.errorCode !== null) return { kind: 'error', code: c.errorCode };
      return { kind: 'blank' };
    },
  };
}

const writeProjected = (
  wb: WorkbookHandle,
  sheet: number,
  row: number,
  col: number,
  projected: ReturnType<SeriesProjection['at']>,
): boolean => {
  if (!projected) return false;
  if (projected.kind === 'number') wb.setNumber({ sheet, row, col }, projected.value);
  else if (projected.kind === 'text') wb.setText({ sheet, row, col }, projected.value);
  else if (projected.kind === 'bool') wb.setBool({ sheet, row, col }, projected.value);
  else if (projected.kind === 'error') wb.setError({ sheet, row, col }, projected.code);
  else if (projected.kind === 'formula') {
    wb.setFormula(
      { sheet, row, col },
      shiftFormulaRefs(projected.formula, row - projected.source.row, col - projected.source.col),
    );
  } else wb.setBlank({ sheet, row, col });
  return true;
};

const projectedCell = (cell: SourceCell | undefined): ProjectedCell => {
  if (!cell || cell.blank) return { kind: 'blank' };
  if (cell.formula !== null) return { kind: 'formula', formula: cell.formula, source: cell };
  if (cell.numeric !== null) return { kind: 'number', value: cell.numeric };
  if (cell.text !== null) return { kind: 'text', value: cell.text };
  if (cell.bool !== null) return { kind: 'bool', value: cell.bool };
  if (cell.errorCode !== null) return { kind: 'error', code: cell.errorCode };
  return { kind: 'blank' };
};

const cloneFormat = (format: CellFormat): CellFormat => {
  const clone: CellFormat = {
    ...format,
    borders: format.borders ? { ...format.borders } : undefined,
  };
  delete clone.comment;
  delete clone.commentAuthor;
  return clone;
};

const preserveDestinationComments = (source: CellFormat, destination?: CellFormat): CellFormat => {
  const next = cloneFormat(source);
  if (destination?.comment !== undefined) next.comment = destination.comment;
  if (destination?.commentAuthor !== undefined) next.commentAuthor = destination.commentAuthor;
  return next;
};

const sourceCoordFor = (value: number, start: number, size: number): number =>
  start + ((((value - start) % size) + size) % size);

function applyFillFormats(state: State, store: SpreadsheetStore, src: Range, dest: Range): boolean {
  const srcRows = src.r1 - src.r0 + 1;
  const srcCols = src.c1 - src.c0 + 1;
  let changed = false;
  store.setState((s) => {
    const formats = new Map(s.format.formats);
    for (let r = dest.r0; r <= dest.r1; r += 1) {
      for (let c = dest.c0; c <= dest.c1; c += 1) {
        if (r >= src.r0 && r <= src.r1 && c >= src.c0 && c <= src.c1) continue;
        const source = {
          sheet: src.sheet,
          row: sourceCoordFor(r, src.r0, srcRows),
          col: sourceCoordFor(c, src.c0, srcCols),
        };
        const targetKey = addrKey({ sheet: dest.sheet, row: r, col: c });
        const sourceFormat = state.format.formats.get(addrKey(source));
        const destinationFormat = formats.get(targetKey);
        if (sourceFormat) {
          formats.set(targetKey, preserveDestinationComments(sourceFormat, destinationFormat));
        } else {
          const preservedComments = destinationFormat
            ? preserveDestinationComments({}, destinationFormat)
            : null;
          if (
            preservedComments &&
            (preservedComments.comment !== undefined ||
              preservedComments.commentAuthor !== undefined)
          )
            formats.set(targetKey, preservedComments);
          else formats.delete(targetKey);
        }
        changed = true;
      }
    }
    return changed ? { ...s, format: { ...s.format, formats } } : s;
  });
  return changed;
}

export type FillFormattingMode = 'with' | 'without' | 'only';

export interface FillOptions {
  /** Bypass series detection and tile the source verbatim (Ctrl-drag). */
  copyOnly?: boolean;
  /** Spreadsheet Auto Fill option for including, suppressing, or only copying format. */
  formatting?: FillFormattingMode;
  /** Explicit date-series Auto Fill option. */
  dateUnit?: DateFillUnit;
  /** Required when `formatting` should write cell formats. */
  store?: SpreadsheetStore;
}

/**
 * Fill the cells in `dest` (which contains `src` as a sub-rect) by projecting
 * the source range outward. Returns true if the fill produced any writes.
 *
 * Formula references are shifted from each source cell to its destination. For
 * pure values, the series detector follows Excel's linear / increment / copy
 * behavior.
 */
export function fillRange(
  state: State,
  wb: WorkbookHandle,
  src: Range,
  dest: Range,
  opts?: FillOptions,
): boolean {
  return wb.withBatchedRecalc(() => applyFillRange(state, wb, src, dest, opts));
}

/** Fill is an all-or-nothing cell operation when a protected destination or a
 * merged destination is involved. Do this check before the first engine write
 * so an invalid selection cannot leave a partially filled range behind. The
 * pointer interaction explicitly unmerges its drag destination first; direct
 * command/ribbon callers preserve the merge and refuse the operation. */
const fillDestinationWritable = (state: State, src: Range, dest: Range): boolean => {
  if (src.sheet !== dest.sheet) return false;
  for (let row = dest.r0; row <= dest.r1; row += 1) {
    for (let col = dest.c0; col <= dest.c1; col += 1) {
      if (row >= src.r0 && row <= src.r1 && col >= src.c0 && col <= src.c1) continue;
      const addr = { sheet: dest.sheet, row, col };
      if (!isCellWritable(state, addr)) return false;
      const mergeAnchor = state.merges.byCell.get(addrKey(addr)) ?? addrKey(addr);
      const merge = mergeAnchor ? state.merges.byAnchor.get(mergeAnchor) : undefined;
      if (merge && rangesIntersect(merge, dest)) return false;
    }
  }
  return true;
};

function applyFillRange(
  state: State,
  wb: WorkbookHandle,
  src: Range,
  dest: Range,
  opts?: FillOptions,
): boolean {
  if (dest.r0 === src.r0 && dest.r1 === src.r1 && dest.c0 === src.c0 && dest.c1 === src.c1) {
    return false;
  }
  if (rangeArea(src) > MAX_FILL_RANGE_CELLS || rangeArea(dest) > MAX_FILL_RANGE_CELLS) {
    return false;
  }
  if (!fillDestinationWritable(state, src, dest)) return false;
  const dir = opts?.copyOnly ? 'copy' : detectDirection(src, dest);
  const sheet = src.sheet;
  const source = readSource(state, src);
  const writeValues = opts?.formatting !== 'only';
  const writeFormats = opts?.store && opts.formatting !== 'without';
  let wroteValues = false;

  if (dir === 'down' || dir === 'up') {
    // Fill column-by-column. Each column gets its own projection.
    const cols = src.c1 - src.c0 + 1;
    for (let c = 0; c < cols; c += 1) {
      const line: SourceCell[] = source.map((row) => row[c] as SourceCell);
      const proj = buildProjection(line, opts?.dateUnit);
      if (dir === 'down') {
        const ext = dest.r1 - src.r1;
        for (let i = 1; i <= ext; i += 1) {
          const destRow = src.r1 + i;
          const destCol = src.c0 + c;
          if (writeValues && writeProjected(wb, sheet, destRow, destCol, proj.at(i)))
            wroteValues = true;
        }
      } else {
        // Up: extension index 1 is the row just above src.r0.
        const ext = src.r0 - dest.r0;
        const reversed = [...line].reverse();
        const projUp = buildProjection(reversed, opts?.dateUnit, -1);
        for (let i = 1; i <= ext; i += 1) {
          const destRow = src.r0 - i;
          const destCol = src.c0 + c;
          if (writeValues && writeProjected(wb, sheet, destRow, destCol, projUp.at(i)))
            wroteValues = true;
        }
      }
    }
    const wroteFormats = writeFormats
      ? applyFillFormats(state, opts.store as SpreadsheetStore, src, dest)
      : false;
    return wroteValues || wroteFormats;
  }

  if (dir === 'right' || dir === 'left') {
    const rows = src.r1 - src.r0 + 1;
    for (let r = 0; r < rows; r += 1) {
      const line: SourceCell[] = source[r] ?? [];
      const proj = buildProjection(line, opts?.dateUnit);
      if (dir === 'right') {
        const ext = dest.c1 - src.c1;
        for (let i = 1; i <= ext; i += 1) {
          const destRow = src.r0 + r;
          const destCol = src.c1 + i;
          if (writeValues && writeProjected(wb, sheet, destRow, destCol, proj.at(i)))
            wroteValues = true;
        }
      } else {
        const ext = src.c0 - dest.c0;
        const reversed = [...line].reverse();
        const projLeft = buildProjection(reversed, opts?.dateUnit, -1);
        for (let i = 1; i <= ext; i += 1) {
          const destRow = src.r0 + r;
          const destCol = src.c0 - i;
          if (writeValues && writeProjected(wb, sheet, destRow, destCol, projLeft.at(i)))
            wroteValues = true;
        }
      }
    }
    const wroteFormats = writeFormats
      ? applyFillFormats(state, opts.store as SpreadsheetStore, src, dest)
      : false;
    return wroteValues || wroteFormats;
  }

  // 2D / arbitrary: tile source over dest.
  const sR = src.r1 - src.r0 + 1;
  const sC = src.c1 - src.c0 + 1;
  for (let r = dest.r0; r <= dest.r1; r += 1) {
    for (let c = dest.c0; c <= dest.c1; c += 1) {
      if (r >= src.r0 && r <= src.r1 && c >= src.c0 && c <= src.c1) continue; // skip source
      const sr = (((r - src.r0) % sR) + sR) % sR;
      const sc = (((c - src.c0) % sC) + sC) % sC;
      const cell = source[sr]?.[sc];
      if (!writeValues) continue;
      if (writeProjected(wb, sheet, r, c, projectedCell(cell))) wroteValues = true;
    }
  }
  const wroteFormats = writeFormats
    ? applyFillFormats(state, opts.store as SpreadsheetStore, src, dest)
    : false;
  return wroteValues || wroteFormats;
}

/** Compute the dest range for a drag from the source's bottom-right corner
 *  to a target cell (the cursor position). Spreadsheets lock the extension to
 *  whichever axis is most extended — diagonal drags don't extend both axes
 *  (unless the target is fully inside both). */
export function fillDestFor(src: Range, target: { row: number; col: number }): Range {
  // Decide whether to extend rows or cols based on which delta is larger.
  const dRow =
    target.row > src.r1 ? target.row - src.r1 : target.row < src.r0 ? src.r0 - target.row : 0;
  const dCol =
    target.col > src.c1 ? target.col - src.c1 : target.col < src.c0 ? src.c0 - target.col : 0;
  if (dRow === 0 && dCol === 0) return src;
  if (dRow >= dCol) {
    return target.row > src.r1
      ? { sheet: src.sheet, r0: src.r0, c0: src.c0, r1: target.row, c1: src.c1 }
      : { sheet: src.sheet, r0: target.row, c0: src.c0, r1: src.r1, c1: src.c1 };
  }
  return target.col > src.c1
    ? { sheet: src.sheet, r0: src.r0, c0: src.c0, r1: src.r1, c1: target.col }
    : { sheet: src.sheet, r0: src.r0, c0: target.col, r1: src.r1, c1: src.c1 };
}

export type RibbonFillAction =
  | 'down'
  | 'right'
  | 'up'
  | 'left'
  | 'flash'
  | 'series'
  | 'days'
  | 'weekdays'
  | 'months'
  | 'years';

export interface ExecuteRibbonFillActionDeps {
  store: SpreadsheetStore;
  workbook: WorkbookHandle;
  history: History;
  action: RibbonFillAction;
}

const cellValueAsText = (value: ReturnType<WorkbookHandle['getValue']>): string => {
  if (value.kind === 'text') return value.value;
  if (value.kind === 'number') return String(value.value);
  if (value.kind === 'bool') return String(value.value);
  return '';
};

const MAX_FLASH_FILL_ROWS = 100_000;

const flashFillInputCandidates = (
  workbook: WorkbookHandle,
  sheet: number,
  range: Range,
): number[] => {
  const cols = new Set<number>();
  for (const cell of workbook.cells(sheet)) {
    if (cell.addr.row < range.r0 || cell.addr.row > range.r1) continue;
    if (cell.addr.col === range.c0) continue;
    if (cellValueAsText(cell.value).length === 0) continue;
    cols.add(cell.addr.col);
  }
  return [...cols].sort((a, b) => {
    const da = Math.abs(a - range.c0);
    const db = Math.abs(b - range.c0);
    if (da !== db) return da - db;
    const aLeft = a < range.c0;
    const bLeft = b < range.c0;
    if (aLeft !== bLeft) return aLeft ? -1 : 1;
    return a - b;
  });
};

const executeRibbonFlashFill = (
  store: SpreadsheetStore,
  workbook: WorkbookHandle,
  history: History,
  range: Range,
): boolean => {
  if (range.c0 !== range.c1) return false;
  if (range.r1 - range.r0 + 1 > MAX_FLASH_FILL_ROWS) return false;
  const candidates = flashFillInputCandidates(workbook, range.sheet, range);
  for (const inputCol of candidates) {
    const examples: { input: string; output: string }[] = [];
    const pending: { row: number; input: string }[] = [];
    for (let row = range.r0; row <= range.r1; row += 1) {
      const inputValue = workbook.getValue({ sheet: range.sheet, row, col: inputCol });
      const outputValue = workbook.getValue({ sheet: range.sheet, row, col: range.c0 });
      const input = cellValueAsText(inputValue);
      if (input.length === 0) continue;
      if (outputValue.kind === 'text' && outputValue.value.length > 0) {
        examples.push({ input, output: outputValue.value });
      } else if (
        outputValue.kind === 'blank' &&
        isCellWritable(store.getState(), { sheet: range.sheet, row, col: range.c0 })
      ) {
        pending.push({ row, input });
      }
    }
    const pattern = inferFlashFillPattern(examples);
    if (!pattern || pending.length === 0) continue;
    const filled = applyFlashFill(
      pattern,
      pending.map((entry) => entry.input),
    );
    history.begin();
    try {
      workbook.withBatchedRecalc(() => {
        pending.forEach((entry, index) => {
          const value = filled[index];
          if (value != null)
            workbook.setText({ sheet: range.sheet, row: entry.row, col: range.c0 }, value);
        });
      });
    } finally {
      history.end();
    }
    mutators.replaceCells(store, workbook.cells(range.sheet));
    return true;
  }
  return false;
};

const DATE_SERIES_ACTIONS = new Set(['days', 'weekdays', 'months', 'years']);

/** Shared "Fill" ribbon split-button action. Covers flash-fill, directional
 *  copy (down/right/up/left), series, and date-series variants. Returns
 *  whether anything actually changed so hosts can short-circuit dropdown
 *  closure or status updates. */
export const executeRibbonFillAction = (deps: ExecuteRibbonFillActionDeps): boolean => {
  const { store, workbook, history, action } = deps;
  const selected = store.getState().selection.range;
  const range: Range = {
    sheet: selected.sheet,
    r0: Math.min(selected.r0, selected.r1),
    c0: Math.min(selected.c0, selected.c1),
    r1: Math.max(selected.r0, selected.r1),
    c1: Math.max(selected.c0, selected.c1),
  };
  if (action === 'flash') return executeRibbonFlashFill(store, workbook, history, range);

  const direction: 'down' | 'right' | 'up' | 'left' =
    action === 'down' || action === 'right' || action === 'up' || action === 'left'
      ? action
      : 'down';
  let src = range;
  if (direction === 'down') src = { ...range, r1: range.r0 };
  else if (direction === 'up') src = { ...range, r0: range.r1 };
  else if (direction === 'right') src = { ...range, c1: range.c0 };
  else src = { ...range, c0: range.c1 };
  if (src.r0 === range.r0 && src.r1 === range.r1 && src.c0 === range.c0 && src.c1 === range.c1)
    return false;

  const isDateSeries = DATE_SERIES_ACTIONS.has(action);
  let changed = false;
  history.begin();
  try {
    recordFormatChange(history, store, () => {
      changed = fillRange(store.getState(), workbook, src, range, {
        copyOnly: action === 'down' || action === 'right' || action === 'up' || action === 'left',
        dateUnit: isDateSeries ? (action as 'days' | 'weekdays' | 'months' | 'years') : undefined,
        formatting: 'with',
        store,
      });
    });
  } finally {
    history.end();
  }
  if (changed) mutators.replaceCells(store, workbook.cells(range.sheet));
  return changed;
};
