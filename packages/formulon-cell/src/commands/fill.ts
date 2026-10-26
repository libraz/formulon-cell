import { addrKey } from '../engine/address.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { excelSerialFromDate, excelSerialToDate } from '../format/date-serial.js';
import { rangeArea, rangesIntersect } from '../store/selection-geometry.js';
import { type CellFormat, mutators, type SpreadsheetStore, type State } from '../store/store.js';
import {
  sameTextSeries,
  type TextSeriesAtom,
  textSeriesAtom,
  textSeriesValue,
} from './autofill-series.js';
import { inputLocaleForWorkbook } from './coerce-input.js';
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
  /** Number format family: dates step by days or months, times by hours. */
  dateKind: 'date' | 'time' | null;
}

const MAX_FILL_RANGE_CELLS = 100_000;

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
          dateKind: null,
        });
        continue;
      }
      const v = cell.value;
      const fmtKind = state.format.formats.get(addrKey({ sheet, row: r, col: c }))?.numFmt?.kind;
      row.push({
        row: r,
        col: c,
        formula: cell.formula,
        numeric: v.kind === 'number' ? v.value : null,
        text: v.kind === 'text' ? v.value : null,
        bool: v.kind === 'bool' ? v.value : null,
        errorCode: v.kind === 'error' ? v.code : null,
        blank: v.kind === 'blank' && !cell.formula,
        dateKind:
          fmtKind === 'date' || fmtKind === 'datetime'
            ? 'date'
            : fmtKind === 'time'
              ? 'time'
              : null,
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

type DateFillUnit = 'days' | 'weekdays' | 'months' | 'years';

interface ProjectionContext {
  locale: string | undefined;
  dateUnit?: DateFillUnit;
  /** Fill Series: a lone number steps by one instead of repeating. */
  series?: boolean;
  /** -1 when filling up or left, where the line arrives reversed. */
  sign: 1 | -1;
}

const serialToDate = (serial: number): Date => excelSerialToDate(serial) ?? new Date(Number.NaN);

const dateToSerial = (date: Date, fraction: number): number =>
  Math.floor(excelSerialFromDate(date) ?? Number.NaN) + fraction;

const daysInMonth = (year: number, month: number): number =>
  new Date(Date.UTC(year, month + 1, 0)).getUTCDate();

const addMonthsClamped = (date: Date, months: number, endOfMonth = false): Date => {
  const first = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth() + months, 1));
  const maxDay = daysInMonth(first.getUTCFullYear(), first.getUTCMonth());
  const day = endOfMonth ? maxDay : Math.min(date.getUTCDate(), maxDay);
  return new Date(Date.UTC(first.getUTCFullYear(), first.getUTCMonth(), day));
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

/** Auto Fill Options' explicit date units step from the last source cell. */
function buildDateUnitProjection(
  line: SourceCell[],
  unit: DateFillUnit,
  sign: 1 | -1,
): SeriesProjection | null {
  if (line.length === 0 || !line.every((c) => c.numeric !== null && c.formula === null)) {
    return null;
  }
  const last = line[line.length - 1]?.numeric ?? 0;
  const fraction = last - Math.floor(last);
  const lastDate = serialToDate(last);
  return {
    at: (i) => {
      const step = i * sign;
      if (unit === 'days') return { kind: 'number', value: last + step };
      const date =
        unit === 'weekdays'
          ? addWeekdays(lastDate, step)
          : addMonthsClamped(lastDate, unit === 'months' ? step : step * 12);
      return { kind: 'number', value: dateToSerial(date, fraction) };
    },
  };
}

/** Least-squares line through `ys` at x = 0..n-1, as AutoFill extends a trend. */
const linearFit = (ys: readonly number[]): ((x: number) => number) => {
  const n = ys.length;
  const meanX = (n - 1) / 2;
  const meanY = ys.reduce((sum, y) => sum + y, 0) / n;
  let covariance = 0;
  let variance = 0;
  ys.forEach((y, x) => {
    covariance += (x - meanX) * (y - meanY);
    variance += (x - meanX) ** 2;
  });
  const slope = variance === 0 ? 0 : covariance / variance;
  return (x) => meanY + slope * (x - meanX);
};

/** A run of consecutive source cells that extend as one series. `at(x)` is the
 *  value at series position x, where the source occupies 0..length-1. */
interface Run {
  length: number;
  at(x: number): ProjectedCell | null;
}

const copyRun = (cells: SourceCell[]): Run => ({
  length: cells.length,
  at: (x) => projectedCell(cells[x % cells.length]),
});

/** Dates with one day of the month (or all month ends) step by months; dates a
 *  whole number of days apart step by days; any other dates repeat. */
function dateRun(cells: SourceCell[], ctx: ProjectionContext): Run {
  const values = cells.map((c) => c.numeric ?? 0);
  const k = values.length;
  const first = values[0] ?? 0;
  if (k === 1) {
    const unit = cells[0]?.dateKind === 'time' ? 1 / 24 : 1;
    return { length: 1, at: (x) => ({ kind: 'number', value: first + ctx.sign * x * unit }) };
  }
  if (cells[0]?.dateKind === 'time') {
    const trend = linearFit(values);
    return { length: k, at: (x) => ({ kind: 'number', value: trend(x) }) };
  }
  const dates = values.map(serialToDate);
  const monthIndex = (d: Date): number => d.getUTCFullYear() * 12 + d.getUTCMonth();
  const isMonthEnd = (d: Date): boolean =>
    d.getUTCDate() === daysInMonth(d.getUTCFullYear(), d.getUTCMonth());
  const months = dates.slice(1).map((d, i) => monthIndex(d) - monthIndex(dates[i] as Date));
  const monthStep = months[0] ?? 0;
  const endOfMonth = dates.every(isMonthEnd);
  if (
    monthStep !== 0 &&
    months.every((m) => m === monthStep) &&
    (endOfMonth || dates.every((d) => d.getUTCDate() === dates[0]?.getUTCDate()))
  ) {
    const anchor = dates[0] as Date;
    const fraction = first - Math.floor(first);
    return {
      length: k,
      at: (x) => ({
        kind: 'number',
        value: dateToSerial(addMonthsClamped(anchor, monthStep * x, endOfMonth), fraction),
      }),
    };
  }
  const days = values.slice(1).map((v, i) => Math.trunc(v - (values[i] ?? 0)));
  const dayStep = days[0] ?? 0;
  if (days.every((d) => d === dayStep)) {
    const last = values[k - 1] ?? 0;
    return { length: k, at: (x) => ({ kind: 'number', value: last + dayStep * (x - k + 1) }) };
  }
  return copyRun(cells);
}

function numberRun(cells: SourceCell[], ctx: ProjectionContext, lone: boolean): Run {
  const values = cells.map((c) => c.numeric ?? 0);
  if (values.length === 1) {
    const v = values[0] ?? 0;
    // A lone number repeats unless Fill Series asks for steps; inside a mixed
    // source it steps by one per cycle.
    const step = lone && !ctx.series ? 0 : ctx.sign;
    return { length: 1, at: (x) => ({ kind: 'number', value: v + step * x }) };
  }
  const trend = linearFit(values);
  return { length: values.length, at: (x) => ({ kind: 'number', value: trend(x) }) };
}

function textRun(cells: SourceCell[], atoms: TextSeriesAtom[], ctx: ProjectionContext): Run {
  const shape = atoms[atoms.length - 1] as TextSeriesAtom;
  const casing = cells[0]?.text ?? '';
  // List positions are unwrapped so Sat, Sun steps forward by one.
  const positions: number[] = [];
  atoms.forEach((atom, i) => {
    const raw = atom.kind === 'list' ? atom.index : atom.n;
    const prev = positions[i - 1];
    if (prev === undefined || atom.kind !== 'list') {
      positions.push(raw);
      return;
    }
    const len = atom.list.length;
    let diff = (((raw - prev) % len) + len) % len;
    if (diff > len / 2) diff -= len;
    positions.push(prev + diff);
  });
  const start = positions[0] ?? 0;
  const trend = positions.length === 1 ? (x: number) => start + ctx.sign * x : linearFit(positions);
  return {
    length: atoms.length,
    at: (x) => ({
      kind: 'text',
      value: textSeriesValue(shape, Math.round(trend(x)), ctx.locale, casing),
    }),
  };
}

type SlotKind =
  | { type: 'number' }
  | { type: 'date' }
  | { type: 'text'; atom: TextSeriesAtom }
  | { type: 'single' };

const slotKind = (cell: SourceCell, locale: string | undefined): SlotKind => {
  if (cell.formula !== null) return { type: 'single' };
  if (cell.numeric !== null) return cell.dateKind ? { type: 'date' } : { type: 'number' };
  if (cell.text !== null) {
    const atom = textSeriesAtom(cell.text, locale);
    if (atom) return { type: 'text', atom };
  }
  return { type: 'single' };
};

const sameRun = (a: SlotKind, b: SlotKind, cellA: SourceCell, cellB: SourceCell): boolean => {
  if (a.type === 'text' && b.type === 'text') return sameTextSeries(a.atom, b.atom);
  if (a.type === 'date' && b.type === 'date') return cellA.dateKind === cellB.dateKind;
  return a.type === 'number' && b.type === 'number';
};

/**
 * Inspect a 1D source line and produce a series projector. The line splits
 * into runs of consecutive cells of one kind (numbers, dates, one list or one
 * numbered text); each run extends as its own series each time the source
 * pattern repeats, and every other cell, formulas included, repeats as-is.
 */
function buildProjection(line: SourceCell[], ctx: ProjectionContext): SeriesProjection {
  if (line.length === 0) return { at: () => null };
  if (ctx.dateUnit) {
    const projection = buildDateUnitProjection(line, ctx.dateUnit, ctx.sign);
    if (projection) return projection;
  }
  const kinds = line.map((cell) => slotKind(cell, ctx.locale));
  const slots: { run: Run; offset: number }[] = [];
  let start = 0;
  while (start < line.length) {
    const kind = kinds[start] as SlotKind;
    let end = start + 1;
    while (
      end < line.length &&
      kind.type !== 'single' &&
      sameRun(kind, kinds[end] as SlotKind, line[start] as SourceCell, line[end] as SourceCell)
    ) {
      end += 1;
    }
    const cells = line.slice(start, end);
    const run =
      kind.type === 'number'
        ? numberRun(cells, ctx, line.length === 1)
        : kind.type === 'date'
          ? dateRun(cells, ctx)
          : kind.type === 'text'
            ? textRun(
                cells,
                kinds.slice(start, end).map((k) => (k as { atom: TextSeriesAtom }).atom),
                ctx,
              )
            : copyRun(cells);
    for (let offset = 0; offset < cells.length; offset += 1) slots.push({ run, offset });
    start = end;
  }
  return {
    at: (i) => {
      const cycle = Math.floor((i - 1) / line.length) + 1;
      const slot = slots[(i - 1) % line.length];
      return slot ? slot.run.at(slot.run.length * cycle + slot.offset) : null;
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
  /** Fill Series Auto Fill option: a lone number steps by one instead of repeating. */
  series?: boolean;
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
  const series = {
    locale: inputLocaleForWorkbook(wb),
    dateUnit: opts?.dateUnit,
    series: opts?.series,
  };
  const writeValues = opts?.formatting !== 'only';
  const writeFormats = opts?.store && opts.formatting !== 'without';
  let wroteValues = false;

  if (dir === 'down' || dir === 'up') {
    // Fill column-by-column. Each column gets its own projection.
    const cols = src.c1 - src.c0 + 1;
    for (let c = 0; c < cols; c += 1) {
      const line: SourceCell[] = source.map((row) => row[c] as SourceCell);
      const proj = buildProjection(line, { ...series, sign: 1 });
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
        const projUp = buildProjection(reversed, { ...series, sign: -1 });
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
      const proj = buildProjection(line, { ...series, sign: 1 });
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
        const projLeft = buildProjection(reversed, { ...series, sign: -1 });
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

/** The cells a fill of `src` into `dest` writes: `dest` without the source rows or columns. */
export function fillOutputRange(src: Range, dest: Range): Range {
  if (dest.r1 > src.r1) return { ...dest, r0: src.r1 + 1 };
  if (dest.r0 < src.r0) return { ...dest, r1: src.r0 - 1 };
  if (dest.c1 > src.c1) return { ...dest, c0: src.c1 + 1 };
  if (dest.c0 < src.c0) return { ...dest, c1: src.c0 - 1 };
  return dest;
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
