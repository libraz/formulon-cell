import { type CellRect, CellShift, type Sheet as EngineSheet } from '@libraz/formulon';
import { collectAllFormulas } from '../../commands/formula-records.js';
import type { HistoryEntry } from '../../commands/history.js';
import { colFromLetters, MAX_COL, MAX_ROW, rangeOfRect } from '../../engine/address.js';
import { writeCell } from '../../engine/value.js';
import type {
  FacadeChangeKind,
  ChangeEvent as HandleChangeEvent,
  WorkbookHandle,
} from '../../engine/workbook-handle.js';
import type {
  JournalChoice,
  JournalCodec,
  JournalFamily,
  JournalScope,
  JournalSpec,
  Restore,
} from './index.js';
import { rowsColsCodec } from './layout.js';
import { DRAWINGS, PIVOTS, TABLES } from './objects.js';
import {
  AUTOFILTER,
  CF,
  COMMENTS,
  NAMES,
  PRINT,
  SHEET_CHROME,
  SHEET_PROT,
  SPARKLINES,
  THREADS,
  VALIDATION,
} from './records.js';

type Method = (...args: unknown[]) => unknown;

const invoke = (sheet: EngineSheet, name: string, args: readonly unknown[]): void => {
  (sheet as unknown as Record<string, Method>)[name]?.apply(sheet, args as unknown[]);
};

/** Facade writes that shift rows, columns or cells; each runs as a segment of its own. */
const STRUCTURAL_MEMBERS: readonly string[] = [
  'insertRows',
  'deleteRows',
  'insertColumns',
  'deleteColumns',
  'insertCells',
  'deleteCells',
];

/** Whether a write shifts cell addresses, so address-keyed snapshots must not span it. */
export const isIndexShifting = (spec: JournalSpec | undefined): boolean =>
  typeof spec === 'object' && spec.hostId === 'sheet' && STRUCTURAL_MEMBERS.includes(spec.name);

/** The inverse of an insert: the matching delete. */
const INSERT_INVERSE: Readonly<Record<string, string>> = {
  insertRows: 'deleteRows',
  insertColumns: 'deleteColumns',
  insertCells: 'deleteCells',
};

/** The inverse of a delete: the matching insert. */
const DELETE_INVERSE: Readonly<Record<string, string>> = {
  deleteRows: 'insertRows',
  deleteColumns: 'insertColumns',
  deleteCells: 'insertCells',
};

const INVERSE_SHIFT: Readonly<Record<number, number>> = {
  [CellShift.Down]: CellShift.Up,
  [CellShift.Right]: CellShift.Left,
  [CellShift.Up]: CellShift.Down,
  [CellShift.Left]: CellShift.Right,
};

/** The arguments of the inverse write: a cell shift flips, spans and ranges carry over. */
const inverseArgs = (name: string, args: readonly unknown[]): unknown[] =>
  name.endsWith('Cells') ? [args[0], INVERSE_SHIFT[args[1] as number]] : [...args];

const ROW_SPAN = /^\s*\$?(\d+)(?::\$?(\d+))?\s*$/;
const COLUMN_SPAN = /^\s*\$?([A-Za-z]+)(?::\$?([A-Za-z]+))?\s*$/;

/** The first and last index of the rows or columns a write names, as a pair or an A1 span. */
const spanOf = (args: readonly unknown[], columns: boolean): [number, number] | null => {
  if (typeof args[0] === 'number') return [args[0], (args[1] ?? args[0]) as number];
  const match = String(args[0]).match(columns ? COLUMN_SPAN : ROW_SPAN);
  if (!match) return null;
  const at = (text: string): number => (columns ? colFromLetters(text) : Number(text) - 1);
  const first = at(match[1] as string);
  return [first, match[2] ? at(match[2]) : first];
};

/** The cells a delete removes and the region it shifts, as rectangles. */
const deleteRects = (scope: JournalScope): { deleted: CellRect; shifted: CellRect } | null => {
  const { name, args } = scope;
  if (name === 'deleteRows' || name === 'deleteColumns') {
    const columns = name === 'deleteColumns';
    const span = spanOf(args, columns);
    if (!span) return null;
    const [first, last] = span;
    return columns
      ? {
          deleted: { firstRow: 0, firstCol: first, lastRow: MAX_ROW, lastCol: last },
          shifted: { firstRow: 0, firstCol: first, lastRow: MAX_ROW, lastCol: MAX_COL },
        }
      : {
          deleted: { firstRow: first, firstCol: 0, lastRow: last, lastCol: MAX_COL },
          shifted: { firstRow: first, firstCol: 0, lastRow: MAX_ROW, lastCol: MAX_COL },
        };
  }
  const range =
    typeof args[0] === 'string'
      ? scope.engine.parseReference(args[0]).range
      : (args[0] as CellRect | undefined);
  if (!range) return null;
  return {
    deleted: range,
    shifted:
      args[1] === CellShift.Left ? { ...range, lastCol: MAX_COL } : { ...range, lastRow: MAX_ROW },
  };
};

const within = (rect: CellRect, row: number, col: number): boolean =>
  row >= rect.firstRow && row <= rect.lastRow && col >= rect.firstCol && col <= rect.lastCol;

/** Every sheet-scope codec a delete can alter, with the store slice each re-hydrates. */
const SHEET_SLICES: readonly (readonly [JournalCodec, FacadeChangeKind])[] = [
  [CF, 'cf'],
  [VALIDATION, 'validation'],
  [SHEET_PROT, 'protection'],
  [PRINT, 'print'],
  [SHEET_CHROME, 'view'],
  [COMMENTS, 'comments'],
  [THREADS, 'threadedComments'],
  [SPARKLINES, 'sparklineGroups'],
  [AUTOFILTER, 'autoFilter'],
  [TABLES, 'table'],
  [rowsColsCodec, 'layout'],
  [DRAWINGS, 'drawing'],
  [PIVOTS, 'pivot'],
];

/** A delete's before-state: every sheet-scope slice, the names, the removed cells, every formula
 *  of the workbook and the styles of the shifted region. */
const captureDelete = (scope: JournalScope): Restore | null => {
  const { handle, name, args } = scope;
  const sheet = scope.sheet as number;
  const rects = deleteRects(scope);
  if (!rects) return null;
  const slices: Restore[] = [];
  for (const codec of [...SHEET_SLICES.map(([c]) => c), NAMES]) {
    const restore = codec.capture(scope);
    if (!restore) return null;
    slices.push(restore);
  }
  const removed = [...handle.cells(sheet)].filter((c) =>
    within(rects.deleted, c.addr.row, c.addr.col),
  );
  const formulas = collectAllFormulas(handle);
  const styles = handle.snapshotRangeStyle(rangeOfRect(sheet, rects.shifted));
  if (!styles) return null;
  const insert = DELETE_INVERSE[name] as string;
  const insertArgs = inverseArgs(name, args);
  return () => {
    invoke(handle.engineWorkbook().sheet(sheet), insert, insertArgs);
    for (const restore of slices) restore();
    for (const c of removed) writeCell(handle, c.addr, c.value, c.formula);
    for (const entry of formulas) handle.setFormula(entry.addr, entry.formula);
    handle.restoreRangeStyle(sheet, styles);
    for (const [, kind] of SHEET_SLICES) handle.noteFacadeChange({ kind, sheet });
    handle.noteFacadeChange({ kind: 'names', sheet: null });
  };
};

/** Scopes whose first capture (before the write) has run; the second one is the after-state. */
const captured = new WeakSet<JournalScope>();

/** Insert, delete and cell-shift writes. Undo of an insert is the matching delete; undo of a delete
 *  restores the sheet's slices, the removed cells and every formula. Redo repeats the write. */
const STRUCTURE: JournalCodec = {
  id: 'structure',
  scopeKey: (scope) => `structure:${scope.sheet}`,
  capture(scope) {
    const { handle, name, args } = scope;
    const sheet = scope.sheet as number;
    if (captured.has(scope)) {
      return () => invoke(handle.engineWorkbook().sheet(sheet), name, args);
    }
    captured.add(scope);
    const inverse = INSERT_INVERSE[name];
    if (inverse) {
      const undoArgs = inverseArgs(name, args);
      return () => invoke(handle.engineWorkbook().sheet(sheet), inverse, undoArgs);
    }
    return captureDelete(scope);
  },
};

type SheetAddEvent = Extract<HandleChangeEvent, { kind: 'sheet-add' | 'sheet-copy' }>;

/** The entry of a write that added the sheet at `index`: undo removes it, redo adds it again
 *  through `repeat` and announces it as the first write did. */
export const sheetAddedEntry = (
  handle: WorkbookHandle,
  index: number,
  announce: (added: EngineSheet) => SheetAddEvent,
  repeat: () => EngineSheet,
): HistoryEntry => ({
  undo: () =>
    handle.withBatchedRecalc(() => {
      const wb = handle.engineWorkbook();
      wb.removeSheet(wb.sheet(index));
      handle.noteSheetChange({ kind: 'sheet-remove', index });
    }),
  redo: () => handle.withBatchedRecalc(() => handle.noteSheetChange(announce(repeat()))),
});

const moveSheetTo = (handle: WorkbookHandle, from: number, to: number): void => {
  if (from === to) return;
  handle.withBatchedRecalc(() => {
    const wb = handle.engineWorkbook();
    const next = to > from ? to + 1 : to;
    wb.sheet(from).move(next >= wb.listSheets().length ? null : next);
    handle.noteSheetChange({ kind: 'sheet-move', from, to });
  });
};

/** The entry of a move from `from` to `to`. */
export const sheetMovedEntry = (
  handle: WorkbookHandle,
  from: number,
  to: number,
): HistoryEntry => ({
  undo: () => moveSheetTo(handle, to, from),
  redo: () => moveSheetTo(handle, from, to),
});

/** The sheet's name now; the returned function renames it back and announces the rename. */
export const captureSheetName = (handle: WorkbookHandle, index: number): Restore => {
  const name = handle.engineWorkbook().sheet(index).name;
  return () => {
    const sheet = handle.engineWorkbook().sheet(index);
    sheet.rename(name);
    handle.noteSheetChange({ kind: 'sheet-rename', index, name: sheet.name });
  };
};

export const STRUCTURE_JOURNAL: JournalFamily = {
  members: {
    sheet: Object.fromEntries(
      STRUCTURAL_MEMBERS.map((name): [string, JournalChoice] => [name, STRUCTURE]),
    ),
  },
};
