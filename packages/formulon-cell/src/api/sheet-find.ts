import {
  type CellAddress,
  type CellRect,
  type Sheet as EngineSheet,
  ValueKind,
} from '@libraz/formulon';
import { coerceInput, inputLocaleForWorkbook, workbookDate1904 } from '../commands/coerce-input.js';
import { applySubstitution, type FindOptions, matcherFor } from '../commands/find.js';
import { appliedBatch, refusedBatch } from '../commands/interaction-policy.js';
import { engineRefusalReasonKey } from '../commands/permission-reasons.js';
import { isFormulonError } from '../engine/result.js';
import { isMacWorkbook } from '../engine/spreadsheet-profile.js';
import { formatBoolean } from '../engine/value.js';
import type { ChangeBatchResult } from './change-batch.js';
import type { MemberSpec, SheetContext } from './facade-context.js';
import { writeCoerced } from './sheet-cells.js';

/** What `Sheet.findAll` looks for and where. */
export interface FindCellsOptions {
  /** The text to find: `?` is one character, `*` any run, `~` makes the next one literal. */
  readonly query: string;
  /** Upper and lower case differ. Off by default. */
  readonly matchCase?: boolean;
  /** The query must match the whole cell. Off by default. */
  readonly matchEntireCell?: boolean;
  /** A half-width voiced kana differs from its composed form. Off by default. */
  readonly matchByte?: boolean;
  /** What is searched: displayed values (the default), formulas (canonical, constants as
   *  entered), notes or threaded comments. */
  readonly lookIn?: 'values' | 'formulas' | 'notes' | 'comments';
  /** The order of the addresses: by rows (the default) or by columns. */
  readonly searchBy?: 'rows' | 'columns';
  /** The cells to search; by default the used range, or the whole sheet for notes and comments. */
  readonly range?: string | CellRect;
}

/** What `Sheet.replaceAll` replaces, and with what. */
export interface ReplaceCellsOptions extends Omit<FindCellsOptions, 'lookIn'> {
  /** The text each match becomes; with `matchEntireCell`, the whole cell content. */
  readonly replacement: string;
  /** Displayed values of constants (the default; formula cells are left alone) or formulas and
   *  constants as entered. */
  readonly lookIn?: 'values' | 'formulas';
}

const findOptionsOf = (options: FindCellsOptions): FindOptions => ({
  query: options.query,
  caseSensitive: options.matchCase ?? false,
  matchWhole: options.matchEntireCell ?? false,
  matchWidth: options.matchByte ?? false,
});

/** A constant as it was entered: numbers and logical values unformatted. */
const enteredText = (sheet: EngineSheet, row: number, col: number): string => {
  const value = sheet.getTypedValue(row, col);
  switch (value.kind) {
    case ValueKind.Number:
      return String(value.number);
    case ValueKind.Bool:
      return formatBoolean(value.boolean);
    case ValueKind.Text:
      return value.text;
    case ValueKind.Error:
      return sheet.getDisplayText(row, col).text;
    default:
      return '';
  }
};

const searchRect = (ctx: SheetContext, range: FindCellsOptions['range']): CellRect | null =>
  range === undefined ? ctx.sheetEngine.getUsedRange() : ctx.rectOf(range);

const inRect = (rect: CellRect, at: CellAddress): boolean =>
  at.row >= rect.firstRow &&
  at.row <= rect.lastRow &&
  at.col >= rect.firstCol &&
  at.col <= rect.lastCol;

const ordered = (cells: CellAddress[], by: FindCellsOptions['searchBy']): CellAddress[] =>
  cells.sort((a, b) =>
    by === 'columns' ? a.col - b.col || a.row - b.row : a.row - b.row || a.col - b.col,
  );

/** Every cell of `rect`, in `by` order. */
function* cellsOf(rect: CellRect, by: FindCellsOptions['searchBy']): Generator<CellAddress> {
  const columns = by === 'columns';
  const outer = columns ? [rect.firstCol, rect.lastCol] : [rect.firstRow, rect.lastRow];
  const inner = columns ? [rect.firstRow, rect.lastRow] : [rect.firstCol, rect.lastCol];
  for (let i = outer[0] ?? 0; i <= (outer[1] ?? -1); i += 1) {
    for (let j = inner[0] ?? 0; j <= (inner[1] ?? -1); j += 1) {
      yield columns ? { row: j, col: i } : { row: i, col: j };
    }
  }
}

const cellText = (
  sheet: EngineSheet,
  at: CellAddress,
  lookIn: 'values' | 'formulas',
): string | null => {
  if (lookIn === 'formulas') {
    const formula = sheet.getFormula(at.row, at.col);
    return formula ?? enteredText(sheet, at.row, at.col);
  }
  return sheet.getDisplayText(at.row, at.col).text;
};

function findAll(ctx: SheetContext, options: FindCellsOptions): CellAddress[] {
  if (typeof options?.query !== 'string') throw new TypeError('findAll: options.query is required');
  const matches = matcherFor(findOptionsOf(options));
  const lookIn = options.lookIn ?? 'values';
  if (options.query === '') return [];
  const sheet = ctx.sheetEngine;
  if (lookIn === 'notes' || lookIn === 'comments') {
    // A note or a thread can sit on an empty cell, so these search the whole sheet by default.
    const within = options.range === undefined ? null : ctx.rectOf(options.range);
    const anchored = (cell: CellAddress): boolean => within === null || inRect(within, cell);
    return annotatedCells(sheet, lookIn, anchored, matches, options.searchBy);
  }
  const rect = searchRect(ctx, options.range);
  if (rect === null) return [];
  const found: CellAddress[] = [];
  for (const at of cellsOf(rect, options.searchBy)) {
    const text = cellText(sheet, at, lookIn);
    if (text && matches(text)) found.push(at);
  }
  return found;
}

const annotatedCells = (
  sheet: EngineSheet,
  lookIn: 'notes' | 'comments',
  anchored: (cell: CellAddress) => boolean,
  matches: (text: string) => boolean,
  searchBy: FindCellsOptions['searchBy'],
): CellAddress[] => {
  if (lookIn === 'notes') {
    const cells = sheet
      .listComments()
      .filter((note) => anchored(note.cell) && matches(note.text))
      .map((note) => ({ row: note.cell.row, col: note.cell.col }));
    return ordered(cells, searchBy);
  }
  const seen = new Set<string>();
  const cells: CellAddress[] = [];
  for (const comment of sheet.listThreadedComments()) {
    const key = `${comment.cell.row}:${comment.cell.col}`;
    if (seen.has(key) || !anchored(comment.cell) || !matches(comment.text)) continue;
    seen.add(key);
    cells.push({ row: comment.cell.row, col: comment.cell.col });
  }
  return ordered(cells, searchBy);
};

function replaceAll(ctx: SheetContext, options: ReplaceCellsOptions): ChangeBatchResult {
  if (typeof options?.replacement !== 'string') {
    throw new TypeError('replaceAll: options.replacement is required');
  }
  const lookIn = options.lookIn ?? 'values';
  const find = findOptionsOf(options);
  const sheet = ctx.sheetEngine;
  const targets = findAll(ctx, { ...options, lookIn }).filter(
    (at) => lookIn === 'formulas' || sheet.getFormula(at.row, at.col) === null,
  );
  const locale = inputLocaleForWorkbook(ctx.handle);
  const coerceOptions = {
    locale,
    date1904: workbookDate1904(ctx.handle),
    mac: isMacWorkbook(ctx.handle),
  };
  const written: CellAddress[] = [];
  let writing: CellAddress | null = null;
  let tracked: ReturnType<typeof ctx.handle.trackChangedCells>;
  try {
    tracked = ctx.handle.trackChangedCells(() =>
      ctx.run(() => {
        for (const at of targets) {
          writing = at;
          const current = cellText(sheet, at, lookIn) ?? '';
          const next = applySubstitution(current, find, options.replacement);
          if (next === current) continue;
          // Formulas are canonical both ways; a constant is read as typed in the workbook locale.
          if (lookIn === 'formulas' && next.startsWith('=')) {
            ctx.cellWrite([at.row, at.col], () => sheet.setFormula(at.row, at.col, next));
          } else {
            const input = coerceInput(next, coerceOptions);
            ctx.cellWrite([at.row, at.col], () => writeCoerced(ctx, at, input, locale));
          }
          written.push(at);
        }
      }),
    );
  } catch (error) {
    // The batch restored every cell it wrote; report the cell the engine refused.
    if (!isFormulonError(error) || writing === null) throw error;
    const at: CellAddress = writing;
    return refusedBatch({
      addr: { sheet: sheet.index, row: at.row, col: at.col },
      code: 'invalid',
      reason: error.message,
      reasonKey: engineRefusalReasonKey(error.parseError),
    });
  }
  const { changed } = tracked;
  if (written.length === 0) return appliedBatch(changed, null);
  const rows = written.map((at) => at.row);
  const cols = written.map((at) => at.col);
  return appliedBatch(changed, {
    sheet: sheet.index,
    r0: Math.min(...rows),
    c0: Math.min(...cols),
    r1: Math.max(...rows),
    c1: Math.max(...cols),
  });
}

/** Find and replace on one sheet: the host's analogue of `Range.Find` and `Range.Replace`. */
export const SHEET_FIND_MEMBERS = {
  methods: { findAll, replaceAll },
} as const satisfies MemberSpec<SheetContext>;

declare module './workbook.js' {
  interface Sheet {
    /**
     * The cells whose displayed value (or formula, note or threaded comment) the query finds, in
     * row or column order. The query compares as the Find dialog does: full-width and half-width
     * forms and hiragana and katakana are the same, and case is ignored unless `matchCase`.
     */
    findAll(options: FindCellsOptions): CellAddress[];
    /**
     * Replaces what the query finds in each cell, as Replace All does, without asking sheet
     * protection: host code is not gated. A replacement that breaks a formula writes nothing and
     * is reported as `refused`, naming that cell with code `invalid`. One undo step and one
     * `change` event.
     */
    replaceAll(options: ReplaceCellsOptions): ChangeBatchResult;
  }
}
