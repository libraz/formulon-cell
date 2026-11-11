import {
  type ArrayFormulaInput,
  type CellAddress,
  type CellRect,
  type CellValueInput,
  type DataTableInput,
  type Sheet as EngineSheet,
  type ErrorCode,
  ValueKind,
} from '@libraz/formulon';
import { type CSVDelimiter, encodeCSV, parseCSV } from '../commands/clipboard/csv.js';
import {
  type CoercedInput,
  canonicalizeLocalizedFormulaForInput,
  coerceInput,
  inputLocaleForWorkbook,
  workbookDate1904,
} from '../commands/coerce-input.js';
import { MAX_CSV_CELLS } from '../commands/import-export.js';
import { appliedBatch } from '../commands/interaction-policy.js';
import { isMacWorkbook } from '../engine/spreadsheet-profile.js';
import { formatBoolean, formatEditableNumber } from '../engine/value.js';
import type { ChangeBatchResult } from './change-batch.js';
import { spreadsheetFailure } from './errors.js';
import type { MemberSpec, SheetContext } from './facade-context.js';

/** Options of `Sheet.importCsv`. */
export interface ImportCsvOptions {
  /** The cell the first field lands in. */
  readonly at: string | CellAddress;
  /** The locale numbers, dates, logical values and function names are read in; the workbook's by default. */
  readonly locale?: string;
  /** The field separator; `,` by default. */
  readonly delimiter?: CSVDelimiter;
}

/** Options of `Sheet.exportCsv`. */
export interface ExportCsvOptions {
  /** The cells to write; the sheet's used range by default. */
  readonly range?: string | CellRect;
  /** The locale numbers and logical values are written in; the workbook's by default. */
  readonly locale?: string;
  /** The field separator; `,` by default. */
  readonly delimiter?: CSVDelimiter;
}

const assertWithinLimit = (rows: number, cols: number, member: string): void => {
  const cells = rows * cols;
  if (cells <= MAX_CSV_CELLS) return;
  throw spreadsheetFailure(
    'limit-exceeded',
    `${member}: ${cells} cells exceed the limit of ${MAX_CSV_CELLS}`,
    { context: { limit: MAX_CSV_CELLS } },
  );
};

export const writeCoerced = (
  ctx: SheetContext,
  { row, col }: CellAddress,
  input: CoercedInput,
  locale: string | undefined,
): void => {
  const sheet = ctx.sheetEngine;
  switch (input.kind) {
    case 'blank':
      sheet.setBlank(row, col);
      break;
    case 'formula':
      sheet.setFormula(
        row,
        col,
        canonicalizeLocalizedFormulaForInput(ctx.handle, input.text, locale),
      );
      break;
    case 'number':
      sheet.setNumber(row, col, input.value);
      break;
    case 'bool':
      sheet.setBool(row, col, input.value);
      break;
    case 'error':
      sheet.setError(row, col, input.code as ErrorCode);
      break;
    case 'text':
      sheet.setText(row, col, input.value);
      break;
  }
};

function importCsv(ctx: SheetContext, text: string, options: ImportCsvOptions): ChangeBatchResult {
  if (typeof text !== 'string') throw new TypeError('importCsv: text must be a string');
  if (options?.at === undefined) throw new TypeError('importCsv: options.at is required');
  const anchor = ctx.cellAt(options.at);
  const rows = parseCSV(text, { delimiter: options.delimiter });
  const width = rows.reduce((widest, row) => Math.max(widest, row.length), 0);
  assertWithinLimit(rows.length, width, 'importCsv');

  const sheet = ctx.sheetEngine.index;
  const locale = options.locale ?? inputLocaleForWorkbook(ctx.handle);
  const coerceOptions = {
    locale,
    date1904: workbookDate1904(ctx.handle),
    mac: isMacWorkbook(ctx.handle),
  };
  const { changed } = ctx.handle.trackChangedCells(() =>
    ctx.run(() => {
      rows.forEach((fields, r) => {
        fields.forEach((raw, c) => {
          const at = { row: anchor.row + r, col: anchor.col + c };
          const forceText = ctx.sheetEngine.getCellFormat(at.row, at.col).numberFormat === '@';
          const input = coerceInput(raw, { ...coerceOptions, forceText });
          ctx.cellWrite([at.row, at.col], () => writeCoerced(ctx, at, input, locale));
        });
      });
    }),
  );
  const range =
    width === 0
      ? null
      : {
          sheet,
          r0: anchor.row,
          c0: anchor.col,
          r1: anchor.row + rows.length - 1,
          c1: anchor.col + width - 1,
        };
  return appliedBatch(changed, range);
}

function exportCsv(ctx: SheetContext, options: ExportCsvOptions = {}): string {
  const sheet = ctx.sheetEngine;
  const rect = options.range === undefined ? sheet.getUsedRange() : ctx.rectOf(options.range);
  if (rect === null) return '';
  assertWithinLimit(
    rect.lastRow - rect.firstRow + 1,
    rect.lastCol - rect.firstCol + 1,
    'exportCsv',
  );
  const locale = options.locale ?? inputLocaleForWorkbook(ctx.handle);
  const field = (row: number, col: number): string => {
    const value = sheet.getTypedValue(row, col);
    switch (value.kind) {
      case ValueKind.Number:
        return locale === undefined
          ? String(value.number)
          : formatEditableNumber(value.number, locale);
      case ValueKind.Text:
        return value.text;
      case ValueKind.Bool:
        return locale === undefined
          ? value.boolean
            ? 'TRUE'
            : 'FALSE'
          : formatBoolean(value.boolean, locale);
      case ValueKind.Error:
        return sheet.getDisplayText(row, col).text;
      default:
        return '';
    }
  };
  const grid: string[][] = [];
  for (let row = rect.firstRow; row <= rect.lastRow; row += 1) {
    const line: string[] = [];
    for (let col = rect.firstCol; col <= rect.lastCol; col += 1) line.push(field(row, col));
    grid.push(line);
  }
  return encodeCSV(grid, { delimiter: options.delimiter });
}

/** `setValues(row, col, values)` or `setValues(a1, values)`: the block anchored at the cell. */
function setValues(ctx: SheetContext, ...args: unknown[]): void {
  const values = args.at(-1) as readonly (readonly CellValueInput[])[];
  const anchor =
    typeof args[0] === 'string'
      ? ctx.cellAt(args[0])
      : ctx.cellAt({ row: args[0] as number, col: args[1] as number });
  const width = values.reduce((widest, row) => Math.max(widest, row.length), 0);
  if (values.length === 0 || width === 0) {
    ctx.run(() => (ctx.sheetEngine.setValues as (...a: unknown[]) => void)(...args));
    return;
  }
  const rect = {
    firstRow: anchor.row,
    firstCol: anchor.col,
    lastRow: anchor.row + values.length - 1,
    lastCol: anchor.col + width - 1,
  };
  ctx.rangeWrite(rect, () => (ctx.sheetEngine.setValues as (...a: unknown[]) => void)(...args));
}

function setArrayFormula(ctx: SheetContext, input: ArrayFormulaInput): void {
  ctx.rangeWrite(ctx.rectOf(input.range), () => ctx.sheetEngine.setArrayFormula(input));
}

function addDataTable(ctx: SheetContext, spec: DataTableInput): CellAddress {
  return ctx.rangeWrite(ctx.rectOf(spec.range), () => ctx.sheetEngine.addDataTable(spec));
}

/** `removeDataTable(row, col)` or `removeDataTable(a1)`: the table whose result block holds the cell. */
function removeDataTable(ctx: SheetContext, ...args: unknown[]): void {
  const at =
    typeof args[0] === 'string'
      ? ctx.cellAt(args[0])
      : ctx.cellAt({ row: args[0] as number, col: args[1] as number });
  const remove = (): void => ctx.sheetEngine.removeDataTable(at.row, at.col);
  const table = ctx.sheetEngine.getDataTable(at.row, at.col);
  if (table === null) ctx.run(remove);
  else ctx.rangeWrite(table.range, remove);
}

/** Cell values, formulas, evaluation and CSV: engine `Sheet` members of the same name, plus CSV. */
export const SHEET_CELL_MEMBERS = {
  read: [
    'getValue',
    'getValues',
    'getTypedValue',
    'getTypedValues',
    'getDate',
    'getDisplayText',
    'getFormula',
    'getFormulaR1C1',
    'getArrayFormula',
    'getLambdaText',
    'getSpill',
    'getDataTable',
    'getUsedRange',
    'getCellsInRange',
    'listCells',
    'iterRows',
    'getPrecedents',
    'getDependents',
    'getCellPhonetic',
    'getCellPhoneticRuns',
    'getCellPhoneticProperties',
    'getCellShowPhonetic',
    'evaluateFormula',
    'evaluateFormulaArray',
    'evaluateTypedFormula',
    'evaluateTypedFormulaArray',
  ],
  cellWrite: ['setValue', 'setNumber', 'setText', 'setBool', 'setError', 'setBlank', 'setFormula'],
  change: {
    format: [
      'setCellPhonetic',
      'setCellPhoneticRuns',
      'setCellPhoneticProperties',
      'setCellShowPhonetic',
    ],
  },
  methods: {
    setValues,
    setArrayFormula,
    addDataTable,
    removeDataTable,
    importCsv,
    exportCsv,
  },
} as const satisfies MemberSpec<SheetContext>;

type CellMemberName =
  | (typeof SHEET_CELL_MEMBERS.read)[number]
  | (typeof SHEET_CELL_MEMBERS.cellWrite)[number]
  | (typeof SHEET_CELL_MEMBERS.change.format)[number]
  | 'setValues'
  | 'setArrayFormula'
  | 'addDataTable'
  | 'removeDataTable';

declare module './workbook.js' {
  interface Sheet extends Pick<EngineSheet, CellMemberName> {
    /**
     * Reads CSV text into the sheet from `options.at`, the way typing each field would: a leading
     * `=` makes a formula, numbers, dates and logical values are recognised in `options.locale`,
     * and a cell formatted as text keeps the field as text. Host code is not gated by sheet
     * protection: every field is written. One undo step and one `change` event. The import the
     * user's policy and protection decide is `SpreadsheetInstance.importCsv`.
     *
     * @throws FormulonError `limit-exceeded` above 100,000 cells.
     */
    importCsv(text: string, options: ImportCsvOptions): ChangeBatchResult;
    /**
     * Writes the values of `options.range` (the used range by default) as CSV text with `\r\n` line
     * ends; an empty sheet gives `''`.
     *
     * @example Export what the user has selected
     * ```ts
     * const { sheet, range } = instance.selection.get();
     * const csv = instance.workbook.sheet(sheet).exportCsv({ range });
     * ```
     *
     * @throws FormulonError `limit-exceeded` above 100,000 cells.
     */
    exportCsv(options?: ExportCsvOptions): string;
  }
}
