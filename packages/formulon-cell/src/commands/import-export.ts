import { addrKey, parseAddrKey } from '../engine/address.js';
/**
 * High-level import/export commands for tabular text formats. The lower-level
 * encoders / parsers live next to clipboard so paste-as-CSV stays trivial.
 */

import type { Addr, Range } from '../engine/types.js';
import { formatBoolean, formatEditableNumber } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { isSingleCell, rangeArea } from '../store/selection-geometry.js';
import type { SpreadsheetStore, State } from '../store/store.js';
import { authorizedWrite } from './authorized-write.js';
import {
  type CSVEncodeOptions,
  type CSVParseOptions,
  encodeCSV,
  parseCSV,
} from './clipboard/csv.js';
import {
  canonicalizeLocalizedFormulaForInput,
  coerceInputForCell,
  inputLocaleForWorkbook,
  writeCoerced,
} from './coerce-input.js';
import {
  blockDataTableEdit,
  blockTableFunction,
  dataTableGuardAlert,
  tableFunctionAlert,
} from './data-table-guard.js';
import { History } from './history.js';
import {
  appliedBatch,
  type ChangeBatchResult,
  type OperationEffect,
  refusedBatch,
  refusedByDecision,
} from './interaction-policy.js';
import { isCellWritable, warnProtected } from './protection.js';

/** The most cells one CSV import or export covers. */
export const MAX_CSV_CELLS = 100_000;

export interface ImportCSVOptions extends CSVParseOptions {
  /**
   * Locale used to coerce numeric and boolean fields. When omitted, the
   * workbook's active profile supplies the input locale.
   */
  locale?: string;
}

/** The first cell of the parsed `rows` at `origin` that sheet protection refuses, or null. */
const firstRefusedCell = (state: State, rows: readonly string[][], origin: Addr): Addr | null => {
  for (let r = 0; r < rows.length; r += 1) {
    for (let c = 0; c < (rows[r]?.length ?? 0); c += 1) {
      const addr = { sheet: origin.sheet, row: origin.row + r, col: origin.col + c };
      if (!isCellWritable(state, addr)) return addr;
    }
  }
  return null;
};

/**
 * Parse a CSV blob and write the resulting cells starting at `anchor` (or the
 * active cell when omitted). Each value runs through `coerceInput` so a leading
 * `=` becomes a formula, `"42"` becomes a number, etc. — same semantics as
 * paste-from-clipboard, but for full-document loads. One authorized `paste` of the target range
 * from the instance API, all or nothing: refused when the interaction policy or sheet protection
 * refuses any target cell, and undone as one step.
 */
export function importCSV(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  text: string,
  anchor?: Addr,
  history?: History | null,
  options: ImportCSVOptions = {},
): ChangeBatchResult {
  const rows = text ? parseCSV(text, options) : [];
  if (rows.length === 0) return appliedBatch([], null);

  const state = store.getState();
  const origin: Addr = anchor ?? state.selection.active;
  const sheet = origin.sheet;
  const width = rows.reduce((widest, row) => Math.max(widest, row.length), 0);
  const target: Range = {
    sheet,
    r0: origin.row,
    c0: origin.col,
    r1: origin.row + rows.length - 1,
    c1: origin.col + Math.max(0, width - 1),
  };
  if (blockDataTableEdit(wb, target))
    return refusedBatch({ code: 'protected', reason: dataTableGuardAlert(wb).message });
  // One write, all or nothing: a target cell the protected sheet refuses refuses the import.
  const refused = firstRefusedCell(state, rows, origin);
  if (refused) {
    warnProtected(refused);
    return refusedBatch({ addr: refused, code: 'protected', reasonKey: 'codeProtected' });
  }
  const inputLocale = options.locale ?? inputLocaleForWorkbook(wb);
  const csvFormulas = rows
    .flat()
    .filter((raw) => raw.trim().startsWith('='))
    .map((raw) => canonicalizeLocalizedFormulaForInput(wb, raw.trim(), inputLocale));
  if (blockTableFunction(wb, csvFormulas))
    return refusedBatch({ code: 'invalid', reason: tableFunctionAlert(wb).message });

  const effect: OperationEffect = {
    kind: 'range',
    range: target,
    ...(csvFormulas.length > 0 ? { includesFormula: true } : {}),
  };
  let changed: Addr[] = [];
  const result = authorizedWrite({
    store,
    history: history ?? new History(),
    intents: (intentOrigin) => [
      { operation: 'paste', origin: intentOrigin, commandId: 'importCSV', effects: [effect] },
    ],
    origin: 'instanceApi',
    write: () => {
      changed = wb.trackChangedCells(() =>
        wb.withBatchedRecalc(() => {
          for (let r = 0; r < rows.length; r += 1) {
            const cells = rows[r] ?? [];
            for (let c = 0; c < cells.length; c += 1) {
              const addr: Addr = { sheet, row: origin.row + r, col: origin.col + c };
              writeCoerced(
                wb,
                addr,
                coerceInputForCell(state, addr, cells[c] ?? '', inputLocale, wb),
              );
            }
          }
        }),
      ).changed;
    },
  });
  return result.ok ? appliedBatch(changed, target) : refusedByDecision(result.decision);
}

export interface ExportOptions extends CSVEncodeOptions {
  /** Range to export. Defaults to the current selection. When the selection is
   *  a single cell, the entire used region of the sheet is exported instead. */
  range?: Range;
  /**
   * Locale used for editable number and boolean output. When omitted, the
   * existing canonical CSV representation is preserved.
   */
  locale?: string;
}

/**
 * Serialize a range to CSV. Values come from the displayed text (formulas
 * collapse to their last computed result) — same rules as TSV copy. When no
 * explicit range is supplied and the selection is a single cell, the export
 * covers the bounding box of populated cells on the active sheet.
 */
export function exportCSV(state: State, opts: ExportOptions = {}): string {
  const selected = opts.range ?? selectionOrUsed(state);
  const range =
    selected && rangeArea(selected) > MAX_CSV_CELLS ? usedRangeInRange(state, selected) : selected;
  if (!range) return '';
  if (rangeArea(range) > MAX_CSV_CELLS) return '';
  const { sheet } = range;

  const grid: string[][] = [];
  for (let row = range.r0; row <= range.r1; row += 1) {
    const line: string[] = [];
    for (let col = range.c0; col <= range.c1; col += 1) {
      line.push(displayValue(state, sheet, row, col, opts.locale));
    }
    grid.push(line);
  }
  return encodeCSV(grid, { delimiter: opts.delimiter, eol: opts.eol, bom: opts.bom });
}

function selectionOrUsed(state: State): Range | null {
  const sel = state.selection.range;
  const isSingle = isSingleCell(sel);
  if (!isSingle) return sel;
  return usedRange(state);
}

/** Bounding box of every populated cell on the active sheet. Returns null
 *  when the sheet is empty. */
function usedRange(state: State): Range | null {
  const sheet = state.data.sheetIndex;
  return usedRangeInRange(state, {
    sheet,
    r0: Number.NEGATIVE_INFINITY,
    c0: Number.NEGATIVE_INFINITY,
    r1: Number.POSITIVE_INFINITY,
    c1: Number.POSITIVE_INFINITY,
  });
}

function usedRangeInRange(state: State, range: Range): Range | null {
  const sheet = range.sheet;
  let r0 = Number.POSITIVE_INFINITY;
  let r1 = Number.NEGATIVE_INFINITY;
  let c0 = Number.POSITIVE_INFINITY;
  let c1 = Number.NEGATIVE_INFINITY;
  let any = false;
  const visit = (key: string): void => {
    const addr = parseAddrKey(key);
    if (!addr) return;
    if (addr.sheet !== sheet) return;
    if (addr.row < range.r0 || addr.row > range.r1) return;
    if (addr.col < range.c0 || addr.col > range.c1) return;
    if (addr.row < r0) r0 = addr.row;
    if (addr.row > r1) r1 = addr.row;
    if (addr.col < c0) c0 = addr.col;
    if (addr.col > c1) c1 = addr.col;
    any = true;
  };
  for (const [key, cell] of state.data.cells) {
    if (cell.value.kind === 'blank' && !cell.formula) continue;
    visit(key);
  }
  for (const [key, format] of state.format.formats) {
    if (typeof format.hyperlinkDisplay !== 'string' || format.hyperlinkDisplay.length === 0)
      continue;
    visit(key);
  }
  if (!any) return null;
  return { sheet, r0, c0, r1, c1 };
}

function displayValue(
  state: State,
  sheet: number,
  row: number,
  col: number,
  locale?: string,
): string {
  const key = addrKey({ sheet, row, col });
  const cell = state.data.cells.get(key);
  const hyperlinkDisplay = state.format.formats.get(key)?.hyperlinkDisplay;
  if (!cell) return hyperlinkDisplay ?? '';
  switch (cell.value.kind) {
    case 'number':
      return locale === undefined
        ? String(cell.value.value)
        : formatEditableNumber(cell.value.value, locale);
    case 'text':
      return cell.value.value;
    case 'bool':
      return locale === undefined
        ? cell.value.value
          ? 'TRUE'
          : 'FALSE'
        : formatBoolean(cell.value.value, locale);
    case 'error':
      return cell.value.text;
    default:
      return hyperlinkDisplay ?? '';
  }
}
