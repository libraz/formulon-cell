import { addrKey, parseAddrKey } from '../engine/address.js';
import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { rangeContainsAddr } from '../store/selection-geometry.js';
import type { CellFormat, SpreadsheetStore, State } from '../store/store.js';
import { coerceInputForCell, inputLocaleForWorkbook, writeCoerced } from './coerce-input.js';
import { isCellWritable, warnProtected } from './protection.js';

const cloneFormat = (fmt: CellFormat | undefined): CellFormat | undefined =>
  fmt ? { ...fmt } : undefined;

export interface TextToColumnsOptions {
  collapseConsecutiveDelimiters?: boolean;
}

const splitText = (
  value: string,
  delimiters: string | readonly string[],
  options: TextToColumnsOptions = {},
): string[] => {
  const list = (Array.isArray(delimiters) ? delimiters : [delimiters]).filter(
    (delimiter) => delimiter !== '',
  );
  if (list.length === 0) return [value];
  if (list.length === 1 && !options.collapseConsecutiveDelimiters)
    return value.split(list[0] ?? '');
  const pattern = new RegExp(
    `(?:${list.map((delimiter) => delimiter.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|')})${
      options.collapseConsecutiveDelimiters ? '+' : ''
    }`,
  );
  return value.split(pattern);
};

/** The cells `textToColumns` writes for `range`: each text cell that splits, and
 *  the cells to its right that receive the extra pieces. */
export function textToColumnsTargets(
  state: State,
  range: Range,
  delimiter: string | readonly string[],
  options: TextToColumnsOptions = {},
): Range[] {
  const targets: Range[] = [];
  for (const [key, cell] of state.data.cells) {
    const addr = parseAddrKey(key);
    if (!addr || !rangeContainsAddr(range, addr) || cell.value.kind !== 'text') continue;
    const pieces = splitText(cell.value.value, delimiter, options).length;
    if (pieces < 2) continue;
    targets.push({
      sheet: range.sheet,
      r0: addr.row,
      c0: addr.col,
      r1: addr.row,
      c1: addr.col + pieces - 1,
    });
  }
  return targets;
}

/** Split each cell in `range` (column-major) by `delimiter` and write the
 *  resulting tokens to the cells immediately to the right of the source.
 *  Returns the maximum number of tokens produced. */
export function textToColumns(
  state: State,
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  range: Range,
  delimiter: string | readonly string[],
  options: TextToColumnsOptions = {},
): number {
  let maxTokens = 0;
  const formatWrites: Array<{ key: string; format: CellFormat }> = [];
  const candidates = [...state.data.cells.entries()]
    .map(([key, cell]) => ({ key, cell, addr: parseAddrKey(key) }))
    .filter((entry): entry is typeof entry & { addr: Addr } => !!entry.addr)
    .filter((entry) => rangeContainsAddr(range, entry.addr))
    .sort((left, right) => left.addr.col - right.addr.col || left.addr.row - right.addr.row);
  const inputLocale = inputLocaleForWorkbook(wb);
  // Operate column by column so consecutive runs land in the same target columns.
  wb.withBatchedRecalc(() => {
    for (const { key: sourceKey, cell, addr: sourceAddr } of candidates) {
      const v = cell.value;
      if (v.kind !== 'text') continue;
      const tokens = splitText(v.value, delimiter, options);
      if (tokens.length < 2) continue;
      maxTokens = Math.max(maxTokens, tokens.length);
      const sourceFormat = cloneFormat(state.format.formats.get(sourceKey));
      for (let t = 0; t < tokens.length; t += 1) {
        const tok = tokens[t] ?? '';
        const dst = { sheet: range.sheet, row: sourceAddr.row, col: sourceAddr.col + t };
        if (!isCellWritable(state, dst)) {
          warnProtected(dst);
          continue;
        }
        writeCoerced(wb, dst, coerceInputForCell(state, dst, tok, inputLocale, wb));
        if (sourceFormat) {
          formatWrites.push({ key: addrKey(dst), format: sourceFormat });
        }
      }
    }
  });
  if (formatWrites.length > 0) {
    store.setState((s) => {
      const formats = new Map(s.format.formats);
      for (const write of formatWrites) formats.set(write.key, { ...write.format });
      return { ...s, format: { ...s.format, formats } };
    });
  }
  wb.recalcAuto();
  return maxTokens;
}
