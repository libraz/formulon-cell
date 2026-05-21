import { addrKey } from '../engine/address.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { CellFormat, SpreadsheetStore, State } from '../store/store.js';
import { coerceInputForCell, writeCoerced } from './coerce-input.js';
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
  // Operate column by column so consecutive runs land in the same target columns.
  for (let c = range.c0; c <= range.c1; c += 1) {
    for (let r = range.r0; r <= range.r1; r += 1) {
      const sourceKey = addrKey({ sheet: range.sheet, row: r, col: c });
      const cell = state.data.cells.get(sourceKey);
      if (!cell) continue;
      const v = cell.value;
      if (v.kind !== 'text') continue;
      const tokens = splitText(v.value, delimiter, options);
      if (tokens.length < 2) continue;
      maxTokens = Math.max(maxTokens, tokens.length);
      const sourceFormat = cloneFormat(state.format.formats.get(sourceKey));
      for (let t = 0; t < tokens.length; t += 1) {
        const tok = tokens[t] ?? '';
        const dst = { sheet: range.sheet, row: r, col: c + t };
        if (!isCellWritable(state, dst)) {
          warnProtected(dst);
          continue;
        }
        writeCoerced(wb, dst, coerceInputForCell(state, dst, tok));
        if (sourceFormat) {
          formatWrites.push({ key: addrKey(dst), format: sourceFormat });
        }
      }
    }
  }
  if (formatWrites.length > 0) {
    store.setState((s) => {
      const formats = new Map(s.format.formats);
      for (const write of formatWrites) formats.set(write.key, { ...write.format });
      return { ...s, format: { ...s.format, formats } };
    });
  }
  wb.recalc();
  return maxTokens;
}
