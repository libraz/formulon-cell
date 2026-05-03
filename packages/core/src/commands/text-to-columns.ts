import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { addrKey } from '../engine/workbook-handle.js';
import type { State } from '../store/store.js';

/** Split each cell in `range` (column-major) by `delimiter` and write the
 *  resulting tokens to the cells immediately to the right of the source.
 *  Returns the maximum number of tokens produced. */
export function textToColumns(
  state: State,
  wb: WorkbookHandle,
  range: Range,
  delimiter: string,
): number {
  if (delimiter === '') return 0;
  let maxTokens = 0;
  // Operate column by column so consecutive runs land in the same target columns.
  for (let c = range.c0; c <= range.c1; c += 1) {
    for (let r = range.r0; r <= range.r1; r += 1) {
      const cell = state.data.cells.get(addrKey({ sheet: range.sheet, row: r, col: c }));
      if (!cell) continue;
      const v = cell.value;
      if (v.kind !== 'text') continue;
      const tokens = v.value.split(delimiter);
      if (tokens.length < 2) continue;
      maxTokens = Math.max(maxTokens, tokens.length);
      for (let t = 0; t < tokens.length; t += 1) {
        const tok = tokens[t] ?? '';
        const dst = { sheet: range.sheet, row: r, col: c + t };
        const num = Number(tok);
        if (tok !== '' && Number.isFinite(num)) wb.setNumber(dst, num);
        else wb.setText(dst, tok);
      }
    }
  }
  wb.recalc();
  return maxTokens;
}
