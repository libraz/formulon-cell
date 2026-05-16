import { addrKey } from '../engine/address.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { State } from '../store/types.js';
import { applyTextScript, type ScriptCommand } from '../toolbar/review-tools.js';
import { isCellWritable } from './protection.js';

export function applyTextScriptToRange(
  state: State,
  workbook: WorkbookHandle,
  range: Range,
  command: ScriptCommand,
): number {
  let changed = 0;
  for (let row = range.r0; row <= range.r1; row += 1) {
    for (let col = range.c0; col <= range.c1; col += 1) {
      const addr = { sheet: range.sheet, row, col };
      if (!isCellWritable(state, addr)) continue;
      const cell = state.data.cells.get(addrKey(addr));
      if (command === 'clear') {
        if (!cell || (cell.value.kind === 'blank' && !cell.formula)) continue;
        workbook.setBlank(addr);
        changed += 1;
        continue;
      }
      if (cell?.value.kind !== 'text') continue;
      const next = applyTextScript(cell.value.value, command);
      if (next === cell.value.value) continue;
      workbook.setText(addr, next);
      changed += 1;
    }
  }
  return changed;
}
