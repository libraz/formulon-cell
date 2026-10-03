import { parseAddrKey } from '../engine/address.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { rangeContainsAddr } from '../store/selection-geometry.js';
import type { State } from '../store/types.js';
import { applyTextScript, type ScriptCommand } from '../toolbar/review-tools.js';
import { isCellWritable } from './protection.js';

export function applyTextScriptToRange(
  state: State,
  workbook: WorkbookHandle,
  range: Range,
  command: ScriptCommand,
): number {
  return workbook.withBatchedRecalc(() => {
    let changed = 0;
    for (const [key, cell] of state.data.cells) {
      const addr = parseAddrKey(key);
      if (!addr || !rangeContainsAddr(range, addr)) continue;
      if (!isCellWritable(state, addr)) continue;
      if (command === 'clear') {
        if (cell.value.kind === 'blank' && !cell.formula) continue;
        workbook.setBlank(addr);
        changed += 1;
        continue;
      }
      if (cell.value.kind !== 'text') continue;
      const next = applyTextScript(cell.value.value, command);
      if (next === cell.value.value) continue;
      workbook.setText(addr, next);
      changed += 1;
    }
    return changed;
  });
}
