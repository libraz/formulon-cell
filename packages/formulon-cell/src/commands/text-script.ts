import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { rangeContainsAddr } from '../store/selection-geometry.js';
import type { State } from '../store/types.js';
import { applyTextScript, type ScriptCommand } from '../toolbar/review-tools.js';
import { isCellWritable } from './protection.js';

const addrFromKey = (key: string): Addr | null => {
  const parts = key.split(':').map(Number);
  const sheet = parts[0];
  const row = parts[1];
  const col = parts[2];
  if (
    typeof sheet !== 'number' ||
    typeof row !== 'number' ||
    typeof col !== 'number' ||
    !Number.isInteger(sheet) ||
    !Number.isInteger(row) ||
    !Number.isInteger(col)
  ) {
    return null;
  }
  return { sheet, row, col };
};

export function applyTextScriptToRange(
  state: State,
  workbook: WorkbookHandle,
  range: Range,
  command: ScriptCommand,
): number {
  let changed = 0;
  for (const [key, cell] of state.data.cells) {
    const addr = addrFromKey(key);
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
}
