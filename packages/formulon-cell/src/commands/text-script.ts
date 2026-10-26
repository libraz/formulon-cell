import { parseAddrKey } from '../engine/address.js';
import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { rangeContainsAddr } from '../store/selection-geometry.js';
import type { State } from '../store/types.js';
import { applyTextScript, type ScriptCommand } from '../toolbar/review-tools.js';
import { blockDataTableCells } from './data-table-guard.js';
import { isCellWritable } from './protection.js';

export function applyTextScriptToRange(
  state: State,
  workbook: WorkbookHandle,
  range: Range,
  command: ScriptCommand,
): number {
  const writes: { addr: Addr; next: string | null }[] = [];
  for (const [key, cell] of state.data.cells) {
    const addr = parseAddrKey(key);
    if (!addr || !rangeContainsAddr(range, addr)) continue;
    if (!isCellWritable(state, addr)) continue;
    if (command === 'clear') {
      if (cell.value.kind !== 'blank' || cell.formula) writes.push({ addr, next: null });
      continue;
    }
    if (cell.value.kind !== 'text') continue;
    const next = applyTextScript(cell.value.value, command);
    if (next !== cell.value.value) writes.push({ addr, next });
  }
  if (
    blockDataTableCells(
      workbook,
      writes.map((write) => write.addr),
    )
  )
    return 0;
  return workbook.withBatchedRecalc(() => {
    for (const { addr, next } of writes) {
      if (next === null) workbook.setBlank(addr);
      else workbook.setText(addr, next);
    }
    return writes.length;
  });
}
