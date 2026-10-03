import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';

export interface FormulaRecord {
  addr: Addr;
  formula: string;
}

/** Every formula cell in the workbook, across all sheets. */
export function collectAllFormulas(wb: WorkbookHandle): FormulaRecord[] {
  const out: FormulaRecord[] = [];
  for (let sheet = 0; sheet < wb.sheetCount; sheet += 1) {
    for (const c of wb.cells(sheet)) {
      if (c.formula !== null) out.push({ addr: c.addr, formula: c.formula });
    }
  }
  return out;
}
