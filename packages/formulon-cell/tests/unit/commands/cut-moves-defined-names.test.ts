// @vitest-environment node

import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { copy } from '../../../src/commands/clipboard/copy.js';
import { pasteSpecial } from '../../../src/commands/clipboard/paste-special.js';
import { captureSnapshotFromCopyResult } from '../../../src/commands/clipboard/snapshot.js';
import { History } from '../../../src/commands/history.js';
import type { Addr, Range } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const at = (row: number, col: number, sheet = 0): Addr => ({ sheet, row, col });

describe('cut-paste moves defined names', () => {
  let wb: WorkbookHandle;
  let store: SpreadsheetStore;
  let history: History;

  const cutPaste = (source: Range, dest: Addr): void => {
    mutators.replaceCells(store, wb.cells(source.sheet));
    mutators.setActive(store, at(source.r0, source.c0, source.sheet));
    mutators.setRange(store, source);
    const result = copy(store.getState());
    if (!result) throw new Error('missing copy result');
    const snap = captureSnapshotFromCopyResult(store.getState(), result, 'cut');
    if (!snap) throw new Error('missing cut snapshot');
    mutators.setCopyRange(store, source, 'cut');
    mutators.setActive(store, dest);
    mutators.setRange(store, {
      sheet: dest.sheet,
      r0: dest.row,
      c0: dest.col,
      r1: dest.row,
      c1: dest.col,
    });
    const transaction = history.begin();
    pasteSpecial(
      store.getState(),
      store,
      wb,
      snap,
      { what: 'all', operation: 'none', skipBlanks: false, transpose: false },
      history,
    );
    history.end(transaction);
  };
  const nameFormula = (name: string, scope = -1): string | undefined =>
    [...wb.definedNames()].find((n) => n.name === name && n.localSheetId === scope)?.formula;
  const num = (addr: Addr): unknown => {
    const v = wb.getValue(addr);
    return v.kind === 'number' ? v.value : v;
  };

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await WorkbookHandle.createDefault();
    wb.renameSheet(0, 'Names');
    wb.addSheet('Calc');
    wb.attachStore(store);
    history = new History();
    wb.attachHistory(history);
    for (let r = 2; r <= 6; r += 1) {
      wb.setNumber(at(r, 2), 30);
      wb.setNumber(at(r, 4), r - 1);
    }
    wb.setNumber(at(0, 0), 0.08);
    wb.setDefinedNameEntry('RateList', '=Names!$C$3:$C$7');
    wb.setDefinedNameEntry('LocalList', '=Names!$E$3:$E$7', 0);
    wb.setDefinedNameEntry('TaxRate', '=Names!$A$1');
    wb.setDefinedNameEntry('Partial', '=Names!$C$3:$C$9');
    wb.setFormula(at(1, 9, 1), '=SUM(RateList)');
    wb.setFormula(at(1, 10, 1), '=TaxRate*100');
    wb.recalc();
  });
  afterEach(() => wb.dispose());

  it('moves a workbook name whose range lies inside the cut', () => {
    cutPaste({ sheet: 0, r0: 2, c0: 2, r1: 6, c1: 2 }, at(19, 2));
    expect(nameFormula('RateList')).toMatch(/Names!\$C\$20:\$C\$24$/);
    expect(num(at(1, 9, 1))).toBe(150);
    expect(nameFormula('Partial')).toMatch(/Names!\$C\$3:\$C\$9$/);
  });

  it('moves a sheet-scoped name and a single-cell name', () => {
    cutPaste({ sheet: 0, r0: 2, c0: 4, r1: 6, c1: 4 }, at(19, 4));
    expect(nameFormula('LocalList', 0)).toMatch(/Names!\$E\$20:\$E\$24$/);
    cutPaste({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, at(19, 0));
    expect(nameFormula('TaxRate')).toMatch(/Names!\$A\$20$/);
    expect(num(at(1, 10, 1))).toBe(8);
  });

  it('undoes together with the cut', () => {
    cutPaste({ sheet: 0, r0: 2, c0: 2, r1: 6, c1: 2 }, at(19, 2));
    history.undo();
    expect(nameFormula('RateList')).toMatch(/Names!\$C\$3:\$C\$7$/);
    expect(num(at(1, 9, 1))).toBe(150);
  });
});
