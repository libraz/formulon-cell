import { describe, expect, it } from 'vitest';
import { commitCellInput } from '../../../src/commands/cell-input-commit.js';
import { clearSelectedContents } from '../../../src/commands/clear-contents.js';
import { pasteTSV } from '../../../src/commands/clipboard/paste.js';
import { setPivotGuardSink } from '../../../src/commands/pivot-guard.js';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en, ja } from '../../../src/i18n/strings.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const setup = async () => {
  const wb = await WorkbookHandle.createDefault();
  const rows: (string | number)[][] = [
    ['Region', 'Sales'],
    ['East', 12],
    ['West', 5],
  ];
  for (const [r, row] of rows.entries()) {
    for (const [c, value] of row.entries()) {
      const addr = { sheet: 0, row: r, col: c };
      if (typeof value === 'number') wb.setNumber(addr, value);
      else wb.setText(addr, value);
    }
  }
  createPivotTableFromRange(wb, {
    source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 },
    destination: { sheet: 0, row: 5, col: 0 },
    rowField: 'Region',
    valueField: 'Sales',
  });
  const store = createSpreadsheetStore();
  mutators.replaceCells(store, wb.cells(0));
  return { wb, store };
};

describe('PivotTable edit guard', () => {
  it('rejects an editor commit onto a report cell and leaves the report alone', async () => {
    const { wb, store } = await setup();
    const before = [...wb.pivotCells(0)].map((c) => c.value);
    const result = commitCellInput({
      store,
      wb,
      addr: { sheet: 0, row: 6, col: 0 },
      raw: 'typed',
      origin: 'editor',
    });
    expect(result).toMatchObject({
      status: 'rejected',
      alert: { severity: 'stop', message: en.pivotTableDialog.cannotChange },
    });
    expect([...wb.pivotCells(0)].map((c) => c.value)).toEqual(before);

    const outside = commitCellInput({
      store,
      wb,
      addr: { sheet: 0, row: 12, col: 0 },
      raw: 'ok',
      origin: 'editor',
    });
    expect(outside.status).toBe('applied');
    wb.dispose();
  });

  it('shows the alert in the UI language', async () => {
    const { wb, store } = await setup();
    setPivotGuardSink(wb, () => undefined, ja.pivotTableDialog.cannotChange);
    const result = commitCellInput({
      store,
      wb,
      addr: { sheet: 0, row: 5, col: 1 },
      raw: '1',
      origin: 'formulaBar',
    });
    expect(result).toMatchObject({
      status: 'rejected',
      alert: { message: 'ピボットテーブルのこの部分は変更できません。' },
    });
    wb.dispose();
  });

  it('rejects Delete over a report and alerts through the sink', async () => {
    const { wb, store } = await setup();
    const alerts: string[] = [];
    setPivotGuardSink(wb, (alert) => alerts.push(alert.message));
    mutators.setRange(store, { sheet: 0, r0: 4, c0: 0, r1: 9, c1: 1 });
    const result = clearSelectedContents({ store, workbook: wb });
    expect(result.status).toBe('rejected');
    expect(alerts).toEqual([en.pivotTableDialog.cannotChange]);
    expect([...wb.pivotCells(0)].length).toBeGreaterThan(0);
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'Region' });

    alerts.length = 0;
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 });
    expect(clearSelectedContents({ store, workbook: wb }).status).toBe('applied');
    expect(alerts).toEqual([]);
    wb.dispose();
  });

  it('refuses a paste that reaches a report', async () => {
    const { wb, store } = await setup();
    mutators.setActive(store, { sheet: 0, row: 4, col: 0 });
    expect(pasteTSV(store.getState(), wb, 'a\tb\nc\td')).toBeNull();
    expect(wb.getValue({ sheet: 0, row: 4, col: 0 }).kind).toBe('blank');
    mutators.setActive(store, { sheet: 0, row: 20, col: 0 });
    expect(pasteTSV(store.getState(), wb, 'a')).not.toBeNull();
    wb.dispose();
  });
});
