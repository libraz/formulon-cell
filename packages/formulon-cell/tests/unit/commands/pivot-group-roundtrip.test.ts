import { afterEach, describe, expect, it } from 'vitest';
import { restorePivotFieldLayout } from '../../../src/commands/pivot-edit.js';
import { recreatePivotTable } from '../../../src/commands/pivot-recreate.js';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const ROWS: (string | number)[][] = [
  ['Region', 'Qty', 'Amt'],
  ['North', 1, 100],
  ['North', 2, 200],
  ['South', 3, 300],
  ['South', 4, 400],
  ['East', 5, 500],
];
const QTY = 1;

const open: WorkbookHandle[] = [];
afterEach(() => {
  for (const wb of open.splice(0)) wb.dispose();
});

const setup = async () => {
  const wb = await WorkbookHandle.createDefault();
  open.push(wb);
  for (const [row, cells] of ROWS.entries()) {
    for (const [col, value] of cells.entries()) {
      if (typeof value === 'number') wb.setNumber({ sheet: 0, row, col }, value);
      else wb.setText({ sheet: 0, row, col }, value);
    }
  }
  const created = createPivotTableFromRange(wb, {
    source: { sheet: 0, r0: 0, c0: 0, r1: 5, c1: 2 },
    destination: { sheet: 0, row: 0, col: 5 },
    rowField: 'Qty',
    valueField: 'Amt',
  });
  if (!created.ok) throw new Error(`pivot fixture failed: ${created.reason}`);
  return { wb, pivotIndex: created.pivotIndex };
};

describe('pivot field grouping read back from the engine', () => {
  it('reads a numeric grouping with its fixed and automatic bounds', async () => {
    const { wb, pivotIndex } = await setup();
    expect(wb.getPivotField(0, pivotIndex, QTY)?.numericGroup).toBeNull();
    expect(
      wb.setPivotFieldNumericGroup(0, pivotIndex, QTY, { interval: 2, start: 1, end: null }),
    ).toBe(true);
    expect(wb.getPivotField(0, pivotIndex, QTY)?.numericGroup).toEqual({
      interval: 2,
      start: 1,
      end: null,
    });
  });

  it('keeps the numeric grouping when a removed report is rebuilt from its definition', async () => {
    const { wb, pivotIndex } = await setup();
    wb.setPivotFieldNumericGroup(0, pivotIndex, QTY, { interval: 2, start: 1, end: 4 });
    const definition = wb.getPivotDefinition(0, pivotIndex);
    if (!definition) throw new Error('definition unreadable');
    expect(wb.removePivotTable(0, pivotIndex)).toBe(true);
    const rebuilt = recreatePivotTable(wb, 0, { row: 0, col: 5 }, definition);
    expect(rebuilt).toBeGreaterThanOrEqual(0);
    expect(wb.getPivotDefinition(0, rebuilt)?.fields).toEqual(definition.fields);
  });

  it('undoes and redoes a numeric grouping by restoring the field layout', async () => {
    const { wb, pivotIndex } = await setup();
    const ungrouped = wb.getPivotDefinition(0, pivotIndex);
    wb.setPivotFieldNumericGroup(0, pivotIndex, QTY, { interval: 2, start: 1, end: 4 });
    const grouped = wb.getPivotDefinition(0, pivotIndex);
    if (!ungrouped || !grouped) throw new Error('definition unreadable');
    expect(grouped.fields.length).toBe(ungrouped.fields.length + 1);

    expect(restorePivotFieldLayout(wb, 0, pivotIndex, ungrouped)).toBe(true);
    expect(wb.getPivotDefinition(0, pivotIndex)?.fields).toEqual(ungrouped.fields);
    expect(restorePivotFieldLayout(wb, 0, pivotIndex, grouped)).toBe(true);
    expect(wb.getPivotDefinition(0, pivotIndex)?.fields).toEqual(grouped.fields);
  });
});
