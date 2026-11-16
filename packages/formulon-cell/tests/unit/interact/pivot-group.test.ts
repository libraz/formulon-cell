import { PivotDatePeriod } from '@libraz/formulon';
import { afterEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { PivotAxis } from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  groupPivotField,
  pivotGroupSettings,
  pivotGroupTargetAt,
} from '../../../src/interact/pivot-group.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';
import { loadFixture } from '../engine/pivot-excel-parity/fixture.js';

// Sheet5 shows Date on rows (2024-01-15 .. 03-10 on rows 4-8); its cache holds dates up to
// 2024-07-22 (serial 45495).
const books: WorkbookHandle[] = [];
afterEach(() => {
  for (const wb of books.splice(0)) wb.dispose();
});

const setup = async () => {
  const wb = await loadFixture('pivot_value_date_filters.xlsx');
  books.push(wb);
  const sheet = Array.from({ length: wb.sheetCount }, (_, i) => i).find(
    (i) => wb.sheetName(i) === 'Sheet5',
  ) as number;
  const history = new History();
  const ctx = {
    wb,
    history,
    store: createSpreadsheetStore(),
    origin: 'ribbon' as const,
    commandId: 'pivotGroupSelection',
    onChanged: () => undefined,
  };
  const dateGroup = () => wb.getPivotDefinition(sheet, 0)?.fields[2]?.dateGroup ?? null;
  return { wb, sheet, history, ctx, dateGroup };
};

describe('PivotTable Grouping of a date field', () => {
  it('targets the field of an item label, and nothing for other cells', async () => {
    const { wb, sheet } = await setup();
    expect(pivotGroupTargetAt(wb, { sheet, row: 4, col: 0 })).toEqual({
      sheet,
      pivotIndex: 0,
      fieldIdx: 2,
      values: 'date',
    });
    expect(pivotGroupTargetAt(wb, { sheet, row: 4, col: 1 })).toBeNull();
    expect(pivotGroupTargetAt(wb, { sheet, row: 2, col: 0 })).toBeNull();
    expect(pivotGroupTargetAt(wb, { sheet, row: 8, col: 0 })).toBeNull();
  });

  it('opens on Months from the first date to the midnight after the last', async () => {
    const { wb, sheet } = await setup();
    const target = pivotGroupTargetAt(wb, { sheet, row: 4, col: 0 });
    if (!target) throw new Error('no target');
    expect(pivotGroupSettings(wb, target)).toEqual({
      kind: 'date',
      periods: PivotDatePeriod.Months,
      autoStart: true,
      autoEnd: true,
      start: 45306,
      end: 45496,
      intervalDays: 1,
    });
  });

  it('groups by several periods in one undo step, and redo brings every period back', async () => {
    const { wb, sheet, history, ctx, dateGroup } = await setup();
    const target = pivotGroupTargetAt(wb, { sheet, row: 4, col: 0 });
    if (!target) throw new Error('no target');
    const periods = PivotDatePeriod.Months | PivotDatePeriod.Quarters | PivotDatePeriod.Years;
    const settings = pivotGroupSettings(wb, target);
    if (settings?.kind !== 'date') throw new Error('no settings');
    expect(groupPivotField(ctx, target, { ...settings, periods })).toBe(true);
    expect(dateGroup()).toMatchObject({ periods, startSerial: -1, endSerial: -1 });
    history.undo();
    expect(dateGroup()).toBeNull();
    history.redo();
    expect(dateGroup()).toMatchObject({ periods });
  });

  it('reopens on the grouping it has, from a label of a group field too', async () => {
    const { wb, sheet, ctx } = await setup();
    const target = pivotGroupTargetAt(wb, { sheet, row: 4, col: 0 });
    const settings = target && pivotGroupSettings(wb, target);
    if (!target || settings?.kind !== 'date') throw new Error('no target');
    groupPivotField(ctx, target, {
      ...settings,
      periods: PivotDatePeriod.Days,
      intervalDays: 7,
      autoStart: false,
      start: 45300,
    });
    const label = Array.from(wb.pivotCells(sheet)).find((c) => c.kind === 1 && c.addr.col === 0);
    if (!label) throw new Error('no label');
    const again = pivotGroupTargetAt(wb, label.addr);
    expect(again?.fieldIdx).toBe(2);
    expect(again && pivotGroupSettings(wb, again)).toEqual({
      kind: 'date',
      periods: PivotDatePeriod.Days,
      autoStart: false,
      autoEnd: true,
      start: 45300,
      end: 45496,
      intervalDays: 7,
    });
  });

  it('leaves text and number fields to other routes', async () => {
    const { wb, sheet } = await setup();
    expect(wb.setPivotFieldAxis(sheet, 0, 0, PivotAxis.Row)).toBe(true);
    const region = Array.from(wb.pivotCells(sheet)).find(
      (c) => c.kind === 1 && c.fieldName === 'Region',
    );
    if (!region) throw new Error('no region label');
    expect(pivotGroupTargetAt(wb, region.addr)?.values).toBe('other');
  });
});
