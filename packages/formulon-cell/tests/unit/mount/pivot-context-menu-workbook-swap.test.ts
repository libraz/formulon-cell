import '../../api/setup-dom.js';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { workbookHandleOf, workbookOf } from '../../../src/api/workbook.js';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import { PivotAggregation } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { mount } from '../../../src/index.js';
import { createHostElement, installDomStubs, uninstallDomStubs } from '../../test-utils/dom.js';
import { sheetFocusTarget } from '../../test-utils/sheet-focus.js';

describe('workbook swap after a PivotTable context menu', () => {
  beforeEach(() => installDomStubs());
  afterEach(() => uninstallDomStubs());

  it('swaps again after the pivot menu was opened and closed', async () => {
    const { host, cleanup } = createHostElement();
    const wb = await WorkbookHandle.createDefault();
    const rows: (string | number)[][] = [
      ['Kind', 'Amt'],
      ['a', 1],
      ['b', 2],
    ];
    for (const [r, row] of rows.entries()) {
      for (const [c, v] of row.entries()) {
        if (typeof v === 'number') wb.setNumber({ sheet: 0, row: r, col: c }, v);
        else wb.setText({ sheet: 0, row: r, col: c }, v);
      }
    }
    expect(
      createPivotTableFromRange(wb, {
        source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 },
        destination: { sheet: 0, row: 0, col: 3 },
        rowField: 'Kind',
        valueField: 'Amt',
        aggregation: PivotAggregation.Sum,
      }).ok,
    ).toBe(true);
    const instance = await mount(host, { workbook: workbookOf(wb, true), locale: 'en' });
    const errors: unknown[] = [];
    instance.on('error', (e) => errors.push(e));
    const swap = async (): Promise<void> => {
      const bytes = instance.workbook.save().bytes;
      const swapped = new Promise((resolve) => {
        instance.on('workbook', resolve);
        instance.on('error', resolve);
      });
      instance.update({ workbook: workbookOf(await WorkbookHandle.loadBytes(bytes), true) });
      await swapped;
    };
    try {
      await swap();
      const pivotCell = { sheet: 0, row: 1, col: 3 };
      expect(workbookHandleOf(instance.workbook).pivotCellInfo(pivotCell)).not.toBeNull();
      instance.selection.set('D2');
      sheetFocusTarget(host).dispatchEvent(
        new KeyboardEvent('keydown', { key: 'F10', shiftKey: true, bubbles: true }),
      );
      expect(document.querySelector('[data-fc-menu="pivot"]')).not.toBeNull();
      document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
      const before = instance.workbook;
      await swap();
      expect(errors).toEqual([]);
      expect(instance.workbook).not.toBe(before);
      expect(() => instance.workbook.save()).not.toThrow();
    } finally {
      instance.dispose();
      cleanup();
    }
  });
});
