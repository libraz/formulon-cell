import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import {
  disposeMacPivotMoveDialog,
  openMacPivotMoveDialog,
} from '../../../src/interact/mac-pivot-move-dialog.js';
import { mutators } from '../../../src/store/store.js';
import {
  byAction,
  byId,
  createMacDataDialogFixture,
  type MacDataDialogFixture,
} from './mac-data-dialog-fixture.js';

describe('Move PivotTable dialog', () => {
  let fx: MacDataDialogFixture;

  beforeEach(async () => {
    fx = await createMacDataDialogFixture();
    const wb = fx.workbook;
    wb.setText({ sheet: 0, row: 0, col: 0 }, 'Region');
    wb.setText({ sheet: 0, row: 0, col: 1 }, 'Sales');
    wb.setText({ sheet: 0, row: 1, col: 0 }, 'East');
    wb.setNumber({ sheet: 0, row: 1, col: 1 }, 3);
    wb.setText({ sheet: 0, row: 2, col: 0 }, 'West');
    wb.setNumber({ sheet: 0, row: 2, col: 1 }, 4);
    createPivotTableFromRange(wb, {
      source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 },
      destination: { sheet: 0, row: 2, col: 4 },
      name: 'Report',
      rowField: 'Region',
      valueField: 'Sales',
    });
    mutators.setActive(fx.store, { sheet: 0, row: 3, col: 4 });
  });

  afterEach(() => {
    disposeMacPivotMoveDialog(fx.instance);
    fx.dispose();
  });

  it('does not open outside a report', () => {
    mutators.setActive(fx.store, { sheet: 0, row: 20, col: 20 });
    expect(openMacPivotMoveDialog(fx.instance)).toBe(false);
  });

  it('starts on the existing-worksheet choice with the report top-left as location', () => {
    expect(openMacPivotMoveDialog(fx.instance)).toBe(true);
    const location = byId<HTMLInputElement>('fc-mac-pivotmove-location');
    expect(location.value).toBe('Sheet1!E3');
    expect((byAction('pivot-move-existing') as unknown as HTMLInputElement).checked).toBe(true);
    byAction('pivot-move-new').click();
    expect(location.disabled).toBe(true);
    expect(location.value).toBe('Sheet1!E3');
  });

  it('moves the report to the typed location on OK', () => {
    openMacPivotMoveDialog(fx.instance);
    byId<HTMLInputElement>('fc-mac-pivotmove-location').value = 'Sheet1!H3';
    byAction('pivot-move-ok').click();
    expect(fx.workbook.pivotBounds(0)[0]?.range).toMatchObject({ r0: 2, c0: 7 });
  });
});
