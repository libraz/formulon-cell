import { CellShift } from '@libraz/formulon';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { mount, type SpreadsheetInstance } from '../../../src/api/spreadsheet/instance.js';
import { spreadsheetInternals } from '../../../src/mount/internals.js';
import type { SpreadsheetStore } from '../../../src/store/store.js';
import { createHostElement, installDomStubs, uninstallDomStubs } from '../../test-utils/dom.js';

describe('facade layout writes keep the store in step with the engine', () => {
  let instance: SpreadsheetInstance;
  let store: SpreadsheetStore;
  let cleanup: () => void;

  beforeEach(async () => {
    installDomStubs();
    let host: HTMLElement;
    ({ host, cleanup } = createHostElement());
    instance = await mount(host);
    store = spreadsheetInternals(instance).store;
    const s = instance.workbook.sheet(0);
    s.setValues('A1', [
      ['Region', 'Sales'],
      ['East', 120],
      ['West', 90],
      ['North', 75],
    ]);
    s.addTable({ range: 'A1:B4', name: 'Sales' });
    instance.workbook.table('Sales').update({ totalsRow: true });
  });
  afterEach(() => {
    instance.dispose();
    cleanup();
    uninstallDomStubs();
  });

  const tableRange = () => store.getState().tables.tables.find((t) => t.name === 'Sales')?.range;

  it('moves the table overlay, totals row included, when insertCells shifts it down', () => {
    expect(tableRange()).toMatchObject({ r0: 0, r1: 4, c0: 0, c1: 1 });
    instance.workbook.sheet(0).insertCells('A4:B4', CellShift.Down);
    expect(tableRange()).toMatchObject({ r0: 0, r1: 5, c0: 0, c1: 1 });
  });

  it('moves the table overlay when whole rows are inserted through it', () => {
    instance.workbook.sheet(0).insertRows(3, 3);
    expect(tableRange()).toMatchObject({ r0: 0, r1: 5 });
  });

  it('moves the table overlay back when rows are deleted from it', () => {
    instance.workbook.sheet(0).deleteRows(2, 2);
    expect(tableRange()).toMatchObject({ r0: 0, r1: 3 });
  });
});
