import { afterEach, describe, expect, it, vi } from 'vitest';
import * as cellShift from '../../../src/commands/cell-shift.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings/en.js';
import {
  disposeCellShiftDialog,
  runCellShiftFlow,
} from '../../../src/interact/cell-shift-dialog.js';
import { disposeOverlayPortal } from '../../../src/interact/overlay-portal.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

describe('runCellShiftFlow', () => {
  const host = document.createElement('div');

  afterEach(() => {
    disposeCellShiftDialog(host);
    disposeOverlayPortal(host);
    document.body.replaceChildren();
    vi.restoreAllMocks();
  });

  it('applies the chosen shift over the range and reports a change', async () => {
    document.body.appendChild(host);
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    const store = createSpreadsheetStore();
    const insertCells = vi.spyOn(cellShift, 'insertCells').mockReturnValue(true);
    const onApplied = vi.fn();
    const target = {
      store,
      wb,
      history: null,
      origin: 'ribbon',
      commandId: 'insertCells',
    } as const;
    const range = { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 2 };

    runCellShiftFlow(host, en, 'insert', target, range, onApplied);
    document.querySelector<HTMLButtonElement>('.fc-cellshift__button--primary')?.click();

    expect(insertCells).toHaveBeenCalledWith(store, wb, null, range, 'down');
    expect(onApplied).toHaveBeenCalledTimes(1);
  });

  it('reports nothing when the shift changes nothing', async () => {
    document.body.appendChild(host);
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    const store = createSpreadsheetStore();
    vi.spyOn(cellShift, 'deleteCells').mockReturnValue(false);
    const onApplied = vi.fn();
    const target = {
      store,
      wb,
      history: null,
      origin: 'ribbon',
      commandId: 'deleteCells',
    } as const;

    runCellShiftFlow(
      host,
      en,
      'delete',
      target,
      { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      onApplied,
    );
    document.querySelector<HTMLButtonElement>('.fc-cellshift__button--primary')?.click();

    expect(onApplied).not.toHaveBeenCalled();
  });
});
