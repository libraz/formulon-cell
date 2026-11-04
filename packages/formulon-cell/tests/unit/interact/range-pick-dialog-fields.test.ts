import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import {
  disposeMacConsolidateDialog,
  openMacConsolidateDialog,
} from '../../../src/interact/mac-consolidate-dialog.js';
import {
  disposeMacGoalSeekDialog,
  openMacGoalSeekDialog,
} from '../../../src/interact/mac-goal-seek-dialog.js';
import { mutators } from '../../../src/store/store.js';
import {
  byAction,
  byId,
  createMacDataDialogFixture,
  type MacDataDialogFixture,
} from './mac-data-dialog-fixture.js';

const pickerButton = (inputId: string): HTMLButtonElement => {
  const btn = byId(inputId).closest('.fc-range-picker')?.querySelector('.fc-range-picker__btn');
  if (!btn) throw new Error(`missing picker for #${inputId}`);
  return btn as HTMLButtonElement;
};

const waitFor = async (predicate: () => boolean): Promise<void> => {
  for (let i = 0; i < 200 && !predicate(); i += 1) await new Promise((r) => setTimeout(r, 10));
  expect(predicate()).toBe(true);
};

describe('Goal Seek reference fields', () => {
  let fx: MacDataDialogFixture;

  beforeEach(async () => {
    fx = await createMacDataDialogFixture();
    fx.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 3);
    fx.workbook.setFormula({ sheet: 0, row: 0, col: 1 }, '=A1*2');
    fx.workbook.recalc();
    openMacGoalSeekDialog(fx.instance);
  });

  afterEach(() => {
    disposeMacGoalSeekDialog(fx.instance);
    fx.dispose();
  });

  it('writes absolute single cells and OK still solves with them', async () => {
    const formula = byId<HTMLInputElement>('fc-mac-goalseek-formula-cell');
    const changing = byId<HTMLInputElement>('fc-mac-goalseek-changing-cell');
    pickerButton('fc-mac-goalseek-formula-cell').click();
    mutators.setActive(fx.store, { sheet: 0, row: 0, col: 1 });
    expect(formula.value).toBe('$B$1');
    pickerButton('fc-mac-goalseek-changing-cell').click();
    mutators.setActive(fx.store, { sheet: 0, row: 0, col: 0 });
    expect(changing.value).toBe('$A$1');
    byId<HTMLInputElement>('fc-mac-goalseek-target-value').value = '10';
    byAction('goal-seek-ok').click();
    await waitFor(
      () =>
        document.querySelector<HTMLElement>('.fc-mac-goalseek__status')?.dataset.state ===
        'success',
    );
    byAction('goal-seek-ok').click();
    expect(fx.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'number',
      value: 5,
    });
  });

  it('qualifies picks once the active sheet has left the opening sheet, and OK resolves them', () => {
    fx.workbook.addSheet('Sheet2');
    const formula = byId<HTMLInputElement>('fc-mac-goalseek-formula-cell');
    pickerButton('fc-mac-goalseek-formula-cell').click();
    mutators.setSheetIndex(fx.store, 1);
    mutators.setActive(fx.store, { sheet: 1, row: 0, col: 0 });
    expect(formula.value).toBe('Sheet2!$A$1');
    mutators.setSheetIndex(fx.store, 0);
    mutators.setActive(fx.store, { sheet: 0, row: 2, col: 2 });
    expect(formula.value).toBe('Sheet1!$C$3');
  });
});

describe('Consolidate reference field', () => {
  let fx: MacDataDialogFixture;

  beforeEach(async () => {
    fx = await createMacDataDialogFixture();
    fx.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    fx.workbook.setNumber({ sheet: 0, row: 1, col: 0 }, 2);
    fx.workbook.recalc();
    fx.store.setState((s) => ({
      ...s,
      selection: {
        ...s.selection,
        active: { sheet: 0, row: 0, col: 2 },
        range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
      },
    }));
    openMacConsolidateDialog(fx.instance);
  });

  afterEach(() => {
    disposeMacConsolidateDialog(fx.instance);
    fx.dispose();
  });

  it('collapses a single cell, writes a drag as an absolute range, and OK consolidates it', () => {
    const source = byId<HTMLInputElement>('fc-mac-consolidate-sources');
    pickerButton('fc-mac-consolidate-sources').click();
    mutators.setRange(fx.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    expect(source.value).toBe('$A$1');
    mutators.setRange(fx.store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 });
    expect(source.value).toBe('$A$1:$A$2');
    byAction('consolidate-ok').click();
    expect(fx.workbook.getValue({ sheet: 0, row: 1, col: 2 })).toEqual({
      kind: 'number',
      value: 2,
    });
  });

  it('qualifies a pick on another sheet', () => {
    fx.workbook.addSheet('Sheet2');
    const source = byId<HTMLInputElement>('fc-mac-consolidate-sources');
    pickerButton('fc-mac-consolidate-sources').click();
    mutators.setSheetIndex(fx.store, 1);
    mutators.setRange(fx.store, { sheet: 1, r0: 0, c0: 0, r1: 0, c1: 0 });
    expect(source.value).toBe('Sheet2!$A$1');
  });
});
