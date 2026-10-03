import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import {
  disposeMacGoalSeekDialog,
  openMacGoalSeekDialog,
} from '../../../src/interact/mac-goal-seek-dialog.js';
import {
  byAction,
  byId,
  createMacDataDialogFixture,
  type MacDataDialogFixture,
  pressEnter,
} from './mac-data-dialog-fixture.js';

const status = (): HTMLElement => {
  const el = document.querySelector<HTMLElement>('.fc-mac-goalseek__status');
  if (!el) throw new Error('missing status');
  return el;
};

const waitFor = async (predicate: () => boolean): Promise<void> => {
  for (let i = 0; i < 200 && !predicate(); i += 1) await new Promise((r) => setTimeout(r, 10));
  expect(predicate()).toBe(true);
};

describe('Goal Seek dialog', () => {
  let fx: MacDataDialogFixture;

  beforeEach(async () => {
    fx = await createMacDataDialogFixture();
    fx.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 3);
    fx.workbook.setFormula({ sheet: 0, row: 0, col: 1 }, '=A1*2');
    fx.workbook.recalc();
    openMacGoalSeekDialog(fx.instance);
    byId<HTMLInputElement>('fc-mac-goalseek-formula-cell').value = 'B1';
    byId<HTMLInputElement>('fc-mac-goalseek-changing-cell').value = 'A1';
    byId<HTMLInputElement>('fc-mac-goalseek-target-value').value = '10';
  });

  afterEach(() => {
    disposeMacGoalSeekDialog(fx.instance);
    fx.dispose();
  });

  it('runs on Enter inside a text input', () => {
    const event = pressEnter(byId('fc-mac-goalseek-target-value'));
    expect(event.defaultPrevented).toBe(true);
    expect(status().textContent).not.toBe('');
  });

  it('ignores Enter on the Cancel button', () => {
    const event = pressEnter(byAction('goal-seek-cancel'));
    expect(event.defaultPrevented).toBe(false);
    expect(status().textContent).toBe('');
    expect(byId<HTMLInputElement>('fc-mac-goalseek-target-value').disabled).toBe(false);
  });

  it('ignores Enter while an IME composition is active', () => {
    pressEnter(byId('fc-mac-goalseek-target-value'), { isComposing: true });
    expect(status().textContent).toBe('');
  });

  it('re-enables the inputs when committing a stale solution fails', async () => {
    byAction('goal-seek-ok').click();
    await waitFor(() => status().dataset.state === 'success');
    expect(byId<HTMLInputElement>('fc-mac-goalseek-target-value').disabled).toBe(true);

    fx.workbook.setNumber({ sheet: 0, row: 5, col: 5 }, 1);
    fx.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 4);
    byAction('goal-seek-ok').click();

    expect(status().dataset.state).toBe('error');
    expect(byId<HTMLInputElement>('fc-mac-goalseek-target-value').disabled).toBe(false);
    expect(byId<HTMLInputElement>('fc-mac-goalseek-formula-cell').disabled).toBe(false);

    // A new run solves again instead of retrying the stale commit.
    byAction('goal-seek-ok').click();
    await waitFor(() => status().dataset.state === 'success');
  });

  it('relabels the dialog when the locale changes', () => {
    expect(document.querySelector('.fc-mac-goalseek .fc-fmtdlg__header')?.textContent).toBe(
      'Goal Seek',
    );
    fx.i18n.setLocale('ja');
    expect(document.querySelector('.fc-mac-goalseek .fc-fmtdlg__header')?.textContent).toBe(
      'ゴール シーク',
    );
  });
});
