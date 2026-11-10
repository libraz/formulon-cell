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

  it('uses a locale-aware decimal text input and rejects an empty target', async () => {
    fx.i18n.setLocale('de');
    const target = byId<HTMLInputElement>('fc-mac-goalseek-target-value');
    expect(target.type).toBe('text');
    expect(target.inputMode).toBe('decimal');

    target.value = '1,5';
    byAction('goal-seek-ok').click();
    await waitFor(() => status().dataset.state === 'success');
    byAction('goal-seek-cancel').click();

    openMacGoalSeekDialog(fx.instance);
    const reopenedTarget = byId<HTMLInputElement>('fc-mac-goalseek-target-value');
    reopenedTarget.value = '';
    byAction('goal-seek-ok').click();
    expect(status().dataset.state).toBe('error');
    expect(reopenedTarget.value).toBe('');
  });

  it('hydrates numeric targets with the active locale', () => {
    fx.i18n.setLocale('de');
    disposeMacGoalSeekDialog(fx.instance);
    openMacGoalSeekDialog(fx.instance);
    expect(byId<HTMLInputElement>('fc-mac-goalseek-target-value').value).toBe('3');

    fx.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 1234.5);
    disposeMacGoalSeekDialog(fx.instance);
    openMacGoalSeekDialog(fx.instance);
    expect(byId<HTMLInputElement>('fc-mac-goalseek-target-value').value).toBe('1234,5');
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

  it.each([
    { action: 'goal-seek-ok', a1: -335543.32 },
    { action: 'goal-seek-cancel', a1: 1 },
  ])('offers the last iterate of a non-converged run ($action)', async ({ action, a1 }) => {
    fx.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    fx.workbook.setFormula({ sheet: 0, row: 0, col: 1 }, '=5');
    fx.workbook.recalc();
    byId<HTMLInputElement>('fc-mac-goalseek-target-value').value = '7';
    byAction('goal-seek-ok').click();
    const notConverged = fx.instance.i18n.strings.mac.data.errors.notConverged;
    await waitFor(() => status().textContent?.startsWith(notConverged) ?? false);
    expect(status().dataset.state).toBe('error');
    expect(status().textContent).toContain('-335543.32');

    byAction(action).click();
    expect(fx.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'number',
      value: a1,
    });
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

  describe('Mac profile', () => {
    const reopenAsMac = (): void => {
      disposeMacGoalSeekDialog(fx.instance);
      fx.host.classList.add('fc-host');
      fx.host.dataset.fcPlatform = 'mac';
      fx.store.setState((s) => ({
        ...s,
        selection: { ...s.selection, active: { sheet: 0, row: 0, col: 0 } },
      }));
      openMacGoalSeekDialog(fx.instance);
    };

    it('prefills only the formula cell and has no iteration field', () => {
      reopenAsMac();
      expect(byId<HTMLInputElement>('fc-mac-goalseek-formula-cell').value).toBe('A1');
      expect(byId<HTMLInputElement>('fc-mac-goalseek-target-value').value).toBe('');
      expect(byId<HTMLInputElement>('fc-mac-goalseek-changing-cell').value).toBe('');
      const iterations = byId<HTMLInputElement>('fc-mac-goalseek-iterations');
      expect(iterations.closest<HTMLElement>('.fc-mac-goalseek__row')?.hidden).toBe(true);
    });

    it('labels the fields with the desktop Japanese wording', () => {
      fx.i18n.setLocale('ja');
      reopenAsMac();
      const labels = [...document.querySelectorAll('.fc-mac-goalseek__row')]
        .filter((row) => !(row as HTMLElement).hidden)
        .map((row) => row.querySelector('span')?.textContent);
      expect(labels).toEqual(['数式入力セル:', '目標値:', '変化させるセル:']);
    });

    it('offers range pickers on the two cell fields only', () => {
      reopenAsMac();
      const pickers = [
        ...document.querySelectorAll<HTMLElement>('.fc-mac-goalseek .fc-range-picker'),
      ];
      expect(pickers.map((p) => p.querySelector('input')?.id)).toEqual([
        'fc-mac-goalseek-formula-cell',
        'fc-mac-goalseek-changing-cell',
      ]);
    });
  });
});
