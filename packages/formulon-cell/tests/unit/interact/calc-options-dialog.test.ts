import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { bindCalcSettingsStore } from '../../../src/engine/calc-settings-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings.js';
import { attachIterativeDialog } from '../../../src/interact/iterative-dialog.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const canLoadWasm = (): boolean => typeof WebAssembly !== 'undefined';

const q = <T extends HTMLElement>(selector: string): T => {
  const el = document.querySelector<T>(selector);
  if (!el) throw new Error(`missing ${selector}`);
  return el;
};
const option = (name: string): HTMLInputElement =>
  q<HTMLInputElement>(`.fc-iterdlg [data-calc-option="${name}"]`);
const toggle = (name: string, checked: boolean): void => {
  const input = option(name);
  input.checked = checked;
  input.dispatchEvent(new Event('change', { bubbles: true }));
};
const ok = (): void => q<HTMLButtonElement>('.fc-iterdlg .fc-iterdlg__btn--primary').click();
const confirmDialog = (): HTMLElement | null =>
  document.querySelector<HTMLElement>('.fc-tb__dlg[role="alertdialog"]');
const settle = (): Promise<void> => new Promise((resolve) => setTimeout(resolve, 0));

describe.skipIf(!canLoadWasm())('Calculation Options dialog', () => {
  let host: HTMLElement;
  let wb: WorkbookHandle;
  let unbind: () => void;
  const store = createSpreadsheetStore();
  const history = new History();

  beforeEach(async () => {
    host = document.createElement('div');
    document.body.appendChild(host);
    wb = await WorkbookHandle.createDefault();
    wb.attachHistory(history);
    unbind = bindCalcSettingsStore(wb, store);
  });

  afterEach(() => {
    unbind();
    wb.dispose();
    mutators.setR1C1(store, false);
    while (document.body.firstChild) document.body.removeChild(document.body.firstChild);
  });

  const open = () => {
    const handle = attachIterativeDialog({
      host,
      getWb: () => wb,
      store,
      history,
      strings: en,
    });
    handle.open();
    return handle;
  };

  it('opens on the settings the workbook carries', () => {
    wb.setCalcMode(1);
    wb.setCalcSetting('calcOnSave', false);
    mutators.setR1C1(store, true);
    wb.setCalcSetting('date1904', true);
    wb.setCalcSetting('fullPrecision', false);
    const handle = open();
    expect(option('manual').checked).toBe(true);
    expect(option('auto').checked).toBe(false);
    expect(option('calcOnSave').checked).toBe(false);
    expect(option('calcOnSave').disabled).toBe(false);
    expect(option('r1c1').checked).toBe(true);
    expect(option('date1904').checked).toBe(true);
    expect(option('precisionAsDisplayed').checked).toBe(true);
    handle.detach();
  });

  it('only offers Recalculate before saving under Manual', () => {
    const handle = open();
    expect(option('auto').checked).toBe(true);
    expect(option('calcOnSave').disabled).toBe(true);
    expect(option('calcOnSave').dataset.disabledReason).toBe(
      en.iterativeDialog.recalcBeforeSaveRequiresManual,
    );
    toggle('manual', true);
    expect(option('calcOnSave').disabled).toBe(false);
    handle.detach();
  });

  it('writes every option on OK and the workbook saves them', async () => {
    wb.setFormula({ sheet: 0, row: 0, col: 0 }, '=DATE(2024,1,1)');
    const handle = open();
    toggle('auto-no-table', true);
    toggle('r1c1', true);
    toggle('date1904', true);
    ok();
    expect(wb.calcMode()).toBe(2);
    expect(store.getState().ui.r1c1).toBe(true);
    expect(store.getState().ui.date1904).toBe(true);
    // The engine marks date formulas dirty; the switch recalcs them.
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 43830 });
    expect(store.getState().data.cells.get('0:0:0')?.value).toEqual({
      kind: 'number',
      value: 43830,
    });
    const restored = await WorkbookHandle.loadBytes(wb.save());
    try {
      expect(restored.calcMode()).toBe(2);
      expect(restored.calcSettings()).toMatchObject({ r1c1: true, date1904: true });
    } finally {
      restored.dispose();
    }
    handle.detach();
  });

  it('confirms Set precision as displayed and rounds constants on OK', async () => {
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1.23456);
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        numFmt: { kind: 'fixed', decimals: 2 },
      },
    );
    const handle = open();
    toggle('precisionAsDisplayed', true);
    const dialog = confirmDialog();
    expect(dialog?.textContent).toContain(en.iterativeDialog.precisionWarning);
    q<HTMLButtonElement>('.fc-tb__dlg[role="alertdialog"] .fc-fmtdlg__btn--primary').click();
    await settle();
    expect(option('precisionAsDisplayed').checked).toBe(true);
    ok();
    expect(wb.calcSettings()?.fullPrecision).toBe(false);
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 1.23 });
    handle.detach();
  });

  it('leaves the option off when the warning is cancelled', async () => {
    const handle = open();
    toggle('precisionAsDisplayed', true);
    const buttons = document.querySelectorAll<HTMLButtonElement>(
      '.fc-tb__dlg[role="alertdialog"] .fc-fmtdlg__btn',
    );
    [...buttons].find((b) => !b.classList.contains('fc-fmtdlg__btn--primary'))?.click();
    await settle();
    expect(option('precisionAsDisplayed').checked).toBe(false);
    ok();
    expect(wb.calcSettings()?.fullPrecision).toBe(true);
    handle.detach();
  });

  it('undoes Set precision as displayed as one step: values and the option', async () => {
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1.23456);
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        numFmt: { kind: 'fixed', decimals: 2 },
      },
    );
    history.clear();
    const handle = open();
    toggle('precisionAsDisplayed', true);
    q<HTMLButtonElement>('.fc-tb__dlg[role="alertdialog"] .fc-fmtdlg__btn--primary').click();
    await settle();
    ok();
    expect(store.getState().ui.precisionAsDisplayed).toBe(true);
    history.undo();
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 1.23456 });
    expect(wb.calcSettings()?.fullPrecision).toBe(true);
    expect(store.getState().ui.precisionAsDisplayed).toBe(false);
    expect(history.canUndo()).toBe(false);
    history.redo();
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 1.23 });
    expect(wb.calcSettings()?.fullPrecision).toBe(false);
    expect(store.getState().ui.precisionAsDisplayed).toBe(true);
    handle.detach();
  });
});
