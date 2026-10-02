import { afterEach, describe, expect, it } from 'vitest';

import { getRecentFunctions } from '../../../src/commands/function-history.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { defaultStrings } from '../../../src/i18n/strings.js';
import { mutators } from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/index.js';

/**
 * Per-feature opt-out tests. `presets.*` covers bundle behavior; this spec
 * locks down that EACH individual flag (set to `false` while the others
 * remain default) drops just its own attach.
 */
describe('mount/host-features — individual feature flags', () => {
  let sheet: MountedStubSheet | undefined;

  afterEach(() => {
    sheet?.dispose();
    sheet = undefined;
  });

  it('formatDialog: false omits the format dialog handle', async () => {
    sheet = await mountStubSheet({ features: { formatDialog: false } });
    expect(sheet.instance.features.formatDialog).toBeFalsy();
  });

  it('findReplace: false omits the find/replace handle', async () => {
    sheet = await mountStubSheet({ features: { findReplace: false } });
    expect(sheet.instance.features.findReplace).toBeFalsy();
  });

  it('fxDialog: false disables the formula-bar fx button with a reason', async () => {
    sheet = await mountStubSheet({ features: { fxDialog: false } });
    expect(sheet.instance.features.fxDialog).toBeFalsy();
    const fx = sheet.host.querySelector<HTMLButtonElement>('.fc-host__formulabar-fx');
    expect(fx?.disabled).toBe(true);
    expect(fx?.dataset.disabledReason).toBe(defaultStrings.fxDialog.fxButtonUnavailable);
    expect(fx?.getAttribute('aria-description')).toBe(defaultStrings.fxDialog.fxButtonUnavailable);
  });

  it('keeps function arguments open and unrecorded when formula-bar commit is rejected', async () => {
    sheet = await mountStubSheet({ features: { fxDialog: true } });
    expect(sheet.instance.features.fxDialog).toBeTruthy();
    mutators.setSheetProtected(sheet.instance.store, 0, true);
    sheet.instance.openFunctionArguments('SUM');
    const overlay = document.querySelector<HTMLElement>('.fc-fxdialog');
    const insertBtn = document.querySelector<HTMLButtonElement>(
      '.fc-fxdialog .fc-fmtdlg__btn--primary',
    );

    insertBtn?.click();

    expect(overlay?.hidden).toBe(false);
    expect(document.activeElement).toBe(insertBtn);
    expect(getRecentFunctions(sheet.instance.store)).toEqual([]);
    expect(sheet.host.querySelector<HTMLTextAreaElement>('.fc-host__formulabar-input')?.value).toBe(
      '=SUM()',
    );
  });

  it('closes and records a function after formula-bar commit is accepted', async () => {
    sheet = await mountStubSheet({ features: { fxDialog: true } });
    expect(sheet.instance.features.fxDialog).toBeTruthy();
    sheet.instance.openFunctionArguments('SUM');
    const argInput = document.querySelector<HTMLInputElement>('.fc-fxdialog__arg-input');
    if (!argInput) throw new Error('expected SUM argument input');
    argInput.value = '1';
    argInput.dispatchEvent(new Event('input'));

    document.querySelector<HTMLButtonElement>('.fc-fxdialog .fc-fmtdlg__btn--primary')?.click();

    expect(document.querySelector<HTMLElement>('.fc-fxdialog')?.hidden).toBe(true);
    expect(getRecentFunctions(sheet.instance.store)).toEqual(['SUM']);
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=SUM(1)');
  });

  it('passes the mounted live workbook catalog to Function Arguments', async () => {
    const workbook = await WorkbookHandle.createDefault();
    expect(workbook.isStub).toBe(false);
    sheet = await mountStubSheet({ workbook, features: { fxDialog: true } });

    sheet.instance.openFunctionArguments('ACOS');

    expect(document.querySelector<HTMLElement>('[data-fx-name="ACOS"]')).toBeTruthy();
    expect(document.querySelector<HTMLElement>('.fc-fxdialog__picker')?.hidden).toBe(true);
    expect(document.querySelector<HTMLElement>('.fc-fxdialog__args')?.hidden).toBe(false);
    expect(document.querySelectorAll<HTMLInputElement>('.fc-fxdialog__arg-input')).toHaveLength(1);
  });

  it('viewToolbar: false omits the .fc-viewbar chrome', async () => {
    sheet = await mountStubSheet({ features: { viewToolbar: false } });
    expect(sheet.host.querySelector('.fc-viewbar')).toBeNull();
    expect(sheet.instance.features.viewToolbar).toBeFalsy();
  });

  it('sheetTabs: false omits the sheet-tab chrome', async () => {
    sheet = await mountStubSheet({ features: { sheetTabs: false } });
    expect(sheet.host.querySelector('.fc-host__sheetbar-tabs')).toBeNull();
  });

  it('statusBar: false omits the statusbar chrome', async () => {
    sheet = await mountStubSheet({ features: { statusBar: false } });
    expect(sheet.host.querySelector('.fc-host__statusbar')).toBeNull();
  });

  it('formulaBar: false omits the formula bar chrome', async () => {
    sheet = await mountStubSheet({ features: { formulaBar: false } });
    expect(sheet.host.querySelector('.fc-host__formulabar')).toBeNull();
  });

  it('conditional + iterative + namedRanges: false drops only those features', async () => {
    sheet = await mountStubSheet({
      features: { conditional: false, iterative: false, namedRanges: false },
    });
    const f = sheet.instance.features;
    expect(f.conditional).toBeFalsy();
    expect(f.iterative).toBeFalsy();
    expect(f.namedRanges).toBeFalsy();
    // Other defaults still on.
    expect(f.findReplace).toBeTruthy();
    expect(f.statusBar).toBeTruthy();
  });

  it('default mount enables charts + pivot dialog + workbook objects', async () => {
    sheet = await mountStubSheet();
    const f = sheet.instance.features;
    expect(f.charts).toBeTruthy();
    expect(f.pivotTableDialog).toBeTruthy();
    expect(f.workbookObjects).toBeTruthy();
  });
});
