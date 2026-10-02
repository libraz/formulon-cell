import { afterEach, describe, expect, it } from 'vitest';

import { getRecentFunctions } from '../../src/commands/function-history.js';
import { WorkbookHandle } from '../../src/engine/workbook-handle.js';
import { mutators } from '../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../test-utils/index.js';

describe('integration: Function Arguments live engine catalog', () => {
  let sheet: MountedStubSheet | undefined;

  afterEach(() => {
    sheet?.dispose();
    sheet = undefined;
  });

  it('opens engine-only ACOS, commits through the mounted formula bar, and records Recent', async () => {
    const workbook = await WorkbookHandle.createDefault();
    expect(workbook.isStub).toBe(false);
    const names = workbook.functionNames();
    expect(names).toContain('ACOS');
    sheet = await mountStubSheet({ workbook, features: { fxDialog: true } });
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });

    sheet.instance.openFunctionArguments('ACOS');
    expect(document.querySelector<HTMLElement>('.fc-fxdialog__args')?.hidden).toBe(false);
    const input = document.querySelector<HTMLInputElement>('.fc-fxdialog__arg-input');
    expect(input).toBeTruthy();
    if (!input) return;
    input.value = '0';
    input.dispatchEvent(new Event('input'));
    document.querySelector<HTMLButtonElement>('.fc-fxdialog .fc-fmtdlg__btn--primary')?.click();

    expect(workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=ACOS(0)');
    workbook.recalc();
    const value = workbook.getValue({ sheet: 0, row: 0, col: 0 });
    expect(value.kind).toBe('number');
    if (value.kind === 'number') expect(value.value).toBeCloseTo(Math.PI / 2, 12);
    expect(getRecentFunctions(sheet.instance.store, new Set(names ?? []))).toContain('ACOS');
  });

  it('re-reads the supplied real workbook after a mounted workbook swap', async () => {
    const first = await WorkbookHandle.createDefault();
    const second = await WorkbookHandle.createDefault();
    expect(first.isStub).toBe(false);
    expect(second.isStub).toBe(false);
    first.setFunctionMetadataProvider({
      ACOS: { aliases: { 'en-US': 'First cosine' } },
    });
    second.setFunctionMetadataProvider({
      ACOS: { aliases: { 'en-US': 'Second cosine' } },
    });
    expect(first.functionMetadata('ACOS', 0)?.localizedName).toBe('First cosine');
    expect(second.functionMetadata('ACOS', 0)?.localizedName).toBe('Second cosine');

    sheet = await mountStubSheet({ workbook: first, features: { fxDialog: true }, locale: 'en' });
    sheet.instance.openFunctionArguments('ACOS');
    expect(
      document.querySelector<HTMLElement>('[data-fx-name="ACOS"] .fc-fxdialog__item-name')
        ?.textContent,
    ).toBe('First cosine');
    document.querySelector<HTMLButtonElement>('.fc-fxdialog .fc-fmtdlg__btn--secondary')?.click();

    await sheet.instance.setWorkbook(second);
    expect(sheet.instance.workbook).toBe(second);
    expect(sheet.instance.workbook.functionMetadata('ACOS', 0)?.localizedName).toBe(
      'Second cosine',
    );
    sheet.instance.openFunctionArguments('ACOS');
    expect(
      document.querySelector<HTMLElement>('[data-fx-name="ACOS"] .fc-fxdialog__item-name')
        ?.textContent,
    ).toBe('Second cosine');
  });
});
