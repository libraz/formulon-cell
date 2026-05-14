import { afterEach, describe, expect, it } from 'vitest';

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
