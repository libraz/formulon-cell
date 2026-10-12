import { expect, it, vi } from 'vitest';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { createBackstageTitle } from '../../../../src/toolbar/ribbon/backstage-title.js';
import { mountStubSheet } from '../../../test-utils/mount.js';

it('searches German displayed values from the title search field', async () => {
  const workbook = await WorkbookHandle.createDefault({ locale: 'de' });
  workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 0);
  workbook.setNumber({ sheet: 0, row: 1, col: 0 }, 1.234);
  const sheet = await mountStubSheet({ workbook, locale: 'de' });
  const input = document.createElement('input');
  input.value = '1,234';
  const noop = (): void => {};
  const documentListeners = vi.spyOn(document, 'addEventListener');
  try {
    const api = createBackstageTitle({
      getInst: () => sheet.instance,
      shellText: {
        save: '',
        saveAs: '',
        autosave: '',
        autosaveOn: '',
        autosaveOff: '',
        comments: '',
        share: '',
        shareReady: '',
        findNoMatches: '{query}',
      },
      ribbonRoot: null,
      titleSearchInput: input,
      autosaveSwitch: null,
      statusMetric: null,
      focusSheet: noop,
      triggerSave: noop,
      triggerSaveAs: noop,
      renderRibbon: noop,
      refreshAutosave: noop,
      projectFormatToolbar: noop,
      showRibbonReport: noop,
      setRibbonDisplayMenuOpen: noop,
      getActiveRibbonTab: () => 'home',
      setActiveRibbonTab: noop,
      getRibbonCollapsed: () => false,
      setRibbonCollapsed: noop,
      getBackstageOpen: () => false,
      setBackstageOpen: noop,
      getBackstageReturnTab: () => 'home',
      setBackstageReturnTab: noop,
      getAutosaveEnabled: () => false,
      setAutosaveEnabled: noop,
    });
    api.runTitleSearch();
    expect(workbook.isStub).toBe(false);
    expect(sheet.instance.store.getState().selection.active).toEqual({ sheet: 0, row: 1, col: 0 });
  } finally {
    for (const [type, listener, options] of documentListeners.mock.calls) {
      document.removeEventListener(type, listener, options);
    }
    documentListeners.mockRestore();
    sheet.dispose();
    workbook.dispose();
  }
});
