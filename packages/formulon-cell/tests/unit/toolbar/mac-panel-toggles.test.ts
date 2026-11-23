import { afterEach, describe, expect, it, vi } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { Spreadsheet } from '../../../src/mount.js';
import { mutators } from '../../../src/store/store.js';
import { projectMacRibbonState } from '../../../src/toolbar/ribbon/mac/state-projection.js';
import { projectRibbonActiveState } from '../../../src/toolbar/ribbon-active-state.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

describe('Mac ribbon panel toggles', () => {
  let sheet: MountedStubSheet | undefined;

  afterEach(() => {
    sheet?.dispose();
    sheet = undefined;
  });

  const ribbonButton = (id: string) => {
    const host = document.createElement('div');
    const button = document.createElement('button');
    button.dataset.ribbonCommand = id;
    host.appendChild(button);
    return { host, button };
  };

  it('shows the Watch Window button pressed while the watch pane is open', async () => {
    sheet = await mountStubSheet({
      workbook: await WorkbookHandle.createDefault({ preferStub: true }),
      locale: 'ja',
    });
    const { host, button } = ribbonButton('mac.formulas.watch');
    projectMacRibbonState(host, sheet.instance);
    expect(button.getAttribute('aria-pressed')).toBe('false');
    mutators.setWatchPanelOpen(sheet.instance.store, true);
    projectMacRibbonState(host, sheet.instance);
    expect(button.getAttribute('aria-pressed')).toBe('true');
    expect(button.classList.contains('fc-tb__rb--active')).toBe(true);
    mutators.setWatchPanelOpen(sheet.instance.store, false);
    projectMacRibbonState(host, sheet.instance);
    expect(button.getAttribute('aria-pressed')).toBe('false');
  });

  it('shows 選択ウインドウ pressed while the Workbook Objects panel is open', async () => {
    sheet = await mountStubSheet({
      workbook: await WorkbookHandle.createDefault({ preferStub: true }),
      locale: 'ja',
    });
    const instance = sheet.instance;
    const { host, button } = ribbonButton('objectSelectionPane');
    const pressed = () => {
      projectMacRibbonState(host, instance);
      return button.getAttribute('aria-pressed');
    };
    const panel = () => instance.host.querySelector<HTMLElement>('.fc-objects');
    expect(pressed()).toBe('false');

    instance.openWorkbookObjects();
    expect(instance.store.getState().ui.workbookObjectsOpen).toBe(true);
    expect(pressed()).toBe('true');
    panel()?.querySelector<HTMLButtonElement>('.fc-objects__close')?.click();
    expect(instance.store.getState().ui.workbookObjectsOpen).toBe(false);
    expect(pressed()).toBe('false');

    instance.openWorkbookObjects();
    expect(pressed()).toBe('true');
    panel()?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(instance.store.getState().ui.workbookObjectsOpen).toBe(false);
    expect(pressed()).toBe('false');
  });

  it('closes the Workbook Objects panel on a second 選択ウインドウ click', async () => {
    sheet = await mountStubSheet({
      workbook: await WorkbookHandle.createDefault({ preferStub: true }),
      locale: 'ja',
    });
    const { instance } = sheet;
    const host = document.createElement('div');
    document.body.appendChild(host);
    const tb = Spreadsheet.mountToolbar(host, instance, {});
    try {
      expect(tb.applyCommand('objectSelectionPane')).toBe(true);
      expect(instance.store.getState().ui.workbookObjectsOpen).toBe(true);
      expect(tb.applyCommand('objectSelectionPane')).toBe(true);
      expect(instance.store.getState().ui.workbookObjectsOpen).toBe(false);
      expect(instance.host.querySelector<HTMLElement>('.fc-objects')?.hidden).toBe(true);
    } finally {
      tb.dispose();
      host.remove();
    }
  });
  it('shows Field List pressed while the PivotTable field list is open', async () => {
    sheet = await mountStubSheet({
      workbook: await WorkbookHandle.createDefault({ preferStub: true }),
      locale: 'ja',
    });
    const { instance } = sheet;
    const { host, button } = ribbonButton('pivotFieldList');
    const text = { background: 'Background', sheetBackgroundClear: 'Delete Background' };
    projectMacRibbonState(host, instance);
    expect(button.getAttribute('aria-pressed')).toBe('false');
    mutators.setPivotFieldListOpen(instance.store, true);
    projectMacRibbonState(host, instance);
    expect(button.getAttribute('aria-pressed')).toBe('true');
    mutators.setPivotFieldListOpen(instance.store, false);
    projectRibbonActiveState(host, instance, text);
    expect(button.getAttribute('aria-pressed')).toBe('false');
    mutators.setPivotFieldListOpen(instance.store, true);
    projectRibbonActiveState(host, instance, text);
    expect(button.getAttribute('aria-pressed')).toBe('true');
  });

  it('hides the open PivotTable field list on a second Field List click', async () => {
    sheet = await mountStubSheet({
      workbook: await WorkbookHandle.createDefault({ preferStub: true }),
      locale: 'ja',
    });
    const { instance } = sheet;
    const openList = vi.spyOn(instance, 'openActivePivotFieldList').mockReturnValue(true);
    const hideList = vi.spyOn(instance, 'hidePivotFieldList');
    const host = document.createElement('div');
    document.body.appendChild(host);
    const tb = Spreadsheet.mountToolbar(host, instance, {});
    try {
      for (const id of ['pivotFieldList', 'pivotFieldListView']) {
        openList.mockClear();
        hideList.mockClear();
        mutators.setPivotFieldListOpen(instance.store, false);
        expect(tb.applyCommand(id)).toBe(true);
        expect(openList).toHaveBeenCalledTimes(1);
        expect(hideList).not.toHaveBeenCalled();
        mutators.setPivotFieldListOpen(instance.store, true);
        expect(tb.applyCommand(id)).toBe(true);
        expect(hideList).toHaveBeenCalledTimes(1);
        expect(openList).toHaveBeenCalledTimes(1);
      }
    } finally {
      tb.dispose();
      host.remove();
    }
  });
});
