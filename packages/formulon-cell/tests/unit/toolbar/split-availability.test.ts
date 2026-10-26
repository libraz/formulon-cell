import { afterEach, describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { projectMacRibbonState } from '../../../src/toolbar/ribbon/mac/state-projection.js';
import { projectRibbonActiveState } from '../../../src/toolbar/ribbon-active-state.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

describe('Split availability', () => {
  let sheet: MountedStubSheet | undefined;

  afterEach(() => {
    sheet?.dispose();
    sheet = undefined;
  });

  const ribbon = () => {
    const host = document.createElement('div');
    host.innerHTML =
      '<button data-ribbon-command="split"></button><button data-ribbon-command="mac.view.split"></button>';
    const button = (id: string) =>
      host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${id}"]`) as HTMLButtonElement;
    return { host, windows: button('split'), mac: button('mac.view.split') };
  };

  it('is offered on an engine that saves split panes', async () => {
    sheet = await mountStubSheet({ workbook: await WorkbookHandle.createDefault(), locale: 'en' });
    const { host, windows, mac } = ribbon();
    projectRibbonActiveState(host, sheet.instance);
    projectMacRibbonState(host, sheet.instance);
    expect(windows.getAttribute('aria-disabled')).not.toBe('true');
    expect(mac.getAttribute('aria-disabled')).not.toBe('true');
  });

  it('is disabled up front, with the reason, where a split cannot be saved', async () => {
    sheet = await mountStubSheet({ locale: 'en' });
    const { host, windows, mac } = ribbon();
    projectRibbonActiveState(host, sheet.instance);
    projectMacRibbonState(host, sheet.instance);
    const reason = sheet.instance.i18n.strings.ribbonMenu.splitNotSaved;
    expect(reason).toBeTruthy();
    for (const button of [windows, mac]) {
      expect(button.getAttribute('aria-disabled')).toBe('true');
      expect(button.dataset.ribbonDisabledReason).toBe(reason);
    }
  });
});
