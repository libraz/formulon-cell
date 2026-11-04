import { afterEach, describe, expect, it } from 'vitest';
import { Spreadsheet } from '../../../../src/mount.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { stubHelpers } from './fixtures.js';

describe('Accessibility checker routing', () => {
  let sheet: MountedStubSheet | undefined;
  afterEach(() => {
    sheet?.dispose();
    sheet = undefined;
    document.body.innerHTML = '';
  });

  it('opens the docked checker pane and shows the status on the Mac host', async () => {
    sheet = await mountStubSheet({ locale: 'ja', ui: { platform: 'mac' } });
    const { instance } = sheet;
    const status = instance.host.querySelector<HTMLButtonElement>('.fc-host__statusbar-a11y');
    expect(status?.textContent).toBe('アクセシビリティ: 問題ありません');

    const host = document.createElement('div');
    document.body.appendChild(host);
    const tb = Spreadsheet.mountToolbar(host, instance, { helpers: stubHelpers() });
    expect(tb.applyCommand('mac.review.accessibility')).toBe(true);
    const pane = instance.host.querySelector<HTMLElement>('.fc-host__taskpane-dock .fc-a11ypane');
    expect(pane?.hidden).toBe(false);
    expect(document.querySelector('.fc-tb__dlg')).toBeNull();
    tb.dispose();

    pane?.querySelector<HTMLButtonElement>('.fc-a11ypane__close')?.click();
    expect(pane?.hidden).toBe(true);
    status?.click();
    expect(pane?.hidden).toBe(false);
    expect(instance.openAccessibilityChecker()).toBe(true);
  });

  it('keeps the report dialog and no status button on other platforms', async () => {
    sheet = await mountStubSheet({ locale: 'en' });
    const { instance } = sheet;
    expect(instance.host.querySelector('.fc-host__statusbar-a11y')).toBeNull();
    expect(instance.openAccessibilityChecker()).toBe(false);
  });
});
