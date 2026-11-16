import { afterEach, describe, expect, it } from 'vitest';
import type { SpreadsheetInstance } from '../../../../src/mount/types.js';
import { Spreadsheet } from '../../../../src/mount.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { stubHelpers } from './fixtures.js';

describe('toolbar mounted before its instance exists', () => {
  let sheet: MountedStubSheet | undefined;
  afterEach(() => {
    sheet?.dispose();
    sheet = undefined;
    document.body.innerHTML = '';
  });

  it('builds the default hooks and dropdown delegation once the instance arrives', async () => {
    sheet = await mountStubSheet({ locale: 'ja', ui: { platform: 'mac' } });
    const { instance } = sheet;
    let late: SpreadsheetInstance | null = null;
    const host = document.createElement('div');
    document.body.appendChild(host);
    const tb = Spreadsheet.mountToolbar(host, () => late, {
      helpers: stubHelpers(),
      dynamicDropdowns: true,
    });
    late = instance;
    tb.rerender();

    expect(tb.applyCommand('mac.review.accessibility')).toBe(true);
    const pane = instance.host.querySelector<HTMLElement>('.fc-host__taskpane-dock .fc-a11ypane');
    expect(pane?.hidden).toBe(false);
    tb.dispose();
  });
});
