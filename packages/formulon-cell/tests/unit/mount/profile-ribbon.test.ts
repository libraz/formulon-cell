import { afterEach, describe, expect, it } from 'vitest';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

let sheet: MountedStubSheet;
afterEach(() => {
  sheet.dispose();
  document.body.replaceChildren();
});

const ribbonShown = (): boolean =>
  (sheet.instance.host.querySelector('.fc-host__ribbon')?.childElementCount ?? 0) > 0;

describe('a ui profile', () => {
  it('mounts no ribbon by itself, at mount or after setUi', async () => {
    sheet = await mountStubSheet({ locale: 'en', ui: { profile: 'full' } });
    expect(ribbonShown()).toBe(false);
    sheet.instance.setUi({ profile: 'standard' });
    expect(ribbonShown()).toBe(false);
  });
});
