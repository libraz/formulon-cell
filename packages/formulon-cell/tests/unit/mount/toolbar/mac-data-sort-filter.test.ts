import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { Spreadsheet } from '../../../../src/mount.js';
import { mutators } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { scaledTimeout } from '../../../timeouts.js';
import { stubHelpers } from './fixtures.js';

vi.setConfig({ testTimeout: scaledTimeout(20_000) });

// Excel for Mac lays Sort & Filter out as a sort-icon pair, two full-height tiles and a column of
// three small rows.
describe('Mac Data tab Sort & Filter group', () => {
  let sheet: MountedStubSheet;
  let host: HTMLElement;

  beforeEach(async () => {
    sheet = await mountStubSheet({ locale: 'en' });
    sheet.instance.host.dataset.fcPlatform = 'mac';
    host = document.createElement('div');
    document.body.appendChild(host);
  });
  afterEach(() => {
    sheet.dispose();
    host.remove();
  });

  const button = (id: string) =>
    host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${id}"]`);

  it('renders Sort and Filter as tiles and Clear, Reapply, Advanced as small rows', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      platform: 'mac',
      helpers: stubHelpers(),
    });
    const group = button('mac.data.sortCustom')?.closest('.fc-tb__ribbon-group');
    expect(group?.classList.contains('fc-tb__ribbon-group--mixed')).toBe(true);
    for (const id of ['mac.data.sortCustom', 'mac.data.filter']) {
      const el = button(id);
      expect(el?.classList.contains('fc-tb__rb--wide')).toBe(true);
      expect(el?.classList.contains('fc-tb__rb--stacked')).toBe(false);
      expect(el?.closest('.fc-tb__ribbon-group')).toBe(group);
    }
    for (const id of ['mac.data.clear', 'mac.data.reapply', 'mac.data.advancedFilter']) {
      const el = button(id);
      expect(el?.classList.contains('fc-tb__rb--stacked')).toBe(true);
      expect(el?.closest('.fc-tb__ribbon-group')).toBe(group);
    }
    for (const id of ['mac.data.sortAsc', 'mac.data.sortDesc']) {
      const el = button(id);
      expect(el?.classList.contains('fc-tb__rb--mac-icon-only')).toBe(true);
      expect(el?.closest('.fc-tb__ribbon-group')).toBe(group);
    }
    tb.dispose();
  });

  it('shows Filter pressed while a filter is applied', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      platform: 'mac',
      helpers: stubHelpers(),
    });
    expect(button('mac.data.filter')?.getAttribute('aria-pressed')).toBe('false');
    mutators.setFilterRanges(sheet.instance.store, [{ sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 }]);
    expect(button('mac.data.filter')?.getAttribute('aria-pressed')).toBe('true');
    tb.dispose();
  });
});
