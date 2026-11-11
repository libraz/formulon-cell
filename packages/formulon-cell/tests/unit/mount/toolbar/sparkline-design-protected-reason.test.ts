import { SparklineType } from '@libraz/formulon';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { setProtectedSheet } from '../../../../src/commands/protection.js';
import {
  defaultSparklineGroupSettings,
  setSparklines,
} from '../../../../src/commands/sparkline.js';
import { Spreadsheet } from '../../../../src/mount.js';
import { mutators } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { stubHelpers } from './fixtures.js';

vi.setConfig({ testTimeout: 20_000 });

describe('Sparkline tab buttons greyed by a protected sheet', () => {
  let sheet: MountedStubSheet;
  let host: HTMLElement;
  let tb: ReturnType<typeof Spreadsheet.mountToolbar>;

  beforeEach(async () => {
    sheet = await mountStubSheet({ locale: 'en' });
    host = document.createElement('div');
    document.body.appendChild(host);
    const store = sheet.instance.store;
    setSparklines(
      store,
      [{ addr: { sheet: 0, row: 0, col: 5 }, source: 'A1:C1' }],
      defaultSparklineGroupSettings(SparklineType.Line),
      sheet.instance.history,
    );
    mutators.setActive(store, { sheet: 0, row: 0, col: 5 });
    tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    tb.setActiveTab('sparklineDesign');
  });

  afterEach(() => {
    tb.dispose();
    sheet.dispose();
    host.remove();
  });

  const command = (id: string): HTMLButtonElement => {
    const found = host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${id}"]`);
    if (!found) throw new Error(`missing ${id}`);
    return found;
  };

  it('keeps the command title and adds the reason, then restores the bare title', () => {
    const own = command('sparklineShowHigh').title;
    expect(own).not.toBe('');
    setProtectedSheet(sheet.instance.store, 0, true);
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 5 });
    const button = command('sparklineShowHigh');
    expect(button.disabled).toBe(true);
    expect(button.title).toBe(
      `${own}\n${sheet.instance.i18n.strings.mac.sparkline.protectedLocation}`,
    );
    setProtectedSheet(sheet.instance.store, 0, false);
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 5 });
    expect(command('sparklineShowHigh').disabled).toBe(false);
    expect(command('sparklineShowHigh').title).toBe(own);
  });

  it('greys the Style strip with the reason and gives it back once unprotected', () => {
    const reason = sheet.instance.i18n.strings.mac.sparkline.protectedLocation;
    const chip = () => host.querySelector<HTMLButtonElement>('[data-sparkline-quick-style]');
    expect(chip()?.disabled).toBe(false);
    setProtectedSheet(sheet.instance.store, 0, true);
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 5 });
    expect(chip()?.disabled).toBe(true);
    expect(chip()?.title).toBe(`${chip()?.getAttribute('aria-label')}\n${reason}`);
    setProtectedSheet(sheet.instance.store, 0, false);
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 5 });
    expect(chip()?.disabled).toBe(false);
  });

  it('keeps the name of the disabled Plot Right to Left entry', () => {
    command('sparklineAxis').click();
    const rtl = host.querySelector<HTMLButtonElement>('[data-sparkline-axis="rtl"]');
    expect(rtl?.disabled).toBe(true);
    expect(rtl?.title).not.toBe('');
  });
});
