import { SparklineType } from '@libraz/formulon';
import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  defaultSparklineGroupSettings,
  groupOfCell,
  setSparklines,
} from '../../../../src/commands/sparkline.js';
import { SPARKLINE_STYLE_PRESETS } from '../../../../src/commands/sparkline-presets.js';
import { Spreadsheet } from '../../../../src/mount.js';
import { mutators } from '../../../../src/store/store.js';
import measured from '../../../fixtures/excel-sparkline-style-presets.json' with { type: 'json' };
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { scaledTimeout } from '../../../timeouts.js';
import { stubHelpers } from './fixtures.js';

vi.setConfig({ testTimeout: scaledTimeout(20_000) });

const at = (row: number) => ({ sheet: 0, row, col: 5 });

describe('Sparkline Style strip', () => {
  let sheet: MountedStubSheet;
  let host: HTMLElement;
  let tb: ReturnType<typeof Spreadsheet.mountToolbar>;

  afterEach(() => {
    tb.dispose();
    sheet.dispose();
    host.remove();
  });

  const mount = async (locale: 'en' | 'ja', platform?: 'mac'): Promise<void> => {
    sheet = await mountStubSheet({ locale });
    host = document.createElement('div');
    document.body.appendChild(host);
    const store = sheet.instance.store;
    setSparklines(
      store,
      [0, 1].map((row) => ({ addr: at(row), source: `A${row + 1}:C${row + 1}` })),
      defaultSparklineGroupSettings(SparklineType.Line),
      sheet.instance.history,
    );
    mutators.setActive(store, at(0));
    tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
      ...(platform ? { platform } : {}),
    });
    tb.setActiveTab('sparklineDesign');
  };
  const chips = () =>
    Array.from(host.querySelectorAll<HTMLButtonElement>('[data-sparkline-quick-style]'));
  const ids = () => chips().map((chip) => chip.dataset.sparklineQuickStyle);

  it('shows a row of seven presets without page arrows on the Mac, then the chevron', async () => {
    await mount('ja', 'mac');
    expect(ids()).toEqual([
      'accentDarker50-1',
      'accentDarker50-2',
      'accentDarker50-3',
      'accentDarker50-4',
      'accentDarker50-5',
      'accentDarker50-6',
      'accentDarker25-1',
    ]);
    const gallery = host.querySelector('.fc-tb__sparkline-quick-gallery');
    expect(gallery?.querySelector('.fc-tb__sparkline-quick-prev')).toBeNull();
    const chevron = host.querySelector('[data-ribbon-command="sparklineStyleGallery"]');
    expect(chevron?.parentElement).toBe(gallery);
    expect(chevron?.getAttribute('aria-label')).toBe('他のアイテムを表示');
  });

  it('lists every preset in the strip but the seven the desktop leaves to the full gallery', async () => {
    await mount('en');
    const next = host.querySelector<HTMLButtonElement>('.fc-tb__sparkline-quick-next');
    const seen: (string | undefined)[] = [...ids()];
    for (let turn = 0; turn < 10 && next && !next.disabled; turn++) {
      next.click();
      seen.push(...ids());
    }
    const family = (name: string, from: number, to: number) =>
      Array.from({ length: to - from + 1 }, (_, i) => `${name}-${from + i}`);
    expect(seen).toEqual([
      ...family('accentDarker50', 1, 6),
      ...family('accentDarker25', 1, 6),
      ...family('accent', 1, 6),
      ...family('accentLighter40', 1, 3),
      ...family('dark', 5, 6),
      ...family('colorful', 1, 6),
    ]);
  });

  it('names every preset as the desktop gallery does in ja', async () => {
    await mount('ja', 'mac');
    const labels = Array.from(
      host.querySelectorAll<HTMLButtonElement>('[data-sparkline-style]'),
      (swatch) => swatch.getAttribute('aria-label'),
    );
    expect(labels).toEqual(measured.map((row) => row.label));
  });

  it('applies a chip and turns to the row of the current preset', async () => {
    await mount('en');
    expect(host.querySelector('.fc-tb__sparkline-quick-prev')).not.toBeNull();
    chips()[2]?.click();
    const settings = groupOfCell(sheet.instance.store.getState(), at(1))?.settings;
    expect(settings?.colorSeries).toEqual(SPARKLINE_STYLE_PRESETS[2]?.series);
    expect(chips()[2]?.getAttribute('aria-pressed')).toBe('true');
    // Another group's preset shows its own row once its cell is active.
    setSparklines(
      sheet.instance.store,
      [{ addr: at(5), source: 'A6:C6' }],
      {
        ...defaultSparklineGroupSettings(SparklineType.Line),
        colorSeries: SPARKLINE_STYLE_PRESETS[30]?.series ?? null,
        colorNegative: SPARKLINE_STYLE_PRESETS[30]?.negative ?? null,
        colorAxis: SPARKLINE_STYLE_PRESETS[30]?.axis ?? null,
        colorMarkers: SPARKLINE_STYLE_PRESETS[30]?.markers ?? null,
        colorFirst: SPARKLINE_STYLE_PRESETS[30]?.first ?? null,
        colorLast: SPARKLINE_STYLE_PRESETS[30]?.last ?? null,
        colorHigh: SPARKLINE_STYLE_PRESETS[30]?.high ?? null,
        colorLow: SPARKLINE_STYLE_PRESETS[30]?.low ?? null,
      },
      sheet.instance.history,
    );
    mutators.setActive(sheet.instance.store, at(5));
    expect(ids()).toContain('colorful-1');
    expect(
      host.querySelector('[data-sparkline-quick-style="colorful-1"]')?.getAttribute('aria-pressed'),
    ).toBe('true');
  });
});
