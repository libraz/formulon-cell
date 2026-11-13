import { SparklineType } from '@libraz/formulon';
import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  defaultSparklineGroupSettings,
  groupOfCell,
  rgbColorSpec,
  setSparklines,
} from '../../../../src/commands/sparkline.js';
import { Spreadsheet } from '../../../../src/mount.js';
import { sparklineColors } from '../../../../src/render/sparkline.js';
import { mutators } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { scaledTimeout } from '../../../timeouts.js';
import { stubHelpers } from './fixtures.js';

vi.setConfig({ testTimeout: scaledTimeout(20_000) });

const at = (row: number) => ({ sheet: 0, row, col: 5 });

describe('Sparkline Color menu', () => {
  let sheet: MountedStubSheet;
  let host: HTMLElement;
  let tb: ReturnType<typeof Spreadsheet.mountToolbar>;

  afterEach(() => {
    tb.dispose();
    sheet.dispose();
    host.remove();
    document.body.replaceChildren();
  });

  const mount = async (): Promise<void> => {
    sheet = await mountStubSheet({ locale: 'ja' });
    host = document.createElement('div');
    document.body.appendChild(host);
    const store = sheet.instance.store;
    setSparklines(
      store,
      [{ addr: at(0), source: 'A1:C1' }],
      defaultSparklineGroupSettings(SparklineType.Line),
      sheet.instance.history,
    );
    mutators.setActive(store, at(0));
    tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
      platform: 'mac',
    });
    tb.setActiveTab('sparklineDesign');
  };
  const settings = () => groupOfCell(sheet.instance.store.getState(), at(0))?.settings;
  const pick = (selector: string): void => {
    const target = host.querySelector<HTMLElement>(selector);
    expect(target, selector).toBeTruthy();
    const click = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(click, 'target', { value: target });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(click)).toBe(true);
  };
  const open = (id: string): HTMLElement => {
    host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${id}"]`)?.click();
    const menu = host.querySelector<HTMLElement>(
      id === 'sparklineColor' ? '#menu-sparkline-color' : '#menu-sparkline-marker-color',
    );
    if (!menu) throw new Error(`no menu for ${id}`);
    return menu;
  };
  /** Pick a colour in the system picker the More Colors row opened. */
  const chooseNative = (css: string): void => {
    const picker = document.body.querySelector<HTMLInputElement>(':scope > input[type="color"]');
    if (!picker) throw new Error('no colour picker');
    picker.value = css;
    picker.dispatchEvent(new Event('change'));
  };

  it('lists high contrast, the theme and standard colours, More Colors and Weight', async () => {
    await mount();
    const menu = open('sparklineColor');
    const contrast = menu.querySelector<HTMLElement>('.fc-colorpalette__contrast');
    expect(contrast?.textContent).toBe('ハイ コントラストのみ');
    expect(contrast?.querySelector<HTMLInputElement>('input')?.disabled).toBe(true);
    expect(menu.textContent).toContain('テーマの色');
    expect(menu.textContent).toContain('標準の色');
    expect(menu.querySelector('.fc-colorpalette__action--more')?.textContent).toBe('その他の色...');
    const weights = Array.from(menu.querySelectorAll<HTMLElement>('[data-sparkline-weight]')).map(
      (el) => el.dataset.sparklineWeight,
    );
    expect(weights).toEqual(['0.25', '0.5', '0.75', '1', '1.5', '2.25', '3', '4.5', '6', 'custom']);
    expect(menu.querySelector('[data-sparkline-weight="custom"]')?.textContent).toBe(
      'ユーザー設定の太さ...',
    );
  });

  it('colours the series and a marker from More Colors', async () => {
    await mount();
    open('sparklineColor');
    pick('[data-sparkline-color="more"]');
    chooseNative('#123456');
    expect(settings()?.colorSeries).toEqual(rgbColorSpec('#123456'));
    open('sparklineMarkerColor');
    pick('[data-sparkline-marker-color="high:more"]');
    chooseNative('#abcdef');
    expect(settings()?.colorHigh).toEqual(rgbColorSpec('#abcdef'));
  });

  it('lists the six points as colour submenus, Markers closing the first block', async () => {
    await mount();
    const menu = open('sparklineMarkerColor');
    const rows = Array.from(menu.children)
      .filter((el) => !el.classList.contains('fc-tb__submenu'))
      .map((el) =>
        el.getAttribute('role') === 'separator' ? '|' : (el.textContent ?? '').replace('›', ''),
      );
    expect(rows).toEqual([
      '負のポイント',
      'マーカー',
      '|',
      '頂点 (山)',
      '頂点 (谷)',
      '始点',
      '終点',
    ]);
    for (const part of ['negative', 'markers', 'high', 'low', 'first', 'last']) {
      const panel = menu.querySelector<HTMLElement>(`#menu-sparkline-marker-color-${part}`);
      expect(
        panel?.querySelector('.fc-colorpalette__contrast input')?.hasAttribute('disabled'),
      ).toBe(true);
      expect(panel?.querySelector('.fc-colorpalette__action--automatic')).toBeNull();
      expect(panel?.querySelector('.fc-colorpalette__action--no-fill')).toBeNull();
      expect(
        panel?.querySelectorAll('.fc-colorpalette__grid--theme .fc-colorpalette__swatch'),
      ).toHaveLength(60);
      expect(
        panel?.querySelectorAll('.fc-colorpalette__grid--standard .fc-colorpalette__swatch'),
      ).toHaveLength(10);
      expect(panel?.querySelector('.fc-colorpalette__action--more')).not.toBeNull();
    }
  });

  it('shows each point colour on its entry and frames it in the palette', async () => {
    await mount();
    open('sparklineColor');
    pick('[data-sparkline-color="#ff0000"]');
    let menu = open('sparklineColor');
    expect(
      menu
        .querySelector('.fc-colorpalette__swatch[aria-pressed="true"]')
        ?.getAttribute('data-color'),
    ).toBe('#ff0000');
    open('sparklineMarkerColor');
    pick('[data-sparkline-marker-color="high:#00b050"]');
    menu = open('sparklineMarkerColor');
    const bar = (part: string) =>
      menu
        .querySelector(`[data-format-submenu="${part}"] .fc-tb__pen-color-bar`)
        ?.getAttribute('stroke');
    const framed = (part: string) =>
      menu
        .querySelector(
          `#menu-sparkline-marker-color-${part} .fc-colorpalette__swatch[aria-pressed="true"]`,
        )
        ?.getAttribute('data-color');
    expect(bar('high')).toBe('#00b050');
    expect(framed('high')).toBe('#00b050');
    const low = sparklineColors(
      settings() ?? defaultSparklineGroupSettings(SparklineType.Line),
      null,
    ).low;
    expect(bar('low')).toBe(low ?? '#ff0000');
  });

  it.each([0.2, 1585])('accepts a custom weight of %s pt', async (weight) => {
    await mount();
    open('sparklineColor');
    pick('[data-sparkline-weight="custom"]');
    await vi.waitFor(() => expect(document.querySelector('.fc-tb__dlg')).toBeTruthy());
    const dlg = document.querySelector('.fc-tb__dlg') as HTMLElement;
    expect(dlg.querySelector('.fc-fmtdlg__header')?.textContent).toBe('スパークラインの太さ');
    const input = dlg.querySelector<HTMLInputElement>('input');
    if (!input) throw new Error('no weight field');
    expect(input.value).toBe('0.75');
    input.value = String(weight);
    dlg.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await vi.waitFor(() => expect(settings()?.lineWidthPt).toBe(weight));
  });
});
