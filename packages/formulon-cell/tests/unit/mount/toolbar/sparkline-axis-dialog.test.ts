import { SparklineAxisType, SparklineType } from '@libraz/formulon';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { setProtectedSheet } from '../../../../src/commands/protection.js';
import {
  defaultSparklineGroupSettings,
  groupOfCell,
  setSparklines,
} from '../../../../src/commands/sparkline.js';
import { Spreadsheet } from '../../../../src/mount.js';
import { mutators } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { stubHelpers } from './fixtures.js';

vi.setConfig({ testTimeout: 20_000 });

const at = (row: number) => ({ sheet: 0, row, col: 5 });

describe('Sparkline Axis dialog on the Mac', () => {
  let sheet: MountedStubSheet;
  let host: HTMLElement;
  let tb: ReturnType<typeof Spreadsheet.mountToolbar>;

  afterEach(() => {
    tb.dispose();
    sheet.dispose();
    host.remove();
    document.body.replaceChildren();
  });

  const mount = async (platform?: 'mac'): Promise<void> => {
    sheet = await mountStubSheet({ locale: 'ja' });
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
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 5, r1: 1, c1: 5 });
    tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
      ...(platform ? { platform } : {}),
    });
    tb.setActiveTab('sparklineDesign');
  };
  const settings = () => groupOfCell(sheet.instance.store.getState(), at(0))?.settings;
  const axisButton = (): HTMLButtonElement => {
    const found = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="sparklineAxisDialog"]',
    );
    if (!found) throw new Error('no Axis dialog button');
    return found;
  };
  const open = async (): Promise<HTMLElement> => {
    axisButton().click();
    await vi.waitFor(() => expect(document.querySelector('.fc-macsparkaxis')).toBeTruthy());
    return document.querySelector('.fc-macsparkaxis') as HTMLElement;
  };
  const radio = (dlg: HTMLElement, label: string, nth = 0): HTMLInputElement => {
    const found = Array.from(dlg.querySelectorAll<HTMLLabelElement>('.fc-fmtdlg__radio')).filter(
      (wrap) => wrap.querySelector('span')?.textContent === label,
    )[nth];
    const input = found?.querySelector<HTMLInputElement>('input');
    if (!input) throw new Error(`no radio ${label}`);
    return input;
  };
  const check = (dlg: HTMLElement, label: string): HTMLInputElement => {
    const found = Array.from(dlg.querySelectorAll<HTMLLabelElement>('.fc-fmtdlg__checkbox')).find(
      (wrap) => wrap.textContent === label,
    );
    const input = found?.querySelector<HTMLInputElement>('input');
    if (!input) throw new Error(`no checkbox ${label}`);
    return input;
  };
  const choose = (input: HTMLInputElement): void => {
    input.click();
    input.dispatchEvent(new Event('change', { bubbles: true }));
  };
  const ok = (dlg: HTMLElement) => dlg.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');

  it('replaces the Axis menu with a dialog button on the Mac profile only', async () => {
    await mount('mac');
    expect(axisButton().dataset.ribbonActivation).toBe('dialog');
    expect(host.querySelector('[data-ribbon-command="sparklineAxis"]')).toBeNull();
    tb.dispose();
    sheet.dispose();
    host.remove();
    await mount();
    expect(host.querySelector('[data-ribbon-command="sparklineAxisDialog"]')).toBeNull();
    expect(
      host.querySelector<HTMLElement>('[data-ribbon-command="sparklineAxis"]')?.dataset
        .ribbonActivation,
    ).toBe('dropdown');
  });

  it('lays out the two tabs with the desktop wording', async () => {
    await mount('mac');
    const dlg = await open();
    expect(dlg.getAttribute('role')).toBe('dialog');
    expect(dlg.getAttribute('aria-label')).toBe('軸');
    expect(Array.from(dlg.querySelectorAll('[role="tab"]')).map((tab) => tab.textContent)).toEqual([
      '横',
      '縦',
    ]);
    const [horizontal, vertical] = Array.from(
      dlg.querySelectorAll<HTMLElement>('[role="tabpanel"]'),
    );
    expect(horizontal?.hidden).toBe(false);
    expect(vertical?.hidden).toBe(true);
    expect(horizontal?.textContent).toBe(
      '軸の種類 (基本)日付軸スパークラインのデータ範囲を選んでください:軸の表示データを右から左にプロット',
    );
    expect(vertical?.textContent).toBe(
      [
        '最小値:',
        'スパークラインごとに自動設定',
        'すべてのスパークラインに同じ値',
        'ユーザー設定値:',
        '最大値:',
        'スパークラインごとに自動設定',
        'すべてのスパークラインに同じ値',
        'ユーザー設定値:',
      ].join(''),
    );
    const dateField = dlg.querySelector<HTMLInputElement>(
      'input[aria-label="スパークラインのデータ範囲を選んでください:"]',
    );
    expect(dateField?.disabled).toBe(true);
    choose(radio(dlg, '日付軸'));
    expect(dateField?.disabled).toBe(false);
    const rtl = check(dlg, 'データを右から左にプロット');
    expect(rtl.disabled).toBe(true);
    expect(rtl.title).toContain('このオプションはまだサポートされていません。');
  });

  it('applies the horizontal and vertical choices together as one undo step', async () => {
    await mount('mac');
    const dlg = await open();
    choose(radio(dlg, '日付軸'));
    const dateField = dlg.querySelector<HTMLInputElement>(
      'input[aria-label="スパークラインのデータ範囲を選んでください:"]',
    ) as HTMLInputElement;
    dateField.value = 'A9:C9';
    choose(check(dlg, '軸の表示'));
    dlg.querySelectorAll<HTMLButtonElement>('[role="tab"]')[1]?.click();
    choose(radio(dlg, 'ユーザー設定値:', 0));
    const values = dlg.querySelectorAll<HTMLInputElement>('.fc-macsparkaxis__value');
    expect(values[0]?.disabled).toBe(false);
    expect(values[1]?.disabled).toBe(true);
    (values[0] as HTMLInputElement).value = '-2.5';
    choose(radio(dlg, 'すべてのスパークラインに同じ値', 1));
    ok(dlg)?.click();
    await vi.waitFor(() => expect(document.querySelector('.fc-macsparkaxis')).toBeNull());
    expect(settings()).toMatchObject({
      dateAxis: true,
      dateAxisRange: { sheet: 0, rect: { firstRow: 8, firstCol: 0, lastRow: 8, lastCol: 2 } },
      displayXAxis: true,
      minAxisType: SparklineAxisType.Custom,
      manualMin: -2.5,
      maxAxisType: SparklineAxisType.Group,
    });
    sheet.instance.history.undo();
    expect(settings()).toMatchObject({
      dateAxis: false,
      displayXAxis: false,
      minAxisType: SparklineAxisType.Individual,
      maxAxisType: SparklineAxisType.Individual,
    });
  });

  it('keeps the dialog open on a bad custom value or date range', async () => {
    await mount('mac');
    const dlg = await open();
    dlg.querySelectorAll<HTMLButtonElement>('[role="tab"]')[1]?.click();
    choose(radio(dlg, 'ユーザー設定値:', 1));
    (
      dlg.querySelectorAll<HTMLInputElement>('.fc-macsparkaxis__value')[1] as HTMLInputElement
    ).value = 'abc';
    ok(dlg)?.click();
    expect(document.querySelector('.fc-macsparkaxis')).toBe(dlg);
    choose(radio(dlg, 'スパークラインごとに自動設定', 1));
    dlg.querySelectorAll<HTMLButtonElement>('[role="tab"]')[0]?.click();
    choose(radio(dlg, '日付軸'));
    ok(dlg)?.click();
    expect(document.querySelector('.fc-macsparkaxis')).toBe(dlg);
    expect(settings()?.dateAxis).toBe(false);
    expect(settings()?.maxAxisType).toBe(SparklineAxisType.Individual);
  });

  it('refuses a date range over rows and columns with an alert, keeping the dialog open', async () => {
    // strman_0008 key 45; when the desktop raises it is unmeasured.
    await mount('mac');
    const dlg = await open();
    choose(radio(dlg, '日付軸'));
    const dateField = dlg.querySelector<HTMLInputElement>(
      'input[aria-label="スパークラインのデータ範囲を選んでください:"]',
    ) as HTMLInputElement;
    dateField.value = 'A9:C10';
    ok(dlg)?.click();
    const alert = document.querySelector('.fc-tb__dlg--alert');
    expect(alert?.querySelector('.fc-tb__dlg__message')?.textContent).toBe(
      '日付軸の参照が有効ではありません。すべてのセルが同じ行または列に含まれていません。すべてのセルが 1 つの行または列に含まれるように選択してください。',
    );
    expect(document.querySelector('.fc-macsparkaxis')).toBe(dlg);
    expect(settings()?.dateAxis).toBe(false);
  });

  it('leaves Axis enabled but inert on a protected sheet, with the rest of the tab greyed', async () => {
    await mount('mac');
    setProtectedSheet(sheet.instance.store, 0, true);
    mutators.setActive(sheet.instance.store, at(0));
    const greyed = [
      'sparklineEditData',
      'sparklineTypeLine',
      'sparklineTypeColumn',
      'sparklineTypeWinLoss',
      'sparklineShowHigh',
      'sparklineShowLow',
      'sparklineShowNegative',
      'sparklineShowFirst',
      'sparklineShowLast',
      'sparklineShowMarkers',
      'sparklineColor',
      'sparklineMarkerColor',
      'sparklineGroup',
      'sparklineUngroup',
      'sparklineClear',
    ];
    for (const id of greyed) {
      const button = host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${id}"]`);
      expect(button?.disabled, id).toBe(true);
    }
    const chips = host.querySelectorAll<HTMLButtonElement>('[data-sparkline-quick-style]');
    expect(chips.length).toBeGreaterThan(0);
    for (const chip of chips) expect(chip.disabled).toBe(true);
    expect(axisButton().disabled).toBe(false);
    expect(axisButton().getAttribute('aria-disabled')).not.toBe('true');
    const before = settings();
    axisButton().click();
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(document.querySelector('.fc-macsparkaxis')).toBeNull();
    expect(document.querySelector('.fc-tb__dlg, [role="alertdialog"]')).toBeNull();
    expect(settings()).toEqual(before);
  });
});
