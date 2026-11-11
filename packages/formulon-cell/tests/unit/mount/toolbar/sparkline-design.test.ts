import { BlanksAs, SparklineAxisType, SparklineType } from '@libraz/formulon';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { setProtectedSheet } from '../../../../src/commands/protection.js';
import {
  defaultSparklineGroupSettings,
  groupOfCell,
  listSparklines,
  rgbColorSpec,
  setSparklines,
} from '../../../../src/commands/sparkline.js';
import { Spreadsheet } from '../../../../src/mount.js';
import { mutators } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { stubHelpers } from './fixtures.js';

vi.setConfig({ testTimeout: 20_000 });

const at = (row: number, col = 5) => ({ sheet: 0, row, col });

describe('Sparkline tab of the ribbon', () => {
  let sheet: MountedStubSheet;
  let host: HTMLElement;
  let tb: ReturnType<typeof Spreadsheet.mountToolbar>;

  beforeEach(async () => {
    sheet = await mountStubSheet({ locale: 'en' });
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    tb?.dispose();
    sheet.dispose();
    host.remove();
  });

  const store = () => sheet.instance.store;
  const settings = (row = 0) => groupOfCell(store().getState(), at(row))?.settings;
  const tabIds = (): string[] =>
    Array.from(host.querySelectorAll<HTMLElement>('[data-ribbon-tab]')).map(
      (tab) => tab.dataset.ribbonTab ?? '',
    );
  const command = (id: string): HTMLButtonElement => {
    const found = host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${id}"]`);
    if (!found) throw new Error(`missing ${id}`);
    return found;
  };
  /** Three sparklines in one column, selected, with the toolbar mounted on their tab. */
  const mountOnSparklines = (type: SparklineType = SparklineType.Line): void => {
    setSparklines(
      store(),
      [0, 1, 2].map((row) => ({ addr: at(row), source: `A${row + 1}:C${row + 1}` })),
      defaultSparklineGroupSettings(type),
      sheet.instance.history,
    );
    mutators.setActive(store(), at(0));
    mutators.setRange(store(), { sheet: 0, r0: 0, c0: 5, r1: 2, c1: 5 });
    tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('sparklineDesign');
  };
  /** Click the menu entry `selector` the way the ribbon's click delegation sees it. */
  const pick = (selector: string): void => {
    const target = host.querySelector<HTMLElement>(selector);
    expect(target, selector).toBeTruthy();
    const click = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(click, 'target', { value: target });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(click)).toBe(true);
  };
  const dialog = async (): Promise<HTMLElement> => {
    await vi.waitFor(() => expect(document.querySelector('.fc-tb__dlg')).toBeTruthy());
    return document.querySelector('.fc-tb__dlg') as HTMLElement;
  };
  const field = (root: ParentNode, selector: string): HTMLInputElement => {
    const found = root.querySelector<HTMLInputElement>(selector);
    if (!found) throw new Error(`missing ${selector}`);
    return found;
  };
  const confirm = (dlg: HTMLElement): void =>
    dlg.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();

  it('shows only while the active cell hosts a sparkline', () => {
    tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    expect(tabIds()).not.toContain('sparklineDesign');
    mountOnSparklines();
    expect(tabIds().at(-1)).toBe('sparklineDesign');
    expect(tb.getActiveTab()).toBe('sparklineDesign');
    mutators.setActive(store(), { sheet: 0, row: 9, col: 9 });
    expect(tabIds()).not.toContain('sparklineDesign');
    expect(tb.getActiveTab()).toBe('home');
  });

  it('orders the groups and controls as the desktop tab does', () => {
    mountOnSparklines();
    const ids = Array.from(host.querySelectorAll<HTMLElement>('[data-ribbon-command]'))
      .map((el) => el.dataset.ribbonCommand ?? '')
      .filter((id) => id.startsWith('sparkline'));
    expect(ids).toEqual([
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
      'sparklineStyleGallery',
      'sparklineColor',
      'sparklineMarkerColor',
      'sparklineAxis',
      'sparklineGroup',
      'sparklineUngroup',
      'sparklineClear',
    ]);
  });

  it('presses the toggles of the group and updates it with one undo step each', () => {
    mountOnSparklines();
    expect(command('sparklineTypeLine').getAttribute('aria-pressed')).toBe('true');
    expect(command('sparklineTypeColumn').getAttribute('aria-pressed')).toBe('false');
    command('sparklineTypeColumn').click();
    expect(settings(2)?.type).toBe(SparklineType.Column);
    expect(command('sparklineTypeColumn').getAttribute('aria-pressed')).toBe('true');
    sheet.instance.history.undo();
    expect(settings(2)?.type).toBe(SparklineType.Line);

    for (const [id, key] of [
      ['sparklineShowHigh', 'high'],
      ['sparklineShowLow', 'low'],
      ['sparklineShowFirst', 'first'],
      ['sparklineShowLast', 'last'],
      ['sparklineShowMarkers', 'markers'],
      ['sparklineShowNegative', 'negative'],
    ] as const) {
      command(id).click();
      expect(settings(1)?.[key], id).toBe(true);
      expect(command(id).getAttribute('aria-pressed'), id).toBe('true');
      sheet.instance.history.undo();
      expect(settings(1)?.[key], id).toBe(false);
    }
  });

  it('greys Markers for a column sparkline', () => {
    mountOnSparklines(SparklineType.Column);
    expect(command('sparklineShowMarkers').disabled).toBe(true);
    expect(command('sparklineShowMarkers').getAttribute('aria-disabled')).toBe('true');
    expect(command('sparklineShowHigh').disabled).toBe(false);
    command('sparklineTypeLine').click();
    expect(command('sparklineShowMarkers').disabled).toBe(false);
  });

  it('colours the series, the weight and each marker colour', async () => {
    mountOnSparklines();
    command('sparklineColor').click();
    pick('[data-sparkline-color="#ff0000"]');
    expect(settings()?.colorSeries).toEqual(rgbColorSpec('#ff0000'));

    command('sparklineColor').click();
    pick('[data-format-submenu="weight"]');
    expect(host.querySelector('[data-sparkline-weight="0.75"]')?.getAttribute('aria-checked')).toBe(
      'true',
    );
    pick('[data-sparkline-weight="2.25"]');
    expect(settings(1)?.lineWidthPt).toBe(2.25);

    command('sparklineColor').click();
    pick('[data-sparkline-weight="custom"]');
    const dlg = await dialog();
    field(dlg, 'input').value = '4';
    confirm(dlg);
    await vi.waitFor(() => expect(settings()?.lineWidthPt).toBe(4));

    command('sparklineMarkerColor').click();
    for (const [part, key] of [
      ['negative', 'colorNegative'],
      ['markers', 'colorMarkers'],
      ['high', 'colorHigh'],
      ['low', 'colorLow'],
      ['first', 'colorFirst'],
      ['last', 'colorLast'],
    ] as const) {
      pick(`[data-sparkline-marker-color="${part}:#00b050"]`);
      expect(settings()?.[key], part).toEqual(rgbColorSpec('#00b050'));
      command('sparklineMarkerColor').click();
    }
  });

  it('sets the axis options, a custom minimum through its dialog, and keeps right-to-left greyed', async () => {
    mountOnSparklines();
    command('sparklineAxis').click();
    const checked = (value: string) =>
      host.querySelector(`[data-sparkline-axis="${value}"]`)?.getAttribute('aria-checked');
    expect(checked('general')).toBe('true');
    expect(checked('min-individual')).toBe('true');
    expect(host.querySelector<HTMLButtonElement>('[data-sparkline-axis="rtl"]')?.disabled).toBe(
      true,
    );

    pick('[data-sparkline-axis="show"]');
    expect(settings()?.displayXAxis).toBe(true);
    command('sparklineAxis').click();
    expect(checked('show')).toBe('true');

    pick('[data-sparkline-axis="max-group"]');
    expect(settings()?.maxAxisType).toBe(SparklineAxisType.Group);

    command('sparklineAxis').click();
    pick('[data-sparkline-axis="min-custom"]');
    const dlg = await dialog();
    field(dlg, 'input').value = '-3';
    confirm(dlg);
    await vi.waitFor(() => expect(settings()?.minAxisType).toBe(SparklineAxisType.Custom));
    expect(settings()?.manualMin).toBe(-3);

    command('sparklineAxis').click();
    pick('[data-sparkline-axis="date"]');
    const dates = await dialog();
    field(dates, 'input').value = 'A1:C1';
    confirm(dates);
    await vi.waitFor(() => expect(settings()?.dateAxis).toBe(true));
    expect(settings()?.dateAxisRange).toEqual({
      sheet: 0,
      rect: { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 2 },
    });
    command('sparklineAxis').click();
    pick('[data-sparkline-axis="general"]');
    expect(settings()?.dateAxis).toBe(false);
  });

  it('ungroups, groups and clears the selection', () => {
    mountOnSparklines();
    command('sparklineUngroup').click();
    expect(store().getState().sparkline.groups.size).toBe(3);
    command('sparklineGroup').click();
    expect(store().getState().sparkline.groups.size).toBe(1);
    sheet.instance.history.undo();
    expect(store().getState().sparkline.groups.size).toBe(3);

    mutators.setRange(store(), { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 });
    command('sparklineClear').click();
    expect(listSparklines(store().getState())).toHaveLength(2);

    mutators.setActive(store(), at(1));
    mutators.setRange(store(), { sheet: 0, r0: 1, c0: 5, r1: 1, c1: 5 });
    command('sparklineClear')
      .querySelector('.fc-tb__rb-split-chevron')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    pick('[data-sparkline-clear="groups"]');
    // After the undo the three sparklines are single-member groups, so only this one goes.
    expect(listSparklines(store().getState()).map(({ addr }) => addr.row)).toEqual([2]);
  });

  it('edits one sparkline, the hidden and empty cell settings and the row/column split', async () => {
    mountOnSparklines();
    mutators.setActive(store(), at(1));
    mutators.setRange(store(), { sheet: 0, r0: 1, c0: 5, r1: 1, c1: 5 });
    command('sparklineEditData').click();
    pick('[data-sparkline-edit="single"]');
    const single = await dialog();
    expect(single.querySelector<HTMLInputElement>('input')?.value).toBe('A2:C2');
    field(single, 'input').value = 'D2:F2';
    confirm(single);
    await vi.waitFor(() =>
      expect(listSparklines(store().getState())[1]?.spec.source).toBe('D2:F2'),
    );

    command('sparklineEditData').click();
    pick('[data-sparkline-edit="hidden-empty"]');
    const hidden = await dialog();
    field(hidden, `input[value="${BlanksAs.Span}"]`).click();
    field(hidden, 'input[type="checkbox"]').click();
    confirm(hidden);
    await vi.waitFor(() => expect(settings(1)?.emptyCellsAs).toBe(BlanksAs.Span));
    expect(settings(1)?.displayHidden).toBe(true);

    // The three sources are rows of a square block again, so the split can switch.
    sheet.instance.history.undo();
    sheet.instance.history.undo();
    mutators.setActive(store(), at(0));
    mutators.setRange(store(), { sheet: 0, r0: 0, c0: 5, r1: 2, c1: 5 });
    command('sparklineEditData').click();
    pick('[data-sparkline-edit="switch"]');
    expect(listSparklines(store().getState()).map(({ spec }) => spec.source)).toEqual([
      'A1:A3',
      'B1:B3',
      'C1:C3',
    ]);
  });

  it('edits the group location and data in the Edit Sparklines dialog', async () => {
    mountOnSparklines();
    command('sparklineEditData').click();
    pick('[data-sparkline-edit="group"]');
    await vi.waitFor(() => expect(document.querySelector('.fc-macsparkdlg')).toBeTruthy());
    const dlg = document.querySelector('.fc-macsparkdlg') as HTMLElement;
    expect(dlg.querySelector<HTMLInputElement>('.fc-macsparkdlg__source')?.value).toBe('A1:C3');
    expect(dlg.querySelector<HTMLInputElement>('.fc-macsparkdlg__destination')?.value).toBe(
      'F1:F3',
    );
    const destination = field(dlg, '.fc-macsparkdlg__destination');
    destination.value = 'H1:H3';
    destination.dispatchEvent(new Event('input', { bubbles: true }));
    dlg.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    expect(listSparklines(store().getState()).map(({ addr }) => addr.col)).toEqual([7, 7, 7]);
    expect(settings()).toBeUndefined();
    expect(groupOfCell(store().getState(), at(0, 7))?.settings.type).toBe(SparklineType.Line);
  });

  // The default protection refuses object edits; allowing them must not change what the tab shows.
  it.each([
    ['default options', undefined],
    ['objects allowed', { permissions: { objects: true } }],
  ] as const)(
    'disables the whole tab on a protected sheet except Axis, whose entries are inert (%s)',
    (_label, options) => {
      mountOnSparklines();
      setProtectedSheet(store(), 0, true, options);
      mutators.setActive(store(), at(0));
      const ids = Array.from(host.querySelectorAll<HTMLButtonElement>('[data-ribbon-command]'))
        .filter((el) => el.dataset.ribbonCommand?.startsWith('sparkline'))
        .map((el) => [el.dataset.ribbonCommand, el.disabled] as const);
      expect(ids.filter(([id]) => id !== 'sparklineAxis').every(([, off]) => off)).toBe(true);
      expect(ids.find(([id]) => id === 'sparklineAxis')?.[1]).toBe(false);
      command('sparklineAxis').click();
      const items = host.querySelectorAll<HTMLButtonElement>('[data-sparkline-axis]');
      expect(items).toHaveLength(10);
      expect(Array.from(items).every((item) => item.disabled)).toBe(true);
      expect(settings()?.type).toBe(SparklineType.Line);
    },
  );

  it('lists the 36 style slots, all disabled', () => {
    mountOnSparklines();
    const swatches = host.querySelectorAll<HTMLButtonElement>('[data-sparkline-style]');
    expect(swatches).toHaveLength(36);
    expect(Array.from(swatches).every((swatch) => swatch.disabled)).toBe(true);
    expect(swatches[0]?.getAttribute('aria-label')).toBe('Sparkline Style Accent 1, Darker 50%');
    expect(swatches[35]?.getAttribute('aria-label')).toBe('Sparkline Style Colorful #6');
    // A slot waits for its measured colours and says so, like other unbuilt commands.
    expect(swatches[0]?.getAttribute('aria-description')).toBe('Coming soon');
    command('sparklineStyleGallery').click();
    expect(Array.from(swatches).every((swatch) => swatch.disabled)).toBe(true);
    expect(swatches[0]?.getAttribute('aria-description')).toBe('Coming soon');
  });
});
