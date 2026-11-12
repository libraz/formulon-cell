import { BlanksAs, SparklineType } from '@libraz/formulon';
import { afterEach, describe, expect, it, vi } from 'vitest';
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

const LARGE = [
  'sparklineEditData',
  'sparklineTypeLine',
  'sparklineTypeColumn',
  'sparklineTypeWinLoss',
  'sparklineColor',
  'sparklineMarkerColor',
  'sparklineAxis',
];
const SHOW = [
  'sparklineShowHigh',
  'sparklineShowLow',
  'sparklineShowNegative',
  'sparklineShowFirst',
  'sparklineShowLast',
  'sparklineShowMarkers',
];
const STACKED = ['sparklineGroup', 'sparklineUngroup', 'sparklineClear'];

describe('Sparkline tab layout', () => {
  let sheet: MountedStubSheet;
  let host: HTMLElement;
  let tb: ReturnType<typeof Spreadsheet.mountToolbar>;

  afterEach(() => {
    tb.dispose();
    sheet.dispose();
    host.remove();
    document.body.replaceChildren();
  });

  /** Three single-row sparklines in one group, all selected, with the tab showing. */
  const mount = async (locale: 'en' | 'ja', platform?: 'mac'): Promise<void> => {
    sheet = await mountStubSheet({ locale });
    host = document.createElement('div');
    document.body.appendChild(host);
    const store = sheet.instance.store;
    setSparklines(
      store,
      [0, 1, 2].map((row) => ({ addr: at(row), source: `A${row + 1}:C${row + 1}` })),
      defaultSparklineGroupSettings(SparklineType.Line),
      sheet.instance.history,
    );
    mutators.setActive(store, at(0));
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 5, r1: 2, c1: 5 });
    tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
      ...(platform ? { platform } : {}),
    });
    tb.setActiveTab('sparklineDesign');
  };
  const command = (id: string): HTMLButtonElement => {
    const found = host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${id}"]`);
    if (!found) throw new Error(`missing ${id}`);
    return found;
  };
  const groupOf = (id: string): HTMLElement | null => command(id).closest('.fc-tb__ribbon-group');

  it('sizes the controls as the desktop tab does', async () => {
    await mount('en');
    for (const id of LARGE) expect(command(id).classList, id).toContain('fc-tb__rb--large');
    for (const id of SHOW) {
      expect(command(id).classList, id).toContain('fc-tb__rb--stacked');
      expect(command(id).classList, id).toContain('fc-tb__rb--mac-checkbox');
    }
    for (const id of STACKED) expect(command(id).classList, id).toContain('fc-tb__rb--stacked');
  });

  it('stacks Show in columns of three and Group, Ungroup and Clear beside Axis', async () => {
    await mount('en');
    expect(groupOf('sparklineShowHigh')?.classList).toContain('fc-tb__ribbon-group--stacked');
    expect(groupOf('sparklineShowMarkers')).toBe(groupOf('sparklineShowHigh'));
    const axisGroup = groupOf('sparklineAxis');
    expect(axisGroup?.classList).toContain('fc-tb__ribbon-group--mixed');
    for (const id of STACKED) expect(groupOf(id)).toBe(axisGroup);
  });

  it('opens Edit Group Location & Data from the face of Edit Data and the menu from its arrow', async () => {
    await mount('en');
    const editData = command('sparklineEditData');
    expect(editData.dataset.ribbonActivation).toBe('splitPrimary');
    editData.click();
    await vi.waitFor(() => expect(document.querySelector('.fc-macsparkdlg')).toBeTruthy());
    const dlg = document.querySelector('.fc-macsparkdlg') as HTMLElement;
    expect(dlg.hidden).toBe(false);
    expect(dlg.querySelector<HTMLInputElement>('.fc-macsparkdlg__source')?.value).toBe(
      'Sheet1!A1:C3',
    );

    editData
      .querySelector('.fc-tb__rb-split-chevron')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(host.querySelector<HTMLElement>('#menu-sparkline-edit-data')?.hidden).toBe(false);
  });

  it('opens Edit Sparklines with the sheet-named data and the pinned location on the Mac', async () => {
    await mount('ja', 'mac');
    command('sparklineEditData').click();
    await vi.waitFor(() => expect(document.querySelector('.fc-macsparkdlg')).toBeTruthy());
    const dlg = document.querySelector('.fc-macsparkdlg') as HTMLElement;
    expect(dlg.querySelector('.fc-fmtdlg__header')?.textContent).toBe('スパークラインの編集');
    const source = dlg.querySelector<HTMLInputElement>('.fc-macsparkdlg__source');
    const destination = dlg.querySelector<HTMLInputElement>('.fc-macsparkdlg__destination');
    expect(source?.getAttribute('aria-label')).toBe('スパークラインのデータ範囲を選んでください:');
    expect(destination?.getAttribute('aria-label')).toBe(
      'スパークラインを配置する場所を選んでください:',
    );
    expect(source?.value).toBe('Sheet1!A1:C3');
    expect(destination?.value).toBe('Sheet1!$F$1:$F$3');
    expect(dlg.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.disabled).toBe(false);
  });

  it('lists the Edit Data menu with the Hidden & Empty Cells submenu on the Mac', async () => {
    await mount('ja', 'mac');
    command('sparklineEditData')
      .querySelector('.fc-tb__rb-split-chevron')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    const menu = host.querySelector<HTMLElement>('#menu-sparkline-edit-data');
    if (!menu) throw new Error('no Edit Data menu');
    const rows = Array.from(menu.children)
      .filter((el) => !el.classList.contains('fc-tb__submenu'))
      .map((el) =>
        el.getAttribute('role') === 'separator' ? '|' : (el.textContent ?? '').replace('›', ''),
      );
    expect(rows).toEqual([
      'グループの位置とデータの編集...',
      '単一スパークラインのデータの編集...',
      '|',
      '非表示および空白のセル',
      '行/列の切り替え',
    ]);
    const panel = menu.querySelector<HTMLElement>('.fc-tb__submenu');
    const entries = Array.from(panel?.children ?? []).map((el) =>
      el.getAttribute('role') === 'separator'
        ? '|'
        : `${el.getAttribute('role')}:${el.textContent ?? ''}`,
    );
    expect(entries).toEqual([
      'menuitemradio:空のセルを空白として表示する',
      'menuitemradio:空のセルをゼロとして表示する',
      'menuitemradio:データ ポイントと線を接続する',
      '|',
      'menuitemcheckbox:非表示の行と列のデータを表示する',
    ]);
  });

  it('opens the Hidden & Empty Cells submenu on a click of its Mac entry, with no dialog', async () => {
    await mount('ja', 'mac');
    command('sparklineEditData')
      .querySelector('.fc-tb__rb-split-chevron')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(host.querySelector('[data-sparkline-edit="hidden-empty"]')).toBeNull();
    const trigger = host.querySelector<HTMLElement>('[data-format-submenu="hidden-empty"]');
    if (!trigger) throw new Error('no Hidden & Empty Cells entry');
    const click = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(click, 'target', { value: trigger });
    tb.dropdownsApi?.dynamicRibbonDropdownClick(click);
    await Promise.resolve();
    expect(host.querySelector<HTMLElement>('#menu-sparkline-edit-data-hidden-empty')?.hidden).toBe(
      false,
    );
    expect(document.querySelector('.fc-tb__dlg')).toBeNull();
  });

  it('opens the Hidden and Empty Cell Settings dialog from a plain Edit Data entry elsewhere', async () => {
    await mount('en');
    command('sparklineEditData')
      .querySelector('.fc-tb__rb-split-chevron')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    const menu = host.querySelector<HTMLElement>('#menu-sparkline-edit-data');
    expect(menu?.querySelector('.fc-tb__submenu')).toBeNull();
    expect(menu?.querySelector('[data-sparkline-empty]')).toBeNull();
    const entry = menu?.querySelector<HTMLElement>('[data-sparkline-edit="hidden-empty"]');
    expect(entry?.textContent).toBe('Hidden & Empty Cells...');
    expect(entry?.hasAttribute('aria-haspopup')).toBe(false);
  });

  it('sets the empty-cell display and hidden data from the Mac submenu, one undo step each', async () => {
    await mount('en', 'mac');
    const settings = () => groupOfCell(sheet.instance.store.getState(), at(0))?.settings;
    const pick = (value: string): void => {
      const target = host.querySelector<HTMLElement>(`[data-sparkline-empty="${value}"]`);
      expect(target, value).toBeTruthy();
      const click = new MouseEvent('click', { bubbles: true });
      Object.defineProperty(click, 'target', { value: target });
      expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(click)).toBe(true);
    };
    const checked = (value: string) =>
      host.querySelector(`[data-sparkline-empty="${value}"]`)?.getAttribute('aria-checked');
    const openEditData = (): void => {
      command('sparklineEditData')
        .querySelector('.fc-tb__rb-split-chevron')
        ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    };
    openEditData();
    expect(checked(String(BlanksAs.Gap))).toBe('true');
    pick(String(BlanksAs.Zero));
    expect(settings()?.emptyCellsAs).toBe(BlanksAs.Zero);
    // The pick hands focus back to the grid, so Cmd+Z undoes it rather than the name box.
    expect(document.activeElement).toBe(sheet.instance.host);
    openEditData();
    expect(checked(String(BlanksAs.Zero))).toBe('true');
    expect(checked(String(BlanksAs.Gap))).toBe('false');
    pick('hidden');
    expect(settings()?.displayHidden).toBe(true);
    openEditData();
    expect(checked('hidden')).toBe('true');
    sheet.instance.history.undo();
    expect(settings()?.displayHidden).toBe(false);
    expect(settings()?.emptyCellsAs).toBe(BlanksAs.Zero);
    sheet.instance.history.undo();
    expect(settings()?.emptyCellsAs).toBe(BlanksAs.Gap);
  });

  it('labels the Group button apart from its group title in ja', async () => {
    await mount('ja', 'mac');
    expect(command('sparklineGroup').textContent).toBe('グループ化');
    expect(command('sparklineUngroup').textContent).toBe('グループ解除');
    expect(groupOf('sparklineGroup')?.getAttribute('aria-label')).toBe('グループ');
  });

  it('greys Group while the selection holds a single group', async () => {
    await mount('en');
    expect(command('sparklineGroup').disabled).toBe(true);
    command('sparklineUngroup').click();
    expect(command('sparklineGroup').disabled).toBe(false);
    command('sparklineGroup').click();
    expect(command('sparklineGroup').disabled).toBe(true);
  });
});
