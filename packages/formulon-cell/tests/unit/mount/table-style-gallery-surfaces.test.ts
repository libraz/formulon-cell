import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import {
  createTableStyleFromActiveTable,
  customTableStyleId,
  formatAsTable,
} from '../../../src/commands/format-as-table.js';
import { Spreadsheet } from '../../../src/mount.js';
import { mutators } from '../../../src/store/store.js';
import type { MountedStubSheet } from '../../test-utils/mount.js';
import { scaledTimeout } from '../../timeouts.js';
import { mountPivotSheet, stubHelpers } from './toolbar/fixtures.js';

vi.setConfig({ testTimeout: scaledTimeout(20_000) });

// Table and PivotTable style galleries as the Mac desktop app shows them (ja).
describe('table style galleries on the Mac profile', () => {
  let sheet: MountedStubSheet;
  let host: HTMLElement;
  let toolbar: ReturnType<typeof Spreadsheet.mountToolbar>;

  beforeEach(async () => {
    sheet = await mountPivotSheet('ja');
    for (const [row, text] of ['Head', 'a', 'b'].entries()) {
      sheet.workbook.setText({ sheet: 0, row: 20 + row, col: 0 }, text);
    }
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    formatAsTable(
      sheet.instance.store,
      { sheet: 0, r0: 20, c0: 0, r1: 22, c1: 0 },
      { workbook: sheet.workbook },
    );
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 21, col: 0 });
    host = document.createElement('div');
    document.body.appendChild(host);
    toolbar = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
      platform: 'mac',
    });
  });

  afterEach(() => {
    toolbar.dispose();
    sheet.dispose();
    host.remove();
    document.body.replaceChildren();
  });

  const strip = () =>
    host.querySelector<HTMLElement>('[data-ribbon-gallery="tableQuickStyles"]') as HTMLElement;
  const chipIds = (): string[] =>
    [...strip().querySelectorAll<HTMLElement>('[data-table-quick-style]')].map(
      (chip) => chip.dataset.tableQuickStyle ?? '',
    );
  const footer = (menuId: string): string[] =>
    [
      ...(host
        .querySelector(`#${menuId}`)
        ?.querySelectorAll<HTMLElement>(':scope > .fc-tb__tablestyle-footer') ?? []),
    ].map((button) => button.dataset.tableStyleFooter ?? '');

  it('shows a strip of seven chips with no page arrows, turned to the current style', () => {
    toolbar.setActiveTab('tableDesign');
    expect(chipIds()).toHaveLength(7);
    expect(strip().querySelector('.fc-tb__tablestyle-quick-prev')).toBeNull();
    expect(chipIds()).toContain('TableStyleMedium2');
  });

  it('puts an applied custom style leftmost in the strip', () => {
    createTableStyleFromActiveTable(sheet.instance.store, sheet.instance.history, 'MyStyle', {
      elements: { headerRow: { bold: true } },
    });
    toolbar.setActiveTab('tableDesign');
    host.querySelector<HTMLButtonElement>('[data-ribbon-command="tableQuickStyles"]')?.click();
    const id = customTableStyleId('MyStyle');
    const item = host.querySelector<HTMLElement>(
      `#menu-table-quick-styles [data-table-style="${id}"]`,
    );
    if (!item) throw new Error('no custom item');
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: item });
    toolbar.dropdownsApi?.dynamicRibbonDropdownClick(event);
    strip().dispatchEvent(new Event('fc-ribbon-control-refresh'));
    expect(chipIds()[0]).toBe(id);
  });

  it('offers None and Clear in the Table Design gallery, New PivotTable Style on Home', () => {
    toolbar.setActiveTab('tableDesign');
    host.querySelector<HTMLButtonElement>('[data-ribbon-command="tableQuickStyles"]')?.click();
    const design = host.querySelector<HTMLElement>('#menu-table-quick-styles');
    const lights = design?.querySelectorAll('.fc-tb__tablestyle-grid')[0];
    expect(lights?.querySelectorAll('.fc-tb__tablestyle-swatch')).toHaveLength(22);
    expect(lights?.querySelector('.fc-tb__tablestyle-swatch')?.getAttribute('aria-label')).toBe(
      'なし',
    );
    expect(footer('menu-table-quick-styles')).toEqual(['new-table-style', 'clear-table-style']);
    toolbar.setActiveTab('home');
    host.querySelector<HTMLButtonElement>('[data-ribbon-command="formatTableHome"]')?.click();
    const home = host.querySelector<HTMLElement>('#menu-table-style-home');
    expect(
      home
        ?.querySelectorAll('.fc-tb__tablestyle-grid')[0]
        ?.querySelectorAll('.fc-tb__tablestyle-swatch'),
    ).toHaveLength(21);
    expect(footer('menu-table-style-home')).toEqual(['new-table-style', 'new-pivot-style']);
  });

  it('clears the table style when None is picked', () => {
    toolbar.setActiveTab('tableDesign');
    host.querySelector<HTMLButtonElement>('[data-ribbon-command="tableQuickStyles"]')?.click();
    const none = host.querySelector<HTMLElement>('#menu-table-quick-styles [data-table-style=""]');
    if (!none) throw new Error('no None');
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: none });
    toolbar.dropdownsApi?.dynamicRibbonDropdownClick(event);
    expect(sheet.instance.store.getState().tables.tables[0]?.styleName).toBe('');
  });

  it('orders the chip menu Apply-Clear first, with mnemonics', () => {
    toolbar.setActiveTab('tableDesign');
    strip()
      .querySelector<HTMLElement>('[data-table-quick-style="TableStyleMedium2"]')
      ?.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true }));
    expect(
      [...host.querySelectorAll<HTMLButtonElement>('.fc-tb__tablestyle-ctx button')].map((b) => [
        b.textContent,
        b.disabled || b.getAttribute('aria-disabled') === 'true',
      ]),
    ).toEqual([
      ['適用 (書式をクリア)(C)', false],
      ['適用 (書式を保持)(M)', false],
      ['変更(Y)...', true],
      ['複製(P)...', false],
      ['削除(L)', true],
      ['既定に設定(D)', false],
    ]);
  });
});
