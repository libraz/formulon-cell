import { expect, it } from 'vitest';
import { effectiveCellFormatFromEngine } from '../../../../src/engine/cell-format-sync.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { createDefaultRibbonHelpers } from '../../../../src/mount/toolbar-defaults.js';
import { Spreadsheet } from '../../../../src/mount.js';
import { mutators } from '../../../../src/store/store.js';
import { mountStubSheet } from '../../../test-utils/mount.js';
import { seedText, stubHelpers } from './fixtures.js';

it('applies a ribbon font size change after selecting the entire sheet', async () => {
  const workbook = await WorkbookHandle.createDefault({ locale: 'ja' });
  expect(workbook.isStub).toBe(false);
  const sheet = await mountStubSheet({ workbook, locale: 'ja' });
  const host = document.createElement('div');
  document.body.appendChild(host);
  seedText(sheet, 0, 0, 'all cells');
  const toolbar = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
  try {
    mutators.selectAll(sheet.instance.store);
    expect(toolbar.applyCommand('fontGrow')).toBe(true);
    expect(effectiveCellFormatFromEngine(workbook, 0, 0, 0)?.fontSize).toBeCloseTo(16, 10);
  } finally {
    toolbar.dispose();
    sheet.dispose();
    host.remove();
  }
});

it('keeps the ribbon value and history unchanged when a protected selection rejects a font size', async () => {
  const workbook = await WorkbookHandle.createDefault({ locale: 'ja' });
  expect(workbook.isStub).toBe(false);
  const sheet = await mountStubSheet({ workbook, locale: 'ja' });
  const host = document.createElement('div');
  document.body.appendChild(host);
  const toolbar = Spreadsheet.mountToolbar(host, sheet.instance, {
    helpers: createDefaultRibbonHelpers(sheet.instance, { lang: 'ja' }),
  });
  try {
    mutators.selectAll(sheet.instance.store);
    sheet.instance.store.setState((state) => ({
      ...state,
      protection: {
        ...state.protection,
        protectedSheets: new Map([[0, {}]]),
      },
    }));

    const value = (): string | undefined =>
      host
        .querySelector<HTMLElement>('[data-ribbon-select="fontSize"] .fc-tb__rb-dd__value')
        ?.textContent?.trim();
    expect(value()).toBe('11');
    expect(sheet.instance.history.canUndo()).toBe(false);

    host.querySelector<HTMLButtonElement>('[data-ribbon-command="fontSize"] button')?.click();
    const option = document.querySelector<HTMLButtonElement>('.fc-tb__rb-dd__opt[data-value="18"]');
    expect(option).not.toBeNull();
    option?.click();

    expect(value()).toBe('11');
    expect(sheet.instance.history.canUndo()).toBe(false);
    expect(effectiveCellFormatFromEngine(workbook, 0, 0, 0)?.fontSize).toBeCloseTo(
      11 * (96 / 72),
      10,
    );

    sheet.instance.store.setState((state) => ({
      ...state,
      protection: { ...state.protection, protectedSheets: new Map() },
    }));
    host.querySelector<HTMLButtonElement>('[data-ribbon-command="fontSize"] button')?.click();
    const accepted = document.querySelector<HTMLButtonElement>(
      '.fc-tb__rb-dd__opt[data-value="18"]',
    );
    expect(accepted).not.toBeNull();
    accepted?.click();
    expect(value()).toBe('18');
    expect(sheet.instance.history.canUndo()).toBe(true);
    expect(effectiveCellFormatFromEngine(workbook, 0, 0, 0)?.fontSize).toBeCloseTo(
      18 * (96 / 72),
      10,
    );
  } finally {
    toolbar.dispose();
    sheet.dispose();
    host.remove();
  }
});
