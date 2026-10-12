import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { addrKey } from '../../../../src/engine/address.js';
import { effectiveCellFormatFromEngine } from '../../../../src/engine/cell-format-sync.js';
import { defaultFontForLocale, WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { BUILT_IN_LOCALES } from '../../../../src/i18n/locales.js';
import { createDefaultRibbonHelpers } from '../../../../src/mount/toolbar-defaults.js';
import { Spreadsheet } from '../../../../src/mount.js';
import { mutators } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { seedText, stubHelpers } from './fixtures.js';

const target = { sheet: 0, row: 0, col: 0 };

const nativeFontAt = (workbook: WorkbookHandle): ReturnType<WorkbookHandle['getFontRecord']> => {
  const xfIndex = workbook.getCellXfIndex(target.sheet, target.row, target.col);
  const xf = workbook.getCellXf(xfIndex ?? 0);
  return xf ? workbook.getFontRecord(xf.fontIndex) : null;
};

describe('font-size point semantics at the mounted ribbon boundary', () => {
  let sheet: MountedStubSheet;
  let host: HTMLElement;

  beforeEach(async () => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'en-US' });
    expect(workbook.isStub).toBe(false);
    sheet = await mountStubSheet({ workbook, locale: 'en' });
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    sheet.dispose();
    host.remove();
  });

  it('font picker 12 pt stores 16 CSS px and writes a native 12 pt font', () => {
    seedText(sheet, target.row, target.col, 'font boundary');
    const toolbar = Spreadsheet.mountToolbar(host, sheet.instance, {
      helpers: createDefaultRibbonHelpers(sheet.instance),
    });
    try {
      const picker = host.querySelector<HTMLButtonElement>(
        '[data-ribbon-command="fontSize"] button',
      );
      expect(picker).not.toBeNull();
      picker?.click();
      const option = document.querySelector<HTMLButtonElement>(
        '.fc-tb__rb-dd__opt[data-value="12"]',
      );
      expect(option).not.toBeNull();
      option?.click();

      expect(
        sheet.instance.store.getState().format.formats.get(addrKey(target))?.fontSize,
      ).toBeCloseTo(16, 10);
      expect(nativeFontAt(sheet.workbook)?.size).toBeCloseTo(12, 10);
    } finally {
      toolbar.dispose();
    }
  });

  it('font grow and shrink step in Excel points while storing CSS pixels', () => {
    seedText(sheet, target.row, target.col, 'font step');
    const toolbar = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    try {
      expect(toolbar.applyCommand('fontGrow')).toBe(true);
      expect(
        sheet.instance.store.getState().format.formats.get(addrKey(target))?.fontSize,
      ).toBeCloseTo(16, 10);
      expect(nativeFontAt(sheet.workbook)?.size).toBeCloseTo(12, 10);

      expect(toolbar.applyCommand('fontShrink')).toBe(true);
      expect(
        sheet.instance.store.getState().format.formats.get(addrKey(target))?.fontSize,
      ).toBeCloseTo((11 * 96) / 72, 10);
      expect(nativeFontAt(sheet.workbook)?.size).toBeCloseTo(11, 10);
    } finally {
      toolbar.dispose();
    }
  });

  it('refreshes the ribbon point label and selects the same 14 pt value in Format Cells', () => {
    seedText(sheet, target.row, target.col, 'font label');
    const toolbar = Spreadsheet.mountToolbar(host, sheet.instance, {
      helpers: createDefaultRibbonHelpers(sheet.instance),
    });
    const ribbonFontSize = (): string | undefined =>
      host
        .querySelector<HTMLElement>('[data-ribbon-select="fontSize"] .fc-tb__rb-dd__value')
        ?.textContent?.trim();
    try {
      expect(ribbonFontSize()).toBe('11');

      expect(toolbar.applyCommand('fontGrow')).toBe(true);
      expect(ribbonFontSize()).toBe('12');
      expect(toolbar.applyCommand('fontGrow')).toBe(true);
      expect(ribbonFontSize()).toBe('14');
      expect(
        sheet.instance.store.getState().format.formats.get(addrKey(target))?.fontSize,
      ).toBeCloseTo((14 * 96) / 72, 10);
      expect(nativeFontAt(sheet.workbook)?.size).toBeCloseTo(14, 10);

      expect(toolbar.applyCommand('fontGrow')).toBe(true);
      expect(ribbonFontSize()).toBe('16');
      expect(nativeFontAt(sheet.workbook)?.size).toBeCloseTo(16, 10);

      expect(toolbar.applyCommand('fontShrink')).toBe(true);
      expect(ribbonFontSize()).toBe('14');
      expect(nativeFontAt(sheet.workbook)?.size).toBeCloseTo(14, 10);

      sheet.instance.openFormatDialog('font');
      const fourteen = document.querySelector<HTMLButtonElement>(
        '.fc-fmtdlg__font-list--size [data-fc-font-size="14"]',
      );
      expect(fourteen?.getAttribute('aria-selected')).toBe('true');
    } finally {
      document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--cancel')?.click();
      toolbar.dispose();
    }
  });
});

describe.each(BUILT_IN_LOCALES)('locale %s workbook font inheritance', (locale) => {
  it('projects the workbook default family and native point size for an unformatted cell', async () => {
    const workbook = await WorkbookHandle.createDefault({ locale });
    expect(workbook.isStub).toBe(false);
    try {
      const expected = defaultFontForLocale(locale);
      expect(workbook.workbookDisplayFont).toEqual(expected);
      workbook.setText(target, 'inherited');

      expect(
        effectiveCellFormatFromEngine(workbook, target.sheet, target.row, target.col),
      ).toMatchObject({
        fontFamily: expected.name,
        fontSize: (expected.size * 96) / 72,
      });
    } finally {
      workbook.dispose();
    }
  });
});

it.each(['de', 'fr'] as const)(
  'localizes a fractional 11.5 pt ribbon value for %s',
  async (locale) => {
    const workbook = await WorkbookHandle.createDefault({ locale });
    const localSheet = await mountStubSheet({ workbook, locale });
    const localHost = document.createElement('div');
    document.body.appendChild(localHost);
    mutators.setCellFormat(localSheet.instance.store, target, {
      fontSize: (11.5 * 96) / 72,
    });
    const toolbar = Spreadsheet.mountToolbar(localHost, localSheet.instance, {
      helpers: createDefaultRibbonHelpers(localSheet.instance, { lang: locale }),
    });
    try {
      expect(
        localHost
          .querySelector<HTMLElement>('[data-ribbon-select="fontSize"] .fc-tb__rb-dd__value')
          ?.textContent?.trim(),
      ).toBe('11,5');
    } finally {
      toolbar.dispose();
      localSheet.dispose();
      localHost.remove();
    }
  },
);
