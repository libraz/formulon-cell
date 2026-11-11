import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { createDefaultRibbonHooks } from '../../../../src/mount/toolbar-defaults.js';
import { mutators } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { seedText } from './fixtures.js';

const reportedCells = (): string[] =>
  [...document.querySelectorAll('.fc-tb__dlg__list strong')].map(
    (label) => label.textContent?.split(' · ')[1] ?? '',
  );

describe('Review > Spelling scope', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet({ locale: 'en' });
    seedText(sheet, 0, 0, 'teh');
    seedText(sheet, 4, 2, 'teh');
  });

  afterEach(() => {
    sheet.dispose();
    document.body.replaceChildren();
  });

  it('checks the whole sheet from one selected cell', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 9, col: 9 });
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 9, c0: 9, r1: 9, c1: 9 });
    createDefaultRibbonHooks(sheet.instance).review?.spelling?.();
    expect(reportedCells()).toEqual(['A1', 'C5']);
  });

  it('checks only a selected range', () => {
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    createDefaultRibbonHooks(sheet.instance).review?.spelling?.();
    expect(reportedCells()).toEqual(['A1']);
  });
});
