import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { createDefaultRibbonHooks } from '../../../../src/mount/toolbar-defaults.js';
import { mutators } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { seedText } from './fixtures.js';

const reportedCells = (): string[] =>
  [...document.querySelectorAll('.fc-tb__dlg__list strong')].map(
    (label) => label.textContent?.split(' · ')[1] ?? '',
  );

describe('Review > Spelling scope and order', () => {
  let sheet: MountedStubSheet;

  const select = (active: [number, number], range: [number, number, number, number]): void => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: active[0], col: active[1] });
    const [r0, c0, r1, c1] = range;
    mutators.setRange(sheet.instance.store, { sheet: 0, r0, c0, r1, c1 });
  };
  const spell = (): void => {
    createDefaultRibbonHooks(sheet.instance).review?.spelling?.();
  };

  beforeEach(async () => {
    sheet = await mountStubSheet({ locale: 'en' });
    seedText(sheet, 0, 0, 'teh');
    seedText(sheet, 1, 1, 'hello');
    seedText(sheet, 4, 2, 'recieve');
  });

  afterEach(() => {
    sheet.dispose();
    document.body.replaceChildren();
  });

  it('from one selected cell checks the whole sheet, starting there and wrapping to the top', () => {
    select([1, 1], [1, 1, 1, 1]);
    spell();
    expect(reportedCells()).toEqual(['C5', 'A1']);
  });

  it('checks only a selected range', () => {
    select([0, 0], [0, 0, 2, 0]);
    spell();
    expect(reportedCells()).toEqual(['A1']);
  });
});
