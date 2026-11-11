import { describe, expect, it } from 'vitest';
import {
  createExcelRibbonSvg,
  type ExcelRibbonIconName,
} from '../../../src/toolbar/excel-ribbon-icons.js';

describe('title-bar glyphs and the public ribbon icon export', () => {
  it.each<ExcelRibbonIconName>(['home', 'repeat', 'moreCommands', 'feedback'])(
    'renders %s as an svg with the class',
    (name) => {
      const svg = createExcelRibbonSvg(name, 'demo__icon');
      expect(svg?.classList.contains('demo__icon')).toBe(true);
      expect(svg?.querySelectorAll('path').length).toBeGreaterThan(0);
    },
  );

  it('returns null for a name the set does not define', () => {
    expect(createExcelRibbonSvg('noSuchGlyph', 'demo__icon')).toBeNull();
  });
});
