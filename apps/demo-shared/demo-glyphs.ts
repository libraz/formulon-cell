import { createExcelRibbonSvg, type ExcelRibbonIconName } from '@libraz/formulon-cell/toolbar';

/** A ribbon glyph from the core icon kit as an SVG element. The core paints its
 *  themed roles, so the glyph follows the title bar's text colour and greys out
 *  when disabled. */
export const createDemoGlyph = (name: ExcelRibbonIconName): SVGSVGElement | null =>
  createExcelRibbonSvg(name, 'fc-tb__rb-icon');
