import { describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

describe('engine theme colour resolution', () => {
  it('resolves selectors for display while preserving them through edits and save/load', async () => {
    const wb = await WorkbookHandle.createDefault();
    let reloaded: WorkbookHandle | null = null;
    try {
      expect(wb.isStub).toBe(false);
      expect(wb.capabilities.colorResolution).toBe(true);
      const normal = wb.getCellStyleXf(0);
      const defaultFont = wb.getFontRecord(0);
      const defaultFill = wb.getFillRecord(0);
      const defaultBorder = wb.getBorderRecord(0);
      expect(normal).toBeDefined();
      expect(defaultFont).toBeDefined();
      expect(defaultFill).toBeDefined();
      expect(defaultBorder).toBeDefined();
      if (!normal || !defaultFont || !defaultFill || !defaultBorder) return;

      const themedFontIndex = wb.addFontRecord({
        ...defaultFont,
        colorArgb: 0xff010203,
        color: { kind: 2, rgb: 0, theme: 4, tint: 0.4, indexed: 0 },
      });
      const indexedAndAutoFillIndex = wb.addFillRecord({
        ...defaultFill,
        pattern: 1,
        fgArgb: 0xff040506,
        fg: { kind: 3, rgb: 0, theme: 0, tint: 0, indexed: 2 },
        bgArgb: 0xff070809,
        bg: { kind: 4, rgb: 0, theme: 0, tint: 0, indexed: 0 },
      });
      const autoBorderIndex = wb.addBorderRecord({
        ...defaultBorder,
        top: {
          ...defaultBorder.top,
          style: 1,
          colorArgb: 0xff0a0b0c,
          color: { kind: 4, rgb: 0, theme: 0, tint: 0, indexed: 0 },
        },
      });

      expect(wb.getFontRecord(themedFontIndex)).toMatchObject({
        colorArgb: 0xff010203,
        color: { kind: 2, theme: 4, tint: 0.4 },
      });
      // The unreleased engine tracks the current Office default theme. Its
      // accent1 is #156082, whose 40% tint resolves to #44B3E1. The older
      // #4472C4 theme would produce #8EA9DB, which is no longer the native
      // default for a workbook without an explicit theme part.
      expect(wb.getResolvedFontRecord(themedFontIndex)).toMatchObject({
        colorArgb: 0xff44b3e1,
        color: { kind: 2, theme: 4, tint: 0.4 },
      });
      expect(wb.getFillRecord(indexedAndAutoFillIndex)).toMatchObject({
        fgArgb: 0xff040506,
        bgArgb: 0xff070809,
        fg: { kind: 3, indexed: 2 },
        bg: { kind: 4 },
      });
      expect(wb.getResolvedFillRecord(indexedAndAutoFillIndex)).toMatchObject({
        fgArgb: 0xffff0000,
        bgArgb: 0xffffffff,
        fg: { kind: 3, indexed: 2 },
        bg: { kind: 4 },
      });
      expect(wb.getBorderRecord(autoBorderIndex)?.top).toMatchObject({
        colorArgb: 0xff0a0b0c,
        color: { kind: 4 },
      });
      expect(wb.getResolvedBorderRecord(autoBorderIndex)?.top).toMatchObject({
        colorArgb: 0xff000000,
        color: { kind: 4 },
      });

      const xfIndex = wb.addXfRecord({
        ...normal,
        fontIndex: themedFontIndex,
        fillIndex: indexedAndAutoFillIndex,
        borderIndex: autoBorderIndex,
      });
      expect(wb.setCellXfIndex(0, 0, 0, xfIndex)).toBe(true);
      wb.setText({ sheet: 0, row: 0, col: 1 }, 'unrelated edit');

      reloaded = await WorkbookHandle.loadBytes(wb.save());
      const loadedXf = reloaded.getCellXf(reloaded.getCellXfIndex(0, 0, 0) ?? -1);
      expect(loadedXf).toBeDefined();
      if (!loadedXf) return;
      expect(reloaded.getFontRecord(loadedXf.fontIndex)?.color).toMatchObject({
        kind: 2,
        theme: 4,
        tint: 0.4,
      });
      expect(reloaded.getResolvedFontRecord(loadedXf.fontIndex)?.colorArgb).toBe(0xff44b3e1);
      expect(reloaded.getFillRecord(loadedXf.fillIndex)?.fg).toMatchObject({
        kind: 3,
        indexed: 2,
      });
      expect(reloaded.getResolvedFillRecord(loadedXf.fillIndex)).toMatchObject({
        fgArgb: 0xffff0000,
        bgArgb: 0xffffffff,
      });
      expect(reloaded.getBorderRecord(loadedXf.borderIndex)?.top.color).toMatchObject({ kind: 4 });
      expect(reloaded.getResolvedBorderRecord(loadedXf.borderIndex)?.top.colorArgb).toBe(
        0xff000000,
      );
    } finally {
      reloaded?.dispose();
      wb.dispose();
    }
  });
});
