import { describe, expect, it } from 'vitest';
import {
  hydrateCellFormatsFromEngine,
  syncCellFormatsToEngine,
} from '../../../../src/engine/cell-format-sync.js';
import { addrKey, WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import {
  type CellAlign,
  type CellFormat,
  type CellVAlign,
  createSpreadsheetStore,
  type FillPattern,
  type NegativeStyle,
  type UnderlineStyle,
} from '../../../../src/store/store.js';

import { canLoadWasm } from './fixtures.js';

describe.skipIf(!canLoadWasm())('real xlsx round-trip', () => {
  it('saves and reloads authored cell XF records when supported', async () => {
    const first = await WorkbookHandle.createDefault();

    try {
      expect(first.isStub).toBe(false);
      if (!first.capabilities.cellFormatting) return;

      first.setText({ sheet: 0, row: 0, col: 0 }, 'Formatted');
      const fontIndex = first.addFontRecord({
        name: 'Arial',
        size: 11,
        bold: true,
        italic: false,
        strike: false,
        underline: 0,
        colorArgb: 0xff006100,
      });
      const fillIndex = first.addFillRecord({ pattern: 1, fgArgb: 0xffe2f0d9, bgArgb: 0 });
      const borderIndex = first.addBorderRecord({
        left: { style: 0, colorArgb: 0 },
        right: { style: 0, colorArgb: 0 },
        top: { style: 0, colorArgb: 0 },
        bottom: { style: 0, colorArgb: 0 },
        diagonal: { style: 0, colorArgb: 0 },
        diagonalUp: false,
        diagonalDown: false,
      });
      const xfIndex = first.addXfRecord({
        fontIndex,
        fillIndex,
        borderIndex,
        numFmtId: 0,
        horizontalAlign: 2,
        verticalAlign: 1,
        wrapText: true,
      });
      expect(xfIndex).toBeGreaterThan(0);
      expect(first.setCellXfIndex(0, 0, 0, xfIndex)).toBe(true);

      const bytes = first.save();
      expect(bytes.length).toBeGreaterThan(0);

      const reloaded = await WorkbookHandle.loadBytes(bytes);
      try {
        expect(reloaded.isStub).toBe(false);
        const reloadedXfIndex = reloaded.getCellXfIndex(0, 0, 0);
        expect(reloadedXfIndex).toBeGreaterThan(0);
        expect(reloaded.getCellXf(reloadedXfIndex ?? -1)).toMatchObject({
          horizontalAlign: 2,
          verticalAlign: 1,
          wrapText: true,
        });
      } finally {
        reloaded.dispose();
      }
    } finally {
      first.dispose();
    }
  });

  it('preserves extended alignment, negative styles, and pattern fills through real xlsx', async () => {
    const first = await WorkbookHandle.createDefault();

    try {
      expect(first.isStub).toBe(false);
      if (!first.capabilities.cellFormatting) return;

      const source = createSpreadsheetStore();
      const formats = new Map<string, CellFormat>();
      const horizontal: CellAlign[] = [
        'left',
        'center',
        'right',
        'fill',
        'justify',
        'centerContinuous',
        'distributed',
      ];
      const vertical: CellVAlign[] = ['top', 'middle', 'bottom', 'justify', 'distributed'];
      for (const [row, align] of horizontal.entries()) {
        for (const [col, vAlign] of vertical.entries()) {
          first.setText({ sheet: 0, row, col }, `${align}/${vAlign}`);
          formats.set(addrKey({ sheet: 0, row, col }), { align, vAlign });
        }
      }

      const negativeStyles: NegativeStyle[] = ['minus', 'parens', 'red', 'red-parens'];
      for (const [offset, negativeStyle] of negativeStyles.entries()) {
        const row = 10 + offset;
        first.setNumber({ sheet: 0, row, col: 0 }, -1234.5);
        formats.set(addrKey({ sheet: 0, row, col: 0 }), {
          numFmt: { kind: 'fixed', decimals: 2, thousands: true, negativeStyle },
        });
        first.setNumber({ sheet: 0, row, col: 1 }, -1234.5);
        formats.set(addrKey({ sheet: 0, row, col: 1 }), {
          numFmt: { kind: 'currency', decimals: 2, symbol: '¥', negativeStyle },
        });
      }

      const patterns: FillPattern[] = [
        'gray50',
        'gray75',
        'gray25',
        'darkHorizontal',
        'darkVertical',
        'darkDown',
        'darkUp',
        'darkGrid',
        'darkTrellis',
        'lightHorizontal',
        'lightVertical',
        'lightDown',
        'lightUp',
        'lightGrid',
        'lightTrellis',
        'gray125',
        'gray0625',
      ];
      for (const [col, fillPattern] of patterns.entries()) {
        first.setText({ sheet: 0, row: 20, col }, fillPattern);
        formats.set(addrKey({ sheet: 0, row: 20, col }), {
          fill: '#abcdef',
          fillPattern,
          fillPatternColor: '#123456',
        });
      }

      const underlines: UnderlineStyle[] = [
        'single',
        'double',
        'singleAccounting',
        'doubleAccounting',
      ];
      for (const [col, underline] of underlines.entries()) {
        first.setText({ sheet: 0, row: 21, col }, underline);
        formats.set(addrKey({ sheet: 0, row: 21, col }), { underline });
      }
      const verticalFontAlignments = ['superscript', 'subscript'] as const;
      for (const [col, fontVertAlign] of verticalFontAlignments.entries()) {
        first.setText({ sheet: 0, row: 22, col }, fontVertAlign);
        formats.set(addrKey({ sheet: 0, row: 22, col }), { fontVertAlign });
      }
      first.setText({ sheet: 0, row: 23, col: 0 }, '漢字');
      formats.set(addrKey({ sheet: 0, row: 23, col: 0 }), {
        phonetic: [{ start: 0, end: 2, text: 'かんじ' }],
      });
      // A guide that annotates each kanji separately, which is what collapses
      // into one whole-cell reading without a per-run writeback.
      first.setText({ sheet: 0, row: 24, col: 0 }, '東京都');
      formats.set(addrKey({ sheet: 0, row: 24, col: 0 }), {
        phonetic: [
          { start: 0, end: 2, text: 'とうきょう' },
          { start: 2, end: 3, text: 'と' },
        ],
      });

      source.setState((state) => ({
        ...state,
        format: { ...state.format, formats },
      }));
      syncCellFormatsToEngine(first, source, 0);

      const bytes = first.save();
      expect(bytes.length).toBeGreaterThan(0);

      const reloaded = await WorkbookHandle.loadBytes(bytes);
      try {
        const target = createSpreadsheetStore();
        hydrateCellFormatsFromEngine(reloaded, target, 0);
        const roundTripped = target.getState().format.formats;

        for (const [row, align] of horizontal.entries()) {
          for (const [col, vAlign] of vertical.entries()) {
            const format = roundTripped.get(addrKey({ sheet: 0, row, col }));
            expect(format?.align).toBe(align);
            expect(format?.vAlign ?? 'bottom').toBe(vAlign);
          }
        }
        for (const [offset, negativeStyle] of negativeStyles.entries()) {
          for (const col of [0, 1]) {
            const numFmt = roundTripped.get(addrKey({ sheet: 0, row: 10 + offset, col }))?.numFmt;
            expect(numFmt?.kind).toBe(col === 0 ? 'fixed' : 'currency');
            if (numFmt?.kind === 'fixed' || numFmt?.kind === 'currency') {
              expect(numFmt.negativeStyle ?? 'minus').toBe(negativeStyle);
            }
          }
        }
        for (const [col, fillPattern] of patterns.entries()) {
          expect(roundTripped.get(addrKey({ sheet: 0, row: 20, col }))).toMatchObject({
            fill: '#abcdef',
            fillPattern,
            fillPatternColor: '#123456',
          });
        }
        for (const [col, underline] of underlines.entries()) {
          expect(roundTripped.get(addrKey({ sheet: 0, row: 21, col }))?.underline).toBe(underline);
        }
        for (const [col, fontVertAlign] of verticalFontAlignments.entries()) {
          expect(roundTripped.get(addrKey({ sheet: 0, row: 22, col }))?.fontVertAlign).toBe(
            fontVertAlign,
          );
        }
        expect(roundTripped.get(addrKey({ sheet: 0, row: 23, col: 0 }))?.phonetic).toEqual([
          { start: 0, end: 2, text: 'かんじ' },
        ]);
        expect(roundTripped.get(addrKey({ sheet: 0, row: 24, col: 0 }))?.phonetic).toEqual([
          { start: 0, end: 2, text: 'とうきょう' },
          { start: 2, end: 3, text: 'と' },
        ]);
      } finally {
        reloaded.dispose();
      }
    } finally {
      first.dispose();
    }
  });

  it('saves gallery styles as named cell styles and rehydrates the tag', async () => {
    const first = await WorkbookHandle.createDefault();

    try {
      if (!first.capabilities.cellStyleMutate) return;
      const source = createSpreadsheetStore();
      first.setText({ sheet: 0, row: 0, col: 0 }, 'styled');
      first.setText({ sheet: 0, row: 1, col: 0 }, 'custom');
      first.setText({ sheet: 0, row: 2, col: 0 }, 'plain');
      source.setState((state) => ({
        ...state,
        format: {
          ...state.format,
          customCellStyles: [
            { id: 'custom:Brand', label: 'Brand', format: { bold: true, color: '#7030a0' } },
          ],
          formats: new Map<string, CellFormat>([
            [
              addrKey({ sheet: 0, row: 0, col: 0 }),
              { cellStyle: 'good', color: '#006100', fill: '#c6efce' },
            ],
            [
              addrKey({ sheet: 0, row: 1, col: 0 }),
              { cellStyle: 'Brand', bold: true, color: '#7030a0' },
            ],
            [addrKey({ sheet: 0, row: 2, col: 0 }), { italic: true }],
          ]),
        },
      }));
      syncCellFormatsToEngine(first, source, 0);

      const reloaded = await WorkbookHandle.loadBytes(first.save());
      try {
        const named = reloaded.getNamedCellStyles();
        expect(named.map((s) => s.name).sort()).toEqual(['Brand', 'Good', 'Normal']);
        expect(named.find((s) => s.name === 'Good')?.builtinId).toBe(26);

        const target = createSpreadsheetStore();
        hydrateCellFormatsFromEngine(reloaded, target, 0);
        const roundTripped = target.getState().format.formats;
        expect(roundTripped.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.cellStyle).toBe('good');
        expect(roundTripped.get(addrKey({ sheet: 0, row: 1, col: 0 }))?.cellStyle).toBe('Brand');
        expect(roundTripped.get(addrKey({ sheet: 0, row: 2, col: 0 }))?.cellStyle).toBeUndefined();

        // The style's own formatting lives on its <cellStyleXfs> row, which is
        // what makes editing the style reach every cell that uses it.
        const good = named.find((s) => s.name === 'Good');
        const styleXf = good ? reloaded.getCellStyleXf(good.xfId) : null;
        expect(styleXf).not.toBeNull();
        expect(reloaded.getFillRecord(styleXf?.fillIndex ?? 0)?.pattern).toBe(1);
      } finally {
        reloaded.dispose();
      }
    } finally {
      first.dispose();
    }
  });
});
