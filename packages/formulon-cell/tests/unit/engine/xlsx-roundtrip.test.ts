import { describe, expect, it } from 'vitest';
import { applyValueFilter, clearFilter } from '../../../src/commands/filter.js';
import { formatAsTable } from '../../../src/commands/format-as-table.js';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import { setSheetRightToLeft } from '../../../src/commands/view.js';
import {
  hydrateAutoFilterFromEngine,
  syncAutoFilterToEngine,
} from '../../../src/engine/auto-filter-sync.js';
import {
  hydrateCellFormatsFromEngine,
  syncCellFormatsToEngine,
} from '../../../src/engine/cell-format-sync.js';
import { syncConditionalRulesToEngine } from '../../../src/engine/cf-writeback.js';
import { hydrateLayoutFromEngine } from '../../../src/engine/layout-sync.js';
import { tableOverlaysFromEngine } from '../../../src/engine/table-sync.js';
import { PivotAggregation, PivotReportLayout, type Range } from '../../../src/engine/types.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { formatPresetPatch } from '../../../src/interact/conditional-dialog-spec.js';
import {
  type CellAlign,
  type CellFormat,
  type CellVAlign,
  type ConditionalRule,
  createSpreadsheetStore,
  type FillPattern,
  type NegativeStyle,
  type SpreadsheetStore,
  type UnderlineStyle,
} from '../../../src/store/store.js';

const canLoadWasm = (): boolean =>
  typeof WebAssembly !== 'undefined' && typeof SharedArrayBuffer !== 'undefined';

/** A two-column header plus two data rows — the smallest grid `formatAsTable`
 * can name every column of. */
const seedTableCells = (wb: WorkbookHandle): void => {
  wb.setText({ sheet: 0, row: 0, col: 0 }, 'Region');
  wb.setText({ sheet: 0, row: 0, col: 1 }, 'Amount');
  wb.setText({ sheet: 0, row: 1, col: 0 }, 'East');
  wb.setNumber({ sheet: 0, row: 1, col: 1 }, 10);
  wb.setText({ sheet: 0, row: 2, col: 0 }, 'West');
  wb.setNumber({ sheet: 0, row: 2, col: 1 }, 20);
};

const seedStoreText = (store: SpreadsheetStore, row: number, col: number, value: string): void => {
  store.setState((state) => {
    const cells = new Map(state.data.cells);
    cells.set(addrKey({ sheet: 0, row, col }), { value: { kind: 'text', value }, formula: null });
    return { ...state, data: { ...state.data, cells } };
  });
};

describe.skipIf(!canLoadWasm())('real xlsx round-trip', () => {
  it('saves and reloads values and formulas through the real engine', async () => {
    const first = await WorkbookHandle.createDefault();

    try {
      first.setNumber({ sheet: 0, row: 0, col: 0 }, 40);
      first.setNumber({ sheet: 0, row: 0, col: 1 }, 2);
      first.setFormula({ sheet: 0, row: 0, col: 2 }, '=A1+B1');
      first.recalc();

      const bytes = first.save();
      expect(bytes.length).toBeGreaterThan(0);

      const cleared = await WorkbookHandle.loadBytes(bytes);
      try {
        expect(cleared.isStub).toBe(false);
        expect(cleared.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
          kind: 'number',
          value: 40,
        });
        expect(cleared.cellFormula({ sheet: 0, row: 0, col: 2 })).toBe('=A1+B1');
        cleared.recalc();
        expect(cleared.getValue({ sheet: 0, row: 0, col: 2 })).toEqual({
          kind: 'number',
          value: 42,
        });
      } finally {
        cleared.dispose();
      }
    } finally {
      first.dispose();
    }
  });

  it('round-trips sheet display flags through the real engine', async () => {
    const first = await WorkbookHandle.createDefault();

    try {
      expect(first.isStub).toBe(false);
      if (!first.capabilities.sheetViewFlags) return;

      expect(first.setSheetShowGridLines(0, false)).toBe(true);
      expect(first.setSheetShowRowColHeaders(0, false)).toBe(true);
      expect(first.setSheetShowZeros(0, false)).toBe(true);
      expect(first.setSheetRightToLeft(0, true)).toBe(true);

      const reloaded = await WorkbookHandle.loadBytes(first.save());
      try {
        expect(reloaded.getSheetView(0)).toMatchObject({
          showGridLines: false,
          showRowColHeaders: false,
          showZeros: false,
          rightToLeft: true,
        });
      } finally {
        reloaded.dispose();
      }
    } finally {
      first.dispose();
    }
  });

  it('round-trips the sheet direction into the store and back out', async () => {
    const first = await WorkbookHandle.createDefault();

    try {
      if (!first.capabilities.sheetViewFlags) return;
      const store = createSpreadsheetStore();
      expect(store.getState().ui.rightToLeft).toBe(false);

      setSheetRightToLeft(store, true, first);
      expect(store.getState().ui.rightToLeft).toBe(true);

      const reloaded = await WorkbookHandle.loadBytes(first.save());
      try {
        const restored = createSpreadsheetStore();
        hydrateLayoutFromEngine(reloaded, restored, 0);
        expect(restored.getState().ui.rightToLeft).toBe(true);
      } finally {
        reloaded.dispose();
      }
    } finally {
      first.dispose();
    }
  });

  it('saves and reloads workbook metadata supported by the real engine', async () => {
    const first = await WorkbookHandle.createDefault();

    try {
      expect(first.isStub).toBe(false);

      if (first.capabilities.definedNameMutate) {
        first.setNumber({ sheet: 0, row: 0, col: 0 }, 4);
        first.setNumber({ sheet: 0, row: 0, col: 1 }, 7);
        expect(first.setDefinedNameEntry('RoundTripName', 'Sheet1!$A$1')).toBe(true);
        expect(first.setDefinedNameEntry('LocalRoundTripName', 'Sheet1!$B$1', 0)).toBe(true);
        first.setFormula({ sheet: 0, row: 0, col: 3 }, '=RoundTripName');
        first.setFormula({ sheet: 0, row: 0, col: 4 }, '=LocalRoundTripName');
        first.recalc();
        expect(first.getValue({ sheet: 0, row: 0, col: 3 })).toEqual({ kind: 'number', value: 4 });
        expect(first.getValue({ sheet: 0, row: 0, col: 4 })).toEqual({ kind: 'number', value: 7 });
      }
      if (first.capabilities.hyperlinks) {
        expect(
          first.addHyperlink(
            0,
            1,
            1,
            'https://example.com/roundtrip',
            'Round Trip',
            'Open round-trip link',
          ),
        ).toBe(true);
      }
      if (first.capabilities.comments) {
        expect(first.setCommentEntry(0, 2, 2, 'Formulon', 'Round-trip comment')).toBe(true);
        expect(first.setCommentEntry(0, 4, 4, 'Formulon', 'Blank-cell comment')).toBe(true);
      }

      const bytes = first.save();
      expect(bytes.length).toBeGreaterThan(0);

      const reloaded = await WorkbookHandle.loadBytes(bytes);
      try {
        expect(reloaded.isStub).toBe(false);

        if (first.capabilities.definedNameMutate) {
          const names = [...reloaded.definedNames()];
          expect(names).toContainEqual({
            name: 'RoundTripName',
            formula: 'Sheet1!$A$1',
            localSheetId: -1,
          });
          expect(names).toContainEqual({
            name: 'LocalRoundTripName',
            formula: 'Sheet1!$B$1',
            localSheetId: 0,
          });
          reloaded.recalc();
          expect(reloaded.getValue({ sheet: 0, row: 0, col: 3 })).toEqual({
            kind: 'number',
            value: 4,
          });
          expect(reloaded.getValue({ sheet: 0, row: 0, col: 4 })).toEqual({
            kind: 'number',
            value: 7,
          });
        }
        if (first.capabilities.hyperlinks) {
          expect(reloaded.getHyperlinks(0)).toContainEqual({
            row: 1,
            col: 1,
            target: 'https://example.com/roundtrip',
            display: 'Round Trip',
            tooltip: 'Open round-trip link',
          });
        }
        if (first.capabilities.comments) {
          expect(reloaded.getComment(0, 2, 2)).toEqual({
            author: 'Formulon',
            text: 'Round-trip comment',
          });
          if (reloaded.capabilities.commentsEnumerable) {
            expect(reloaded.getComments(0)).toContainEqual({
              row: 4,
              col: 4,
              author: 'Formulon',
              text: 'Blank-cell comment',
            });
          }
        }
      } finally {
        reloaded.dispose();
      }
    } finally {
      first.dispose();
    }
  });

  it('saves and reloads static error values when supported', async () => {
    const first = await WorkbookHandle.createDefault();

    try {
      expect(first.isStub).toBe(false);
      if (!first.capabilities.staticErrorValues) return;

      first.setError({ sheet: 0, row: 0, col: 0 }, 1);
      first.setError({ sheet: 0, row: 1, col: 0 }, 5);

      const bytes = first.save();
      expect(bytes.length).toBeGreaterThan(0);

      const reloaded = await WorkbookHandle.loadBytes(bytes);
      try {
        expect(reloaded.isStub).toBe(false);
        expect(reloaded.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
          kind: 'error',
          code: 1,
          text: '#DIV/0!',
        });
        expect(reloaded.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({
          kind: 'error',
          code: 5,
          text: '#NUM!',
        });
      } finally {
        reloaded.dispose();
      }
    } finally {
      first.dispose();
    }
  });

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
      formats.set(addrKey({ sheet: 0, row: 23, col: 0 }), { phonetic: 'かんじ' });

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
        expect(roundTripped.get(addrKey({ sheet: 0, row: 23, col: 0 }))?.phonetic).toBe('かんじ');
      } finally {
        reloaded.dispose();
      }
    } finally {
      first.dispose();
    }
  });

  it('saves and reloads validation and sheet-protection metadata when supported', async () => {
    const first = await WorkbookHandle.createDefault();

    try {
      expect(first.isStub).toBe(false);

      if (first.capabilities.dataValidation) {
        expect(
          first.addValidationEntry(0, {
            type: 3,
            ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 }],
            formula1: '"Yes,No"',
            allowBlank: false,
            showErrorMessage: false,
            showDropDown: true,
          }),
        ).toBe(true);
      }
      if (first.capabilities.sheetProtectionRoundtrip) {
        expect(
          first.setSheetProtection(0, {
            enabled: true,
            legacyPassword: 'ABCD',
            sheet: true,
            selectLockedCells: true,
            selectUnlockedCells: true,
            sort: true,
            autoFilter: true,
          }),
        ).toBe(true);
      }

      const bytes = first.save();
      expect(bytes.length).toBeGreaterThan(0);

      const reloaded = await WorkbookHandle.loadBytes(bytes);
      try {
        expect(reloaded.isStub).toBe(false);

        if (first.capabilities.dataValidation) {
          const validations = reloaded.getValidationsForSheet(0);
          expect(validations).toHaveLength(1);
          expect(validations[0]).toMatchObject({
            type: 3,
            formula1: '"Yes,No"',
            allowBlank: false,
            showErrorMessage: false,
            ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 }],
          });
        }
        if (first.capabilities.sheetProtectionRoundtrip) {
          expect(reloaded.getSheetProtection(0)).toMatchObject({
            enabled: true,
            legacyPassword: 'ABCD',
            sheet: true,
            selectLockedCells: true,
            selectUnlockedCells: true,
            sort: true,
            autoFilter: true,
          });
        }
      } finally {
        reloaded.dispose();
      }
    } finally {
      first.dispose();
    }
  });

  it('saves and reloads data-validation hidden dropdown visibility', async () => {
    const first = await WorkbookHandle.createDefault();

    try {
      expect(first.isStub).toBe(false);
      if (!first.capabilities.dataValidation) return;

      expect(
        first.addValidationEntry(0, {
          type: 3,
          ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }],
          formula1: '"Yes,No"',
          showDropDown: true,
        }),
      ).toBe(true);

      const bytes = first.save();
      expect(bytes.length).toBeGreaterThan(0);

      const reloaded = await WorkbookHandle.loadBytes(bytes);
      try {
        expect(reloaded.isStub).toBe(false);
        const validations = reloaded.getValidationsForSheet(0);
        expect(validations).toHaveLength(1);
        const validation = validations[0];
        expect(validation).toBeDefined();
        expect(validation).toMatchObject({
          type: 3,
          formula1: '"Yes,No"',
          ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }],
        });
        expect(validation?.showDropDown).toBe(true);
      } finally {
        reloaded.dispose();
      }
    } finally {
      first.dispose();
    }
  });

  it('saves and reloads conditional-format visual payloads and dxfs when supported', async () => {
    const first = await WorkbookHandle.createDefault();

    try {
      expect(first.isStub).toBe(false);
      if (!first.capabilities.conditionalFormatMutate) return;

      first.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
      first.setNumber({ sheet: 0, row: 1, col: 0 }, 2);
      first.setNumber({ sheet: 0, row: 2, col: 0 }, 3);

      let dxfId: number | undefined;
      if (first.capabilities.conditionalFormatDxf) {
        dxfId = first.addDxf({
          fill: { pattern: 1, fgArgb: 0xffe2f0d9, bgArgb: 0 },
          font: {
            name: 'Calibri',
            size: 11,
            bold: true,
            italic: false,
            strike: false,
            underline: 0,
            vertAlign: 0,
            colorArgb: 0xff006100,
          },
        });
        expect(dxfId).toBeGreaterThanOrEqual(0);
      }

      expect(
        first.addConditionalFormat(0, {
          sqref: [{ firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 0 }],
          type: 1,
          op: 5,
          formula1: '1',
          ...(dxfId !== undefined ? { dxfId } : {}),
        }),
      ).toBeGreaterThanOrEqual(0);
      if (first.capabilities.conditionalFormatVisualMutate) {
        expect(
          first.addConditionalFormat(0, {
            sqref: [{ firstRow: 0, firstCol: 1, lastRow: 2, lastCol: 1 }],
            type: 2,
            colorScale: {
              thresholds: [{ type: 3 }, { type: 4 }],
              colors: [
                { a: 255, r: 255, g: 0, b: 0 },
                { a: 255, r: 0, g: 128, b: 0 },
              ],
            },
          }),
        ).toBeGreaterThanOrEqual(0);
      }

      const bytes = first.save();
      expect(bytes.length).toBeGreaterThan(0);

      const reloaded = await WorkbookHandle.loadBytes(bytes);
      try {
        expect(reloaded.isStub).toBe(false);
        const formats = reloaded.getConditionalFormats(0);
        expect(formats.some((entry) => entry.type === 1 && entry.formula1 === '1')).toBe(true);
        if (first.capabilities.conditionalFormatDxf) {
          const dxfRule = formats.find((entry) => entry.type === 1 && entry.dxfId !== undefined);
          expect(dxfRule?.dxfId).toBeGreaterThanOrEqual(0);
          expect(reloaded.getDxf(dxfRule?.dxfId ?? -1)).toMatchObject({
            fill: { pattern: 1, fgArgb: 0xffe2f0d9 },
            font: { bold: true, colorArgb: 0xff006100 },
          });
        }
        if (first.capabilities.conditionalFormatVisualMutate) {
          expect(
            formats.some(
              (entry) =>
                entry.type === 2 &&
                entry.colorScale?.colors.length === 2 &&
                entry.colorScale.thresholds.length === 2,
            ),
          ).toBe(true);
        }
      } finally {
        reloaded.dispose();
      }
    } finally {
      first.dispose();
    }
  });

  it('round-trips every classic conditional-format preset dxf through the real engine', async () => {
    const first = await WorkbookHandle.createDefault();
    try {
      expect(first.isStub).toBe(false);
      if (!first.capabilities.conditionalFormatMutate || !first.capabilities.conditionalFormatDxf)
        return;

      const presets = [
        'red-fill',
        'yellow-fill',
        'green-fill',
        'light-red-fill',
        'red-text',
        'red-border',
      ] as const;
      const rules: ConditionalRule[] = presets.map((preset, row) => ({
        kind: 'cell-value',
        range: { sheet: 0, r0: row, c0: 0, r1: row, c1: 0 },
        op: '>',
        a: row,
        apply: formatPresetPatch(preset),
      }));
      rules.push({
        kind: 'cell-value',
        range: { sheet: 0, r0: presets.length, c0: 0, r1: presets.length, c1: 0 },
        op: '>',
        a: presets.length,
        apply: { numFmt: { kind: 'percent', decimals: 0 }, bold: true },
      });

      expect(syncConditionalRulesToEngine(first, rules, 0)).toEqual({ written: 7, skipped: 0 });
      const reloaded = await WorkbookHandle.loadBytes(first.save());
      try {
        const formats = reloaded.getConditionalFormats(0);
        expect(formats).toHaveLength(7);
        const dxfs = formats.map((format) => reloaded.getDxf(format.dxfId ?? -1));
        expect(dxfs[0]).toMatchObject({
          fill: { fgArgb: 0xffffc7ce },
          font: { colorArgb: 0xff9c0006 },
        });
        expect(dxfs[1]).toMatchObject({
          fill: { fgArgb: 0xffffeb9c },
          font: { colorArgb: 0xff9c6500 },
        });
        expect(dxfs[2]).toMatchObject({
          fill: { fgArgb: 0xffc6efce },
          font: { colorArgb: 0xff006100 },
        });
        expect(dxfs[3]).toMatchObject({ fill: { fgArgb: 0xffffc7ce } });
        expect(dxfs[4]).toMatchObject({ font: { colorArgb: 0xffc00000 } });
        expect(dxfs[5]).toMatchObject({ border: { top: { style: 1, colorArgb: 0xffff0000 } } });
        expect(dxfs[6]).toMatchObject({ numFmt: { formatCode: '0%' }, font: { bold: true } });
      } finally {
        reloaded.dispose();
      }
    } finally {
      first.dispose();
    }
  });

  it('saves and reloads PivotTable source metadata and report layout when supported', async () => {
    const first = await WorkbookHandle.createDefault();

    try {
      expect(first.isStub).toBe(false);
      if (
        !first.capabilities.pivotTableMutate ||
        !first.capabilities.pivotCacheSource ||
        !first.capabilities.pivotReportLayout
      ) {
        return;
      }

      first.setText({ sheet: 0, row: 0, col: 0 }, 'Region');
      first.setText({ sheet: 0, row: 0, col: 1 }, 'Sales');
      first.setText({ sheet: 0, row: 1, col: 0 }, 'East');
      first.setNumber({ sheet: 0, row: 1, col: 1 }, 12);
      first.setText({ sheet: 0, row: 2, col: 0 }, 'West');
      first.setNumber({ sheet: 0, row: 2, col: 1 }, 8);

      const created = createPivotTableFromRange(first, {
        source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 },
        destination: { sheet: 0, row: 5, col: 0 },
        name: 'SalesPivot',
        rowField: 'Region',
        valueField: 'Sales',
        aggregation: PivotAggregation.Sum,
      });
      expect(created).toMatchObject({ ok: true });
      if (!created.ok) return;
      expect(first.setPivotReportLayout(0, created.pivotIndex, PivotReportLayout.Tabular)).toBe(
        true,
      );

      const bytes = first.save();
      expect(bytes.length).toBeGreaterThan(0);

      const reloaded = await WorkbookHandle.loadBytes(bytes);
      try {
        expect(reloaded.isStub).toBe(false);
        const pivots = reloaded.getPivotTables();
        expect(pivots).toHaveLength(1);
        expect(pivots[0]).toMatchObject({
          sheetIndex: 0,
          pivotIndex: 0,
          top: 5,
          left: 0,
        });
        const cacheId = reloaded.pivotCacheIds()[0] ?? -1;
        expect(cacheId).toBeGreaterThanOrEqual(0);
        expect(reloaded.getPivotCacheWorksheetSource(cacheId)).toMatchObject({
          present: true,
          ref: 'A1:B3',
          sheet: 'Sheet1',
        });
        expect(reloaded.getPivotReportLayout(0, 0)).toBe(PivotReportLayout.Tabular);
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

  it('writes Format as Table as a real ListObject and reads it back as an overlay', async () => {
    const first = await WorkbookHandle.createDefault();

    try {
      if (!first.capabilities.tableMutate) return;
      const range: Range = { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 };
      seedTableCells(first);

      const store = createSpreadsheetStore();
      const overlay = formatAsTable(store, range, { workbook: first, style: 'medium' });
      expect(overlay).not.toBeNull();

      const reloaded = await WorkbookHandle.loadBytes(first.save());
      try {
        const tables = reloaded.getTables();
        expect(tables).toHaveLength(1);
        expect(tables[0]).toMatchObject({ sheetIndex: 0, ref: 'A1:B3' });

        const overlays = tableOverlaysFromEngine(reloaded);
        expect(overlays).toHaveLength(1);
        expect(overlays[0]?.source).toBe('engine');
        expect(overlays[0]?.range).toEqual(range);
      } finally {
        reloaded.dispose();
      }
    } finally {
      first.dispose();
    }
  });

  it('keeps a table across a load and save that never touches it', async () => {
    const authored = await WorkbookHandle.createDefault();
    let bytes: Uint8Array;
    try {
      if (!authored.capabilities.tableMutate) return;
      seedTableCells(authored);
      formatAsTable(
        createSpreadsheetStore(),
        { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 },
        {
          workbook: authored,
        },
      );
      bytes = authored.save();
    } finally {
      authored.dispose();
    }

    const passthrough = await WorkbookHandle.loadBytes(bytes);
    let resaved: Uint8Array;
    try {
      expect(passthrough.getTables()).toHaveLength(1);
      resaved = passthrough.save();
    } finally {
      passthrough.dispose();
    }

    const reloaded = await WorkbookHandle.loadBytes(resaved);
    try {
      expect(reloaded.getTables()).toHaveLength(1);
      expect(reloaded.getTables()[0]).toMatchObject({ sheetIndex: 0, ref: 'A1:B3' });
    } finally {
      reloaded.dispose();
    }
  });

  it('round-trips an AutoFilter definition and rehydrates its range', async () => {
    const first = await WorkbookHandle.createDefault();

    try {
      if (!first.capabilities.autoFilter) return;
      const range: Range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 };
      const store = createSpreadsheetStore();
      seedStoreText(store, 0, 0, 'Region');
      seedStoreText(store, 1, 0, 'East');
      seedStoreText(store, 2, 0, 'West');
      seedStoreText(store, 3, 0, 'North');
      applyValueFilter(store.getState(), store, range, 0, ['West']);

      expect(syncAutoFilterToEngine(first, store.getState(), 0)).toBe(true);

      const reloaded = await WorkbookHandle.loadBytes(first.save());
      try {
        const xml = reloaded.getSheetAutoFilterXml(0) ?? '';
        expect(xml).toContain('ref="A1:B4"');
        expect(xml).toContain('colId="0"');
        expect(xml).toContain('<filter val="East"/>');
        expect(xml).not.toContain('<filter val="West"/>');

        const restored = createSpreadsheetStore();
        hydrateAutoFilterFromEngine(reloaded, restored, 0);
        expect(restored.getState().ui.filterRange).toEqual(range);
      } finally {
        reloaded.dispose();
      }
    } finally {
      first.dispose();
    }
  });

  it('clears a saved AutoFilter when the user removes the filter', async () => {
    const first = await WorkbookHandle.createDefault();

    try {
      if (!first.capabilities.autoFilter) return;
      const range: Range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 };
      const store = createSpreadsheetStore();
      seedStoreText(store, 1, 0, 'East');
      applyValueFilter(store.getState(), store, range, 0, []);
      syncAutoFilterToEngine(first, store.getState(), 0);

      clearFilter(store.getState(), store);
      expect(syncAutoFilterToEngine(first, store.getState(), 0)).toBe(true);

      const reloaded = await WorkbookHandle.loadBytes(first.save());
      try {
        expect(reloaded.getSheetAutoFilterXml(0) ?? '').toBe('');
        const restored = createSpreadsheetStore();
        hydrateAutoFilterFromEngine(reloaded, restored, 0);
        expect(restored.getState().ui.filterRange).toBeNull();
      } finally {
        reloaded.dispose();
      }
    } finally {
      first.dispose();
    }
  });
});
