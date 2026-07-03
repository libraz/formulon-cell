import { describe, expect, it } from 'vitest';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import { PivotAggregation, PivotReportLayout } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const canLoadWasm = (): boolean =>
  typeof WebAssembly !== 'undefined' && typeof SharedArrayBuffer !== 'undefined';

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
});
