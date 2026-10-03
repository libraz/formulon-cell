import { describe, expect, it } from 'vitest';
import {
  evaluateCfFromEngine,
  hydrateConditionalRulesFromEngine,
} from '../../../../src/engine/cf-sync.js';
import {
  conditionalRuleToEngineInput,
  syncConditionalRulesToEngine,
} from '../../../../src/engine/cf-writeback.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { formatPresetPatch } from '../../../../src/interact/conditional-dialog-spec.js';
import { type ConditionalRule, createSpreadsheetStore } from '../../../../src/store/store.js';

import { canLoadWasm } from './fixtures.js';

describe.skipIf(!canLoadWasm())('real xlsx round-trip', () => {
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

  it('preserves 0.12 data-bar and icon-set metadata through store hydration and writeback', async () => {
    const first = await WorkbookHandle.createDefault();
    try {
      expect(first.isStub).toBe(false);
      expect(first.capabilities.conditionalFormatVisualMutate).toBe(true);

      expect(
        first.addConditionalFormat(0, {
          sqref: [{ firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 0 }],
          type: 3,
          dataBar: {
            min: { type: 3 },
            max: { type: 4 },
            fill: { a: 255, r: 0, g: 120, b: 212 },
            showValue: true,
            minLengthPct: 0,
            maxLengthPct: 100,
            gradient: false,
            direction: 2,
          },
        }),
      ).toBeGreaterThanOrEqual(0);
      first.setNumber({ sheet: 0, row: 0, col: 0 }, -50);
      first.setNumber({ sheet: 0, row: 1, col: 0 }, 100);
      const rawBars = first
        .evaluateCfRange(0, 0, 0, 1, 0)
        .flatMap((cell) =>
          cell.matches
            .filter((match) => match.kind === 2)
            .map((match) => ({ col: cell.col, match })),
        );
      expect(rawBars).toHaveLength(2);
      for (const { match } of rawBars) {
        expect(match.barLengthPct).toBeCloseTo(100);
        expect(match.barAxisPositionPct).toBeCloseTo(33.333333, 3);
      }
      const engineOverlay = evaluateCfFromEngine(first, 0, 0, 0, 1, 0);
      expect(engineOverlay.get('0:0:0')?.bar).toBeCloseTo(1 / 3);
      expect(engineOverlay.get('0:0:0')?.barAxis).toBeCloseTo(2 / 3);
      expect(engineOverlay.get('0:0:0')?.barDirection).toBe('right');
      expect(engineOverlay.get('0:1:0')?.bar).toBeCloseTo(2 / 3);
      expect(engineOverlay.get('0:1:0')?.barAxis).toBeCloseTo(2 / 3);
      expect(engineOverlay.get('0:1:0')?.barDirection).toBe('left');
      expect(
        first.addConditionalFormat(0, {
          sqref: [{ firstRow: 0, firstCol: 1, lastRow: 2, lastCol: 1 }],
          type: 4,
          iconSet: {
            name: 3,
            thresholds: [
              { type: 1, value: '33' },
              { type: 1, value: '67' },
            ],
            reverse: false,
            showValue: true,
            percent: true,
            floor: { type: 1, value: '10' },
          },
        }),
      ).toBeGreaterThanOrEqual(0);

      const loaded = await WorkbookHandle.loadBytes(first.save());
      try {
        const store = createSpreadsheetStore();
        hydrateConditionalRulesFromEngine(loaded, store, 0);
        const imported = store
          .getState()
          .conditional.rules.filter(
            (rule): rule is Extract<ConditionalRule, { kind: 'data-bar' | 'icon-set' }> =>
              rule.kind === 'data-bar' || rule.kind === 'icon-set',
          );
        expect(imported).toEqual([
          expect.objectContaining({
            kind: 'data-bar',
            gradient: false,
            direction: 'right-to-left',
          }),
          expect.objectContaining({
            kind: 'icon-set',
            floor: { kind: 'percent', value: 10, gte: true },
            thresholds: [
              { kind: 'percent', value: 33, gte: true },
              { kind: 'percent', value: 67, gte: true },
            ],
          }),
        ]);

        const written = await WorkbookHandle.createDefault();
        try {
          for (const rule of imported) {
            const input = conditionalRuleToEngineInput({ ...rule, engineId: undefined });
            if (!input) throw new Error(`Expected a writable ${rule.kind} rule.`);
            expect(written.addConditionalFormat(0, input)).toBeGreaterThanOrEqual(0);
          }
          const formats = written.getConditionalFormats(0);
          expect(formats.find((entry) => entry.type === 3)?.dataBar).toMatchObject({
            gradient: false,
            direction: 2,
          });
          expect(formats.find((entry) => entry.type === 4)?.iconSet).toMatchObject({
            thresholds: [
              { type: 1, value: '33' },
              { type: 1, value: '67' },
            ],
            floor: { type: 1, value: '10' },
          });
        } finally {
          written.dispose();
        }
      } finally {
        loaded.dispose();
      }
    } finally {
      first.dispose();
    }
  });

  it('preserves strict icon-set floor and threshold comparisons through the real engine', async () => {
    const first = await WorkbookHandle.createDefault();
    const iconCells = (wb: WorkbookHandle) =>
      wb
        .evaluateCfRange(0, 0, 0, 0, 2)
        .flatMap((cell) =>
          cell.matches
            .filter((match) => match.kind === 3)
            .map((match) => ({ col: cell.col, iconIndex: match.iconIndex })),
        );
    try {
      expect(first.isStub).toBe(false);
      expect(first.capabilities.conditionalFormatVisualMutate).toBe(true);
      first.setNumber({ sheet: 0, row: 0, col: 0 }, 14);
      first.setNumber({ sheet: 0, row: 0, col: 1 }, 15);
      first.setNumber({ sheet: 0, row: 0, col: 2 }, 20);
      expect(
        first.addConditionalFormat(0, {
          sqref: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 2 }],
          type: 4,
          iconSet: {
            name: 3,
            thresholds: [
              { type: 0, value: '20', gte: false },
              { type: 0, value: '40', gte: false },
            ],
            reverse: false,
            showValue: true,
            percent: true,
            floor: { type: 0, value: '15', gte: false },
          },
        }),
      ).toBeGreaterThanOrEqual(0);
      expect(iconCells(first)).toEqual([{ col: 2, iconIndex: 0 }]);

      const loaded = await WorkbookHandle.loadBytes(first.save());
      try {
        expect(iconCells(loaded)).toEqual([{ col: 2, iconIndex: 0 }]);
        const store = createSpreadsheetStore();
        hydrateConditionalRulesFromEngine(loaded, store, 0);
        const imported = store
          .getState()
          .conditional.rules.find(
            (rule): rule is Extract<ConditionalRule, { kind: 'icon-set' }> =>
              rule.kind === 'icon-set',
          );
        expect(imported).toMatchObject({
          floor: { kind: 'number', value: 15, gte: false },
          thresholds: [
            { kind: 'number', value: 20, gte: false },
            { kind: 'number', value: 40, gte: false },
          ],
        });
        if (!imported) throw new Error('Expected an imported icon-set rule.');
        const input = conditionalRuleToEngineInput({ ...imported, engineId: undefined });
        expect(input?.iconSet).toMatchObject({
          thresholds: [
            { type: 0, value: '20', gte: false },
            { type: 0, value: '40', gte: false },
          ],
          floor: { type: 0, value: '15', gte: false },
        });

        const written = await WorkbookHandle.createDefault();
        try {
          written.setNumber({ sheet: 0, row: 0, col: 0 }, 14);
          written.setNumber({ sheet: 0, row: 0, col: 1 }, 15);
          written.setNumber({ sheet: 0, row: 0, col: 2 }, 20);
          if (!input) throw new Error('Expected a writable icon-set rule.');
          expect(written.addConditionalFormat(0, input)).toBeGreaterThanOrEqual(0);
          expect(
            written.getConditionalFormats(0).find((entry) => entry.type === 4)?.iconSet,
          ).toMatchObject({
            thresholds: [
              { type: 0, value: '20', gte: false },
              { type: 0, value: '40', gte: false },
            ],
            floor: { type: 0, value: '15', gte: false },
          });
          expect(iconCells(written)).toEqual([{ col: 2, iconIndex: 0 }]);
        } finally {
          written.dispose();
        }
      } finally {
        loaded.dispose();
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
});
