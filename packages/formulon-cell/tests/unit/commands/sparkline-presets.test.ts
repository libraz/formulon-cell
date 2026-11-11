import type { Color } from '@libraz/formulon';
import { ColorKind, SparklineType } from '@libraz/formulon';
import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import {
  defaultSparklineGroupSettings,
  groupOfCell,
  setSparklines,
} from '../../../src/commands/sparkline.js';
import {
  applySparklineStylePresetAction,
  SPARKLINE_STYLE_PRESETS,
  sparklineStylePresetIndex,
} from '../../../src/commands/sparkline-presets.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';
import measured from '../../fixtures/excel-sparkline-style-presets.json' with { type: 'json' };

type MeasuredColor = { theme: string; tint?: string } | { rgb: string };

const PARTS = [
  ['Series', 'series'],
  ['Negative', 'negative'],
  ['Axis', 'axis'],
  ['Markers', 'markers'],
  ['First', 'first'],
  ['Last', 'last'],
  ['High', 'high'],
  ['Low', 'low'],
] as const;

const expected = (m: MeasuredColor): Pick<Color, 'kind' | 'argb' | 'theme' | 'tint'> =>
  'rgb' in m
    ? { kind: ColorKind.Rgb, argb: Number.parseInt(m.rgb, 16), theme: 0, tint: 0 }
    : { kind: ColorKind.Theme, argb: 0, theme: Number(m.theme), tint: Number(m.tint ?? 0) };

const at = (row: number) => ({ sheet: 0, row, col: 5 });

const setup = () => {
  const store = createSpreadsheetStore();
  const history = new History();
  setSparklines(
    store,
    [0, 1].map((row) => ({ addr: at(row), source: `A${row + 1}:C${row + 1}` })),
    defaultSparklineGroupSettings(SparklineType.Line),
    history,
  );
  mutators.setActive(store, at(0));
  mutators.setRange(store, { sheet: 0, r0: 0, c0: 5, r1: 1, c1: 5 });
  const ctx = { store, history, origin: 'ribbon' as const, commandId: 'sparklineStyleGallery' };
  return { store, history, ctx, settings: () => groupOfCell(store.getState(), at(0))?.settings };
};

describe('sparkline style presets', () => {
  it('holds the 36 measured presets in gallery order', () => {
    expect(SPARKLINE_STYLE_PRESETS).toHaveLength(36);
    expect(measured.map((row) => row.k)).toEqual(SPARKLINE_STYLE_PRESETS.map((_, k) => k));
  });

  for (const row of measured) {
    it(`preset ${row.k + 1} carries the measured colours`, () => {
      const preset = SPARKLINE_STYLE_PRESETS[row.k];
      for (const [from, part] of PARTS) {
        const color = preset?.[part];
        expect(
          color && { kind: color.kind, argb: color.argb, theme: color.theme, tint: color.tint },
        ).toEqual(expected((row.colors as Record<string, MeasuredColor>)[from] as MeasuredColor));
      }
    });
  }

  it('applies every colour of a preset to the selected groups as one undo step', () => {
    const { ctx, settings, history } = setup();
    expect(sparklineStylePresetIndex(settings())).toBe(-1);
    expect(applySparklineStylePresetAction(ctx, 30)).toBe(true);
    const preset = SPARKLINE_STYLE_PRESETS[30];
    expect(settings()?.colorSeries).toEqual(preset?.series);
    expect(settings()?.colorLow).toEqual(preset?.low);
    expect(settings()?.colorAxis).toEqual(preset?.axis);
    expect(sparklineStylePresetIndex(settings())).toBe(30);
    history.undo();
    expect(sparklineStylePresetIndex(settings())).toBe(-1);
  });

  it('is refused on a protected sheet', () => {
    const { ctx, store, settings } = setup();
    setProtectedSheet(store, 0, true);
    expect(applySparklineStylePresetAction(ctx, 0)).toBe(false);
    expect(sparklineStylePresetIndex(settings())).toBe(-1);
  });
});
