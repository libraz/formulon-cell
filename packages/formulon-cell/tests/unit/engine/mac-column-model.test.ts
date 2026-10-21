import { describe, expect, it } from 'vitest';
import {
  macDefaultColumnPt,
  macMdw,
  macPtToRaw,
  macRawToPt,
} from '../../../src/engine/mac-column-model.js';

// Raw widths Mac Excel saved for 8.43 / 10 / 20 displayed characters in a
// 游ゴシック 12 book, re-read under each Normal font (xlwings measurements).
const SAVED_RAW = [9.140625, 10.7109375, 20.7109375] as const;

const MEASURED: ReadonlyArray<readonly [string, number, readonly number[], number]> = [
  ['Calibri', 11, [55, 64, 124], 65],
  ['Arial', 10, [55, 64, 124], 65],
  ['ＭＳ Ｐゴシック', 11, [55, 64, 124], 65],
  ['Aptos Narrow', 11, [55, 64, 124], 65],
  ['Aptos Narrow', 12, [55, 64, 124], 65],
  ['游ゴシック', 11, [55, 64, 124], 65],
  ['Meiryo', 11, [64, 75, 145], 75],
  ['Helvetica Neue', 12, [64, 75, 145], 75],
  ['Calibri', 12, [55, 64, 124], 65],
  ['Arial', 12, [64, 75, 145], 75],
  ['游ゴシック', 12, [64, 75, 145], 75],
];

describe('engine/mac-column-model', () => {
  it.each(MEASURED)('reproduces the measured %s %d pt columns', (font, size, widths, base) => {
    const mdw = macMdw(font, size);
    expect(SAVED_RAW.map((raw) => macRawToPt(raw, mdw))).toEqual(widths);
    expect(macDefaultColumnPt(10, mdw)).toBe(base);
  });

  it.each([
    ['等线', 12, 6],
    ['맑은 고딕', 12, 7],
    ['Tahoma', 12, 7],
  ])('derives the locale default %s %d pt digit width', (font, size, mdw) => {
    expect(macMdw(font, size)).toBe(mdw);
  });

  it('writes back raw widths the way Mac Excel saves them', () => {
    expect(macPtToRaw(64, 7)).toBe(9.140625);
    expect(macPtToRaw(75, 7)).toBe(10.7109375);
    expect(macPtToRaw(145, 7)).toBe(20.7109375);
    for (const mdw of [6, 7]) {
      for (const pt of [1, 21, 55, 64, 75, 84, 96, 124, 145, 300]) {
        const raw = macPtToRaw(pt, mdw);
        expect(raw).toBe(Math.trunc((pt / mdw) * 256) / 256);
        expect(macRawToPt(raw, mdw)).toBe(pt);
      }
    }
  });

  it('rounds a fractional width to whole points before writing back', () => {
    expect(macPtToRaw(74.6, 7)).toBe(macPtToRaw(75, 7));
    expect(macPtToRaw(0, 7)).toBe(0);
    expect(macPtToRaw(-3, 7)).toBe(0);
  });

  it('resolves aliases through NFKC and case folding', () => {
    expect(macMdw('メイリオ', 11)).toBe(7);
    expect(macMdw('meiryo', 11)).toBe(7);
    expect(macMdw('DengXian', 12)).toBe(6);
    expect(macMdw('Yu Gothic', 12)).toBe(7);
    expect(macMdw('Malgun Gothic', 12)).toBe(7);
    expect(macMdw('MS PGothic', 11)).toBe(6);
    expect(macMdw('MS Pｺﾞｼｯｸ', 11)).toBe(6);
    expect(macMdw('ms pゴシック', 11)).toBe(6);
    expect(macMdw('CALIBRI', 12)).toBe(6);
  });

  it('uses the unknown-family em for fonts outside the table', () => {
    expect(macMdw('Menlo', 12)).toBe(7);
    expect(macMdw('Menlo', 10)).toBe(6);
  });
});
