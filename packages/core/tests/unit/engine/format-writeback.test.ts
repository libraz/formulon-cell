import { describe, expect, it } from 'vitest';
import {
  argbToCssColor,
  borderRecordFromFormat,
  borderRecordToFormat,
  buildXfRecord,
  cssColorToArgb,
  fillRecordFromFormat,
  fillRecordToFormat,
  fontRecordFromFormat,
  fontRecordToFormat,
  formatCodeToNumFmt,
  numFmtToFormatCode,
} from '../../../src/engine/format-writeback.js';

describe('cssColorToArgb / argbToCssColor', () => {
  it('parses #rrggbb', () => {
    expect(cssColorToArgb('#FF0000')).toBe(0xffff0000);
    expect(cssColorToArgb('#00ff00')).toBe(0xff00ff00);
  });

  it('parses #rgb shorthand', () => {
    expect(cssColorToArgb('#f00')).toBe(0xffff0000);
    expect(cssColorToArgb('#0f0')).toBe(0xff00ff00);
  });

  it('parses #rrggbbaa', () => {
    expect(cssColorToArgb('#ff000080')).toBe(0x80ff0000);
  });

  it('parses rgb()/rgba()', () => {
    expect(cssColorToArgb('rgb(255, 0, 0)')).toBe(0xffff0000);
    expect(cssColorToArgb('rgba(0, 255, 0, 0.5)')).toBe(0x8000ff00);
  });

  it('parses named colors', () => {
    expect(cssColorToArgb('red')).toBe(0xffff0000);
    expect(cssColorToArgb('white')).toBe(0xffffffff);
    expect(cssColorToArgb('black')).toBe(0xff000000);
  });

  it('returns null for unknown input', () => {
    expect(cssColorToArgb('not-a-color')).toBeNull();
    expect(cssColorToArgb('')).toBeNull();
  });

  it('round-trips opaque colors as #rrggbb', () => {
    expect(argbToCssColor(0xffff0000)).toBe('#ff0000');
    expect(argbToCssColor(0xff00ff00)).toBe('#00ff00');
    expect(argbToCssColor(0xff000000)).toBe('#000000');
  });

  it('emits rgba() for non-opaque colors', () => {
    expect(argbToCssColor(0x80ff0000)).toMatch(/rgba\(255, 0, 0, 0\.50\)/);
  });
});

describe('fontRecordFromFormat / fontRecordToFormat', () => {
  it('emits defaults when CellFormat is empty', () => {
    const rec = fontRecordFromFormat({});
    expect(rec.name).toBe('Calibri');
    expect(rec.size).toBe(11);
    expect(rec.bold).toBe(false);
    expect(rec.italic).toBe(false);
    expect(rec.underline).toBe(0);
    expect(rec.colorArgb).toBe(0xff000000);
  });

  it('encodes bold/italic/underline/strike + custom font', () => {
    const rec = fontRecordFromFormat({
      bold: true,
      italic: true,
      underline: true,
      strike: true,
      fontFamily: 'Arial',
      fontSize: 14,
      color: '#ff0000',
    });
    expect(rec.bold).toBe(true);
    expect(rec.italic).toBe(true);
    expect(rec.underline).toBe(1);
    expect(rec.strike).toBe(true);
    expect(rec.name).toBe('Arial');
    expect(rec.size).toBe(14);
    expect(rec.colorArgb).toBe(0xffff0000);
  });

  it('round-trips font fields back to CellFormat', () => {
    const rec = fontRecordFromFormat({ bold: true, fontSize: 16, color: '#0000ff' });
    const fmt = fontRecordToFormat(rec);
    expect(fmt.bold).toBe(true);
    expect(fmt.fontSize).toBe(16);
    expect(fmt.color).toBe('#0000ff');
  });

  it('hydrate strips workbook-default font name+size', () => {
    const fmt = fontRecordToFormat({
      name: 'Calibri',
      size: 11,
      bold: false,
      italic: false,
      strike: false,
      underline: 0,
      colorArgb: 0xff000000,
    });
    expect(fmt.fontFamily).toBeUndefined();
    expect(fmt.fontSize).toBeUndefined();
  });
});

describe('fillRecordFromFormat / fillRecordToFormat', () => {
  it('uses pattern=0 when no fill', () => {
    expect(fillRecordFromFormat({}).pattern).toBe(0);
  });

  it('uses pattern=1 (solid) when fill present', () => {
    const rec = fillRecordFromFormat({ fill: '#ffff00' });
    expect(rec.pattern).toBe(1);
    expect(rec.fgArgb).toBe(0xffffff00);
  });

  it('round-trips back to CellFormat', () => {
    const rec = fillRecordFromFormat({ fill: '#abcdef' });
    expect(fillRecordToFormat(rec)).toEqual({ fill: '#abcdef' });
  });

  it('hydrate produces empty fmt when pattern=0', () => {
    expect(fillRecordToFormat({ pattern: 0, fgArgb: 0, bgArgb: 0 })).toEqual({});
  });
});

describe('borderRecordFromFormat / borderRecordToFormat', () => {
  it('encodes thin border on top side', () => {
    const rec = borderRecordFromFormat({ borders: { top: true } });
    expect(rec.top.style).toBe(1); // thin
    expect(rec.left.style).toBe(0);
  });

  it('encodes styled+colored border', () => {
    const rec = borderRecordFromFormat({
      borders: { left: { style: 'medium', color: '#ff0000' } },
    });
    expect(rec.left.style).toBe(2);
    expect(rec.left.colorArgb).toBe(0xffff0000);
  });

  it('round-trips top side back', () => {
    const rec = borderRecordFromFormat({ borders: { top: true } });
    const fmt = borderRecordToFormat(rec);
    expect(fmt.borders?.top).toEqual({ style: 'thin' });
  });
});

describe('numFmtToFormatCode / formatCodeToNumFmt', () => {
  it('general → null', () => {
    expect(numFmtToFormatCode({ kind: 'general' })).toBeNull();
    expect(formatCodeToNumFmt('General')).toBeNull();
  });

  it('fixed with thousands round-trip', () => {
    const code = numFmtToFormatCode({ kind: 'fixed', decimals: 2, thousands: true });
    expect(code).toBe('#,##0.00');
    expect(formatCodeToNumFmt(code ?? '')).toEqual({
      kind: 'fixed',
      decimals: 2,
      thousands: true,
    });
  });

  it('percent round-trip', () => {
    expect(numFmtToFormatCode({ kind: 'percent', decimals: 2 })).toBe('0.00%');
    expect(formatCodeToNumFmt('0.00%')).toEqual({ kind: 'percent', decimals: 2 });
  });

  it('scientific round-trip', () => {
    expect(numFmtToFormatCode({ kind: 'scientific', decimals: 3 })).toBe('0.000E+00');
    expect(formatCodeToNumFmt('0.000E+00')).toEqual({ kind: 'scientific', decimals: 3 });
  });

  it('text → "@"', () => {
    expect(numFmtToFormatCode({ kind: 'text' })).toBe('@');
    expect(formatCodeToNumFmt('@')).toEqual({ kind: 'text' });
  });

  it('date pattern passes through', () => {
    expect(numFmtToFormatCode({ kind: 'date', pattern: 'yyyy-mm-dd' })).toBe('yyyy-mm-dd');
    expect(formatCodeToNumFmt('yyyy-mm-dd')).toEqual({ kind: 'date', pattern: 'yyyy-mm-dd' });
  });

  it('unknown patterns surface as custom', () => {
    expect(formatCodeToNumFmt('???xx???')).toEqual({ kind: 'custom', pattern: '???xx???' });
  });
});

describe('buildXfRecord', () => {
  it('packs alignment ordinals + wrap flag', () => {
    const xf = buildXfRecord(1, 2, 3, 4, { align: 'center', vAlign: 'middle', wrap: true });
    expect(xf).toEqual({
      fontIndex: 1,
      fillIndex: 2,
      borderIndex: 3,
      numFmtId: 4,
      horizontalAlign: 2,
      verticalAlign: 1,
      wrapText: true,
    });
  });

  it('defaults map to general/bottom/no-wrap', () => {
    const xf = buildXfRecord(0, 0, 0, 0, {});
    expect(xf.horizontalAlign).toBe(0);
    expect(xf.verticalAlign).toBe(2);
    expect(xf.wrapText).toBe(false);
  });
});
