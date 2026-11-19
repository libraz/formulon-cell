import { mount, type SpreadsheetInstance } from '@libraz/formulon-cell';
import { afterEach, describe, expect, it } from 'vitest';
import { spreadsheetInternals } from '../../src/mount/internals.js';
import { makeHost } from './helpers.js';

const live: Array<{ dispose(): void }> = [];

const mountSheet = async (): Promise<SpreadsheetInstance> => {
  const instance = await mount(makeHost(), { locale: 'en' });
  live.push(instance);
  return instance;
};

afterEach(() => {
  for (const d of live.splice(0).reverse()) d.dispose();
});

const formatAt = (instance: SpreadsheetInstance, row: number, col: number) =>
  spreadsheetInternals(instance).store.getState().format.formats.get(`0:${row}:${col}`);

const THEME = [
  '#111111',
  '#eeeeee',
  '#222222',
  '#dddddd',
  '#ff0000',
  '#00ff00',
  '#0000ff',
  '#ffff00',
  '#ff00ff',
  '#00ffff',
  '#123456',
  '#654321',
];

describe('facade journal format', () => {
  it('round-trips updateCellFormat', async () => {
    const inst = await mountSheet();
    const s = inst.workbook.sheet(0);
    s.updateCellFormat('A1', { font: { bold: true } });
    expect(s.ooxml.getCellXfIndex('A1')).not.toBeNull();
    expect(formatAt(inst, 0, 0)?.bold).toBe(true);
    inst.history.undo();
    expect(s.ooxml.getCellXfIndex('A1')).toBeNull();
    expect(inst.history.canUndo).toBe(false);
    inst.history.redo();
    expect(s.ooxml.getCellXfIndex('A1')).not.toBeNull();
    expect(formatAt(inst, 0, 0)?.bold).toBe(true);
  });

  it('clears the store format when updateCellFormat is undone', async () => {
    const inst = await mountSheet();
    const s = inst.workbook.sheet(0);
    s.updateCellFormat('A1', { font: { bold: true } });
    inst.history.undo();
    expect(formatAt(inst, 0, 0)?.bold).toBeFalsy();
  });

  it('round-trips setHyperlink with the link cell format', async () => {
    const inst = await mountSheet();
    const s = inst.workbook.sheet(0);
    s.setHyperlink({ range: 'B2', target: 'https://example.com/' });
    expect(s.listHyperlinks()).toHaveLength(1);
    expect(formatAt(inst, 1, 1)?.hyperlink).toBeDefined();
    inst.history.undo();
    expect(s.listHyperlinks()).toHaveLength(0);
    expect(s.ooxml.getCellXfIndex('B2') ?? 0).toBe(0);
    expect(formatAt(inst, 1, 1)?.hyperlink).toBeUndefined();
    inst.history.redo();
    expect(s.listHyperlinks()).toHaveLength(1);
    expect(formatAt(inst, 1, 1)?.hyperlink).toBeDefined();
  });

  it('restores the link and its format after removeHyperlink', async () => {
    const inst = await mountSheet();
    const s = inst.workbook.sheet(0);
    s.setHyperlink({ range: 'B2', target: 'https://example.com/' });
    const xf = s.ooxml.getCellXfIndex('B2');
    expect(xf).not.toBeNull();
    inst.history.clear();
    s.removeHyperlink('B2');
    expect(s.listHyperlinks()).toHaveLength(0);
    expect(s.ooxml.getCellXfIndex('B2')).not.toBe(xf);
    inst.history.undo();
    expect(s.listHyperlinks()).toHaveLength(1);
    expect(s.ooxml.getCellXfIndex('B2')).toBe(xf);
    expect(formatAt(inst, 1, 1)?.hyperlink).toBeDefined();
    inst.history.redo();
    expect(s.listHyperlinks()).toHaveLength(0);
    expect(formatAt(inst, 1, 1)?.hyperlink).toBeUndefined();
  });

  it('round-trips clearHyperlinks', async () => {
    const inst = await mountSheet();
    const s = inst.workbook.sheet(0);
    s.setHyperlink({ range: 'B2', target: 'https://example.com/' });
    s.setHyperlink({ range: 'C3', target: 'https://example.org/' });
    inst.history.clear();
    s.clearHyperlinks();
    expect(s.listHyperlinks()).toHaveLength(0);
    inst.history.undo();
    expect(s.listHyperlinks()).toHaveLength(2);
    expect(formatAt(inst, 2, 2)?.hyperlink).toBeDefined();
    inst.history.redo();
    expect(s.listHyperlinks()).toHaveLength(0);
  });

  it('round-trips setCellPhonetic', async () => {
    const inst = await mountSheet();
    const s = inst.workbook.sheet(0);
    s.setText('A1', 'kanji');
    inst.history.clear();
    s.setCellPhonetic('A1', 'kana');
    expect(s.getCellPhonetic('A1')).toBe('kana');
    inst.history.undo();
    expect(s.getCellPhonetic('A1')).toBeNull();
    inst.history.redo();
    expect(s.getCellPhonetic('A1')).toBe('kana');
  });

  it('round-trips updateFormatDefaults', async () => {
    const inst = await mountSheet();
    const s = inst.workbook.sheet(0);
    const before = s.getFormatDefaults().defaultColumnWidthChars;
    s.updateFormatDefaults({ defaultColumnWidthChars: 17 });
    expect(s.getFormatDefaults().defaultColumnWidthChars).toBe(17);
    inst.history.undo();
    expect(s.getFormatDefaults().defaultColumnWidthChars).toBe(before);
    inst.history.redo();
    expect(s.getFormatDefaults().defaultColumnWidthChars).toBe(17);
  });

  it('round-trips setThemeColors', async () => {
    const inst = await mountSheet();
    const wb = inst.workbook;
    const before = wb.getThemeColors().colors;
    wb.setThemeColors({ colors: THEME });
    expect(wb.getThemeColors().colors[4]).toBe(0xffff0000);
    inst.history.undo();
    expect(wb.getThemeColors().colors).toEqual(before);
    inst.history.redo();
    expect(wb.getThemeColors().colors[4]).toBe(0xffff0000);
  });

  it('round-trips setCellStyle', async () => {
    const inst = await mountSheet();
    const wb = inst.workbook;
    const names = () => wb.listCellStyles().map((style) => style.name);
    expect(names()).not.toContain('Journal Style');
    wb.setCellStyle({ name: 'Journal Style' });
    expect(names()).toContain('Journal Style');
    inst.history.undo();
    expect(names()).not.toContain('Journal Style');
    inst.history.redo();
    expect(names()).toContain('Journal Style');
  });

  it('round-trips ooxml.setFont', async () => {
    const inst = await mountSheet();
    const ooxml = inst.workbook.ooxml;
    const before = ooxml.getFont(0);
    ooxml.setFont(0, { ...before, bold: !before.bold });
    expect(ooxml.getFont(0).bold).toBe(!before.bold);
    inst.history.undo();
    expect(ooxml.getFont(0).bold).toBe(before.bold);
    inst.history.redo();
    expect(ooxml.getFont(0).bold).toBe(!before.bold);
  });

  it('round-trips a sheet XF index write', async () => {
    const inst = await mountSheet();
    const s = inst.workbook.sheet(0);
    s.updateCellFormat('A1', { font: { bold: true } });
    const xf = s.ooxml.getCellXfIndex('A1') as number;
    inst.history.clear();
    s.ooxml.setCellXfIndex('C3', xf);
    expect(s.ooxml.getCellXfIndex('C3')).toBe(xf);
    inst.history.undo();
    expect(s.ooxml.getCellXfIndex('C3')).toBeNull();
    inst.history.redo();
    expect(s.ooxml.getCellXfIndex('C3')).toBe(xf);
  });

  it('pushes nothing for ooxml.addFont', async () => {
    const inst = await mountSheet();
    const ooxml = inst.workbook.ooxml;
    ooxml.addFont({ ...ooxml.getFont(0), bold: true });
    expect(inst.history.canUndo).toBe(false);
  });
});
