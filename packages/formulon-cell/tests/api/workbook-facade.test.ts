import { ChartPlotType, FormulonError, StatusCode, ValueKind } from '@libraz/formulon';
import { SAVE_FORMATS, Workbook } from '@libraz/formulon-cell';
import { afterEach, describe, expect, it } from 'vitest';
import { expectFailure, expectFormulonError } from './helpers.js';

const open: Workbook[] = [];

const createWorkbook = async (): Promise<Workbook> => {
  const wb = await Workbook.createDefault({ locale: 'en' });
  open.push(wb);
  return wb;
};

afterEach(() => {
  for (const wb of open.splice(0)) wb.dispose();
});

describe('workbook root', () => {
  it('creates a workbook on the wasm engine with one sheet', async () => {
    const wb = await createWorkbook();
    expect(wb.engine).toBe('wasm');
    expect(wb.listSheets()).toHaveLength(1);
    expect(wb.disposed).toBe(false);
  });

  it('returns one facade per sheet, whichever key reaches it', async () => {
    const wb = await createWorkbook();
    const first = wb.sheet(0);
    expect(wb.sheet(first.name)).toBe(first);
    expect(wb.sheet(first)).toBe(first);
    expect(wb.listSheets()[0]).toBe(first);
  });

  it('adds, renames and removes sheets; a removed sheet facade is stale', async () => {
    const wb = await createWorkbook();
    const data = wb.addSheet('Data');
    data.rename('Figures');
    expect(wb.sheet('Figures')).toBe(data);
    wb.removeSheet(data);
    expectFormulonError(() => wb.sheet('Figures'), StatusCode.NotFound);
    expectFormulonError(() => data.getTypedValue('A1'), StatusCode.NotFound);
  });

  it('round-trips through save and load, defaulting to the loaded format', async () => {
    const wb = await createWorkbook();
    wb.sheet(0).setNumber('A1', 42);
    const saved = wb.save();
    expect(saved.format).toBe('xlsx');
    expect(saved.bytes).toBeInstanceOf(Uint8Array);
    expect(saved.losses).toEqual([]);

    const xlsb = wb.save({ format: 'xlsb' });
    const reloaded = await Workbook.load(xlsb.bytes);
    open.push(reloaded);
    expect(reloaded.save().format).toBe('xlsb');
    expect(reloaded.loadLosses).toEqual([]);
    expect(reloaded.sheet(0).getTypedValue('A1')).toEqual({ kind: ValueKind.Number, number: 42 });
  });

  it('describes every save format by extension and MIME type', () => {
    expect(Object.keys(SAVE_FORMATS).sort()).toEqual(['xlsb', 'xlsx']);
    expect(SAVE_FORMATS.xlsx.extension).toBe('.xlsx');
    expect(typeof SAVE_FORMATS.xlsb.mimeType).toBe('string');
  });

  it('formats a value with the engine number formatter', async () => {
    const wb = await createWorkbook();
    expect(wb.formatValue(0.25, '0.0%').text).toBe('25.0%');
  });

  it('batches writes into one change event and keeps cell writes atomic', async () => {
    const wb = await createWorkbook();
    const s = wb.sheet(0);
    const events: unknown[] = [];
    wb.on('change', (e) => events.push(e));

    wb.batch(() => {
      for (let r = 0; r < 10; r++) s.setNumber(r, 0, r);
      wb.batch(() => s.setFormula('B1', '=SUM(A1:A10)'));
    });
    expect(events).toHaveLength(1);
    expect(s.getTypedValue('B1')).toEqual({ kind: ValueKind.Number, number: 45 });

    s.setNumber('C1', 1);
    const boom = new Error('boom');
    expect(() =>
      wb.batch(() => {
        s.setNumber('C1', 2);
        s.addMerge('E1:F1');
        throw boom;
      }),
    ).toThrow(boom);
    expect(s.getTypedValue('C1')).toEqual({ kind: ValueKind.Number, number: 1 });
    // Structural writes are not rolled back.
    expect(s.listMerges()).toHaveLength(1);
  });
});

describe('stub engine', () => {
  it('reports the stub engine and throws NotImplemented for members it lacks', async () => {
    const wb = await Workbook.createDefault({ engine: 'stub' });
    open.push(wb);
    expect(wb.engine).toBe('stub');
    expectFormulonError(() => wb.sheet(0).addMerge('A1:B2'), StatusCode.NotImplemented);
  });
});

describe('sheet cells', () => {
  it('writes and reads values and formulas in A1 and (row, col) form', async () => {
    const wb = await createWorkbook();
    const s = wb.sheet(0);
    s.setNumber('A1', 10);
    s.setFormula(0, 1, '=A1*2');
    expect(s.getTypedValue(0, 1)).toEqual({ kind: ValueKind.Number, number: 20 });
    expect(s.getFormula('B1')).toBe('=A1*2');
  });

  it('writes a block with setValues that a failed batch restores as a whole', async () => {
    const wb = await createWorkbook();
    const s = wb.sheet(0);
    s.setValues('A1', [
      [1, 2],
      [3, 4],
    ]);
    expect(s.getTypedValue('B2')).toEqual({ kind: ValueKind.Number, number: 4 });
    const boom = new Error('abort');
    expect(() =>
      wb.batch(() => {
        s.setValues(0, 0, [
          [9, 9],
          [9, 9],
        ]);
        throw boom;
      }),
    ).toThrow(boom);
    expect(s.getTypedValue('A1')).toEqual({ kind: ValueKind.Number, number: 1 });
    expect(s.getTypedValue('B2')).toEqual({ kind: ValueKind.Number, number: 4 });
  });
});

describe('evaluate', () => {
  it('evaluates a formula at a cell without writing it', async () => {
    const wb = await createWorkbook();
    const s = wb.sheet(0);
    s.setNumber('A1', 3);
    s.setNumber('B1', 4);
    expect(s.evaluateTypedFormula('C1', '=SUM(A1:B1)')).toEqual({
      kind: ValueKind.Number,
      number: 7,
    });
    expect(s.evaluateTypedFormulaArray('C1', '=A1:B1*2')).toEqual([
      [
        { kind: ValueKind.Number, number: 6 },
        { kind: ValueKind.Number, number: 8 },
      ],
    ]);
    expect(s.getTypedValue('C1').kind).toBe(ValueKind.Blank);
  });
});

describe('csv', () => {
  it('imports at an anchor and exports a range', async () => {
    const wb = await createWorkbook();
    const s = wb.sheet(0);
    const result = s.importCsv('a,1\nb,2', { at: 'B2' });
    expect(result.status).toBe('applied');
    expect(result.rejections).toEqual([]);
    expect(s.getTypedValue('C3')).toEqual({ kind: ValueKind.Number, number: 2 });
    expect(s.exportCsv({ range: 'B2:C3' })).toBe('a,1\r\nb,2');
  });

  it('exports an empty sheet as empty text', async () => {
    const wb = await createWorkbook();
    expect(wb.sheet(0).exportCsv()).toBe('');
  });

  it('refuses more cells than the limit without writing any', async () => {
    const wb = await createWorkbook();
    const s = wb.sheet(0);
    const wide = Array.from({ length: 100_001 }, () => '1').join(',');
    expectFailure(() => s.importCsv(wide, { at: 'A1' }), 'limit-exceeded');
    expect(s.getTypedValue('A1').kind).toBe(ValueKind.Blank);
    expectFailure(() => s.exportCsv({ range: 'A1:XFD100' }), 'limit-exceeded');
  });
});

describe('sheet structure', () => {
  it('inserts rows and merges cells through engine verbs', async () => {
    const wb = await createWorkbook();
    const s = wb.sheet(0);
    s.setNumber('A1', 1);
    s.insertRows(0, 0);
    expect(s.getTypedValue('A2')).toEqual({ kind: ValueKind.Number, number: 1 });
    s.addMerge('C1:D2');
    expect(s.listMerges()).toHaveLength(1);
  });
});

describe('sheet format', () => {
  it('reads back a patched cell format', async () => {
    const wb = await createWorkbook();
    const s = wb.sheet(0);
    s.updateCellFormat('A1:A2', { numberFormat: '0.00%' });
    expect(s.getCellFormat('A2').numberFormat).toBe('0.00%');
  });
});

describe('protection', () => {
  it('protects with a password and unprotects only with it', async () => {
    const wb = await createWorkbook();
    const s = wb.sheet(0);
    await s.protect({ password: 'secret' });
    expect(s.getSheetProtection()).not.toBeNull();
    await expect(s.unprotect('wrong')).rejects.toBeInstanceOf(FormulonError);
    expect(s.getSheetProtection()).not.toBeNull();
    await s.unprotect('secret');
    expect(s.getSheetProtection()).toBeNull();
  });
});

describe('print', () => {
  it('sets the print area through the engine record', async () => {
    const wb = await createWorkbook();
    const s = wb.sheet(0);
    s.setPrintArea('A1:C10');
    expect(s.getPrintArea()).toEqual([{ firstRow: 0, firstCol: 0, lastRow: 9, lastCol: 2 }]);
  });
});

describe('ooxml', () => {
  it('reaches the workbook and sheet XML-level surfaces', async () => {
    const wb = await createWorkbook();
    expect(Array.isArray(wb.ooxml.listNumberFormats())).toBe(true);
    expect(wb.sheet(0).ooxml.getAutoFilterXml()).toBeNull();
  });
});

describe('table', () => {
  it('adds, finds, updates and removes a table; the removed facade is stale', async () => {
    const wb = await createWorkbook();
    const s = wb.sheet(0);
    const t = s.addTable({ range: 'A1:B5', name: 'Sales' });
    expect(wb.table('Sales')).toBe(t);
    expect(s.table('Sales')).toBe(t);
    t.update({ totalsRow: true });
    expect(t.getProperties().totalsRow).toBe(true);
    s.removeTable(t);
    expectFormulonError(() => wb.table('Sales'), StatusCode.NotFound);
    expectFormulonError(() => t.getProperties(), StatusCode.NotFound);
  });
});

describe('drawing', () => {
  it('lists every drawing object and removes one by handle', async () => {
    const wb = await createWorkbook();
    const s = wb.sheet(0);
    const shape = s.addShape({ presetGeometry: 'rect', text: 'Note' });
    expect(s.listDrawings()).toEqual([shape]);
    expect(s.drawing(shape.objectId)).toBe(shape);
    s.removeDrawing(shape);
    expectFormulonError(() => shape.getProperties(), StatusCode.NotFound);
  });
});

describe('shape', () => {
  it('updates a shape through its handle', async () => {
    const wb = await createWorkbook();
    const shape = wb.sheet(0).addShape({ presetGeometry: 'rect' });
    shape.update({ text: 'Hello' });
    expect(shape.getProperties().text).toBe('Hello');
  });
});

describe('chart', () => {
  it('adds a chart over a range and reads its data', async () => {
    const wb = await createWorkbook();
    const s = wb.sheet(0);
    s.setNumber('A1', 1);
    s.setNumber('A2', 2);
    const chart = s.addChart({ type: ChartPlotType.Bar, range: 'A1:A2' });
    expect(s.chart(chart.objectId)).toBe(chart);
    expect(chart.getData().series.length).toBeGreaterThan(0);
  });
});

describe('picture', () => {
  it('stale picture handle lookups throw NotFound', async () => {
    const wb = await createWorkbook();
    expectFormulonError(() => wb.sheet(0).picture(999), StatusCode.NotFound);
  });
});

describe('slicer', () => {
  it('stale slicer handle lookups throw NotFound', async () => {
    const wb = await createWorkbook();
    expectFormulonError(() => wb.sheet(0).slicer(999), StatusCode.NotFound);
  });
});

describe('pivot cache', () => {
  it('adds a cache with a worksheet source and removes it', async () => {
    const wb = await createWorkbook();
    const s = wb.sheet(0);
    s.setText('A1', 'Region');
    s.setText('A2', 'East');
    const cache = wb.addPivotCache();
    cache.setWorksheetSource({ range: 'A1:A2', sheet: s });
    expect(wb.pivotCache(cache.cacheId)).toBe(cache);
    expect(cache.getWorksheetSource()).not.toBeNull();
    wb.removePivotCache(cache);
    expectFormulonError(() => cache.getWorksheetSource(), StatusCode.NotFound);
  });
});

describe('pivot', () => {
  it('pivot lookups by an unknown name throw NotFound', async () => {
    const wb = await createWorkbook();
    expectFormulonError(() => wb.sheet(0).pivot('Missing'), StatusCode.NotFound);
    expect(wb.sheet(0).listPivots()).toEqual([]);
  });
});
