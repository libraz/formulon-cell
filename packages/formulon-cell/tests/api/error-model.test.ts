import { FormulonError, StatusCode, ValueKind } from '@libraz/formulon';
import { mount, spreadsheetFailureOf, viewerPolicy, Workbook } from '@libraz/formulon-cell';
import { afterEach, describe, expect, it } from 'vitest';
import { expectFailure, expectFormulonError, makeHost, thrownBy } from './helpers.js';

const disposers: Array<{ dispose(): void }> = [];

const createWorkbook = async (): Promise<Workbook> => {
  const wb = await Workbook.createDefault({ locale: 'en' });
  disposers.push(wb);
  return wb;
};

afterEach(() => {
  for (const d of disposers.splice(0).reverse()) d.dispose();
});

describe('error model: engine failures propagate untouched', () => {
  it('Workbook: an unknown sheet name', async () => {
    const wb = await createWorkbook();
    expectFormulonError(() => wb.sheet('Nope'), StatusCode.NotFound);
  });

  it('Sheet: an invalid sheet name', async () => {
    const wb = await createWorkbook();
    expectFormulonError(() => wb.sheet(0).rename('a/b'), StatusCode.InvalidSheetName);
  });

  it('Table: a duplicate name', async () => {
    const wb = await createWorkbook();
    const s = wb.sheet(0);
    s.addTable({ range: 'A1:B3', name: 'One' });
    const two = s.addTable({ range: 'D1:E3', name: 'Two' });
    expectFormulonError(() => two.rename('One'), StatusCode.AlreadyExists);
  });

  it('PivotCache: an unknown cache id', async () => {
    const wb = await createWorkbook();
    expectFormulonError(() => wb.pivotCache(999), StatusCode.NotFound);
  });

  it('Pivot: an unknown pivot', async () => {
    const wb = await createWorkbook();
    expectFormulonError(() => wb.sheet(0).pivot(0), StatusCode.IndexOutOfRange);
  });

  it('Drawing: an unknown object id', async () => {
    const wb = await createWorkbook();
    expectFormulonError(() => wb.sheet(0).drawing(999), StatusCode.NotFound);
  });

  it('Chart: an unknown object id', async () => {
    const wb = await createWorkbook();
    expectFormulonError(() => wb.sheet(0).chart(999), StatusCode.NotFound);
  });

  it('Shape: an unknown patch key', async () => {
    const wb = await createWorkbook();
    const shape = wb
      .sheet(0)
      .addShape({ presetGeometry: 'rect', anchor: { widthEmu: 952500, heightEmu: 476250 } });
    expect(thrownBy(() => shape.update({ bogus: 1 } as never))).toBeInstanceOf(TypeError);
  });

  it('a formula that does not parse carries its parse position', async () => {
    const wb = await createWorkbook();
    const error = thrownBy(() => wb.sheet(0).setFormula('A1', '=SUM(')) as FormulonError;
    expect(error).toBeInstanceOf(FormulonError);
    expect(error.parseError).toBeDefined();
  });
});

describe('error model: argument shape', () => {
  it('throws RangeError for a non-integer, non-finite or negative coordinate', async () => {
    const wb = await createWorkbook();
    const s = wb.sheet(0);
    expect(thrownBy(() => s.setNumber(0.5, 0, 1))).toBeInstanceOf(RangeError);
    expect(thrownBy(() => s.setNumber(Number.NaN, 0, 1))).toBeInstanceOf(RangeError);
    expect(thrownBy(() => s.setNumber(2 ** 32, 0, 1))).toBeInstanceOf(RangeError);
  });

  it('throws TypeError for a boolean where a number is required', async () => {
    const wb = await createWorkbook();
    expect(thrownBy(() => wb.sheet(0).setNumber(true as never, 0, 1))).toBeInstanceOf(TypeError);
  });
});

describe('error model: formulon-cell preconditions', () => {
  it('a disposed workbook throws BindingInvalidHandle with the disposed failure', async () => {
    const wb = await Workbook.createDefault();
    wb.dispose();
    const error = expectFailure(() => wb.listSheets(), 'disposed');
    expect(error.code).toBe(StatusCode.BindingInvalidHandle);
  });

  it('opening a disabled dialog throws feature-disabled', async () => {
    const instance = await mount(makeHost(), { ui: { features: { pivotTable: false } } });
    disposers.push(instance);
    const error = expectFailure(() => instance.dialogs.open('pivotTable'), 'feature-disabled');
    expect(error.code).toBe(StatusCode.PreconditionFailed);
    expect(instance.dialogs.canOpen('pivotTable')).toBe(false);
  });

  it('a command the policy denies throws policy-denied', async () => {
    const instance = await mount(makeHost(), { policy: viewerPolicy() });
    disposers.push(instance);
    expectFailure(() => instance.commands.execute('insertRows'), 'policy-denied');
    expect(instance.commands.can('insertRows')).toBe(false);
  });

  it('a command on a protected sheet throws protected', async () => {
    const instance = await mount(makeHost());
    disposers.push(instance);
    await instance.workbook.sheet(0).protect();
    expectFailure(() => instance.commands.execute('insertRows'), 'protected');
  });

  it('spreadsheetFailureOf returns null for anything that is not a formulon-cell failure', () => {
    expect(spreadsheetFailureOf(new Error('x'))).toBeNull();
    expect(spreadsheetFailureOf(new FormulonError(StatusCode.NotFound, 'x'))).toBeNull();
    expect(spreadsheetFailureOf('protected')).toBeNull();
    expect(
      spreadsheetFailureOf(
        new FormulonError(StatusCode.PreconditionFailed, 'x', 'failure=protected sheet=0'),
      ),
    ).toBe('protected');
  });
});

describe('commands.can: intents', () => {
  it('decides an edit as applyChanges decides the same cells on a protected sheet', async () => {
    const instance = await mount(makeHost());
    disposers.push(instance);
    const sheet = instance.workbook.sheet(0);
    sheet.updateCellFormat('B1', { protection: { locked: false } });
    await sheet.protect();
    for (const at of ['A1', 'B1'] as const) {
      const allowed = instance.commands.can({ kind: 'edit', operation: 'valueEdit', at });
      const result = instance.applyChanges([{ at, input: '1' }], { history: 'record' });
      expect(result.status === 'applied', at).toBe(allowed);
    }
    expect(instance.commands.can({ kind: 'edit', operation: 'valueEdit', at: 'B1' })).toBe(true);
    expect(
      instance.commands.can({
        kind: 'edit',
        operation: 'valueEdit',
        sheet: 0,
        at: { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 1 },
      }),
    ).toBe(false);
    expect(
      instance.commands.can({ kind: 'edit', operation: 'valueEdit', at: { row: 0, col: 1 } }),
    ).toBe(true);
  });

  it('opens the hyperlink of the active cell through its command id', async () => {
    const instance = await mount(makeHost());
    disposers.push(instance);
    const sheet = instance.workbook.sheet(0);
    sheet.setHyperlink('A1', { target: null, location: `'${sheet.name}'!C5` });
    instance.selection.set('A1');
    expect(instance.commands.can('openHyperlink')).toBe(true);
    instance.commands.execute('openHyperlink');
    expect(instance.selection.get().active).toEqual({ row: 4, col: 2 });
  });

  it('answers select and copy from the policy', async () => {
    const instance = await mount(makeHost(), { policy: { ...viewerPolicy(), copy: false } });
    disposers.push(instance);
    expect(instance.commands.can({ kind: 'select' })).toBe(true);
    expect(instance.commands.can({ kind: 'copy' })).toBe(false);
    expect(instance.commands.can({ kind: 'edit', operation: 'valueEdit', at: 'A1' })).toBe(false);
  });
});

describe('instance.importCsv: the gated import', () => {
  it('imports into unlocked cells as one undo step and refuses locked ones whole', async () => {
    const instance = await mount(makeHost());
    disposers.push(instance);
    const sheet = instance.workbook.sheet(0);
    sheet.updateCellFormat('B1:B2', { protection: { locked: false } });
    await sheet.protect();

    const applied = instance.importCsv('1\n2', { at: 'B1' });
    expect(applied.status).toBe('applied');
    expect(sheet.getValue('B2')).toBe(2);
    instance.history.undo();
    expect(sheet.getTypedValue('B2').kind).toBe(ValueKind.Blank);

    const refused = instance.importCsv('1,2', { at: 'A1' });
    expect(refused.status).toBe('refused');
    expect(refused.rejections[0]?.code).toBe('protected');
    expect(sheet.getTypedValue('B1').kind).toBe(ValueKind.Blank);
  });

  it('reports more fields than the limit as a refusal', async () => {
    const instance = await mount(makeHost());
    disposers.push(instance);
    const line = Array.from({ length: 400 }, () => '1').join(',');
    const text = Array.from({ length: 300 }, () => line).join('\n');
    const result = instance.importCsv(text, { at: 'A1' });
    expect(result.status).toBe('refused');
    expect(result.rejections[0]?.code).toBe('effectTooLarge');
  });
});

describe('error model: all-or-nothing batches', () => {
  it('applyChanges writes nothing when one change is refused and reports that one refusal', async () => {
    const instance = await mount(makeHost());
    disposers.push(instance);
    const sheet = instance.workbook.sheet(0);
    sheet.updateCellFormat('B1', { protection: { locked: false } });
    await sheet.protect();
    const result = instance.applyChanges(
      [
        { at: 'B1', input: '2' },
        { at: 'A1', input: '1' },
      ],
      { history: 'record' },
    );
    expect(result.status).toBe('refused');
    expect(result.applied).toEqual([]);
    expect(result.rejections).toHaveLength(1);
    expect(sheet.getTypedValue('B1').kind).toBe(ValueKind.Blank);
  });

  it("gates applyChanges whatever its history mode: 'reset' on a locked cell writes nothing", async () => {
    const instance = await mount(makeHost());
    disposers.push(instance);
    const sheet = instance.workbook.sheet(0);
    await sheet.protect();
    const result = instance.applyChanges([{ at: 'A1', input: '1' }], { history: 'reset' });
    expect(result.status).toBe('refused');
    expect(result.rejections).toHaveLength(1);
    expect(sheet.getTypedValue('A1').kind).toBe(ValueKind.Blank);
  });
});
