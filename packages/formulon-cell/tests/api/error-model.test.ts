import { FormulonError, StatusCode } from '@libraz/formulon';
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

describe('error model: partial application', () => {
  it('applyChanges reports refused cells instead of throwing or returning false', async () => {
    const instance = await mount(makeHost());
    disposers.push(instance);
    await instance.workbook.sheet(0).protect();
    const result = instance.applyChanges([{ at: 'A1', input: '1' }], { history: 'record' });
    expect(result.status).toBe('refused');
    expect(result.rejections.length).toBeGreaterThan(0);
  });
});
