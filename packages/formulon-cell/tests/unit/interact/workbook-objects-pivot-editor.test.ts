import { beforeEach, describe, expect, it, vi } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import { PivotAggregation, PivotAxis } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings.js';
import { attachWorkbookObjectsPanel } from '../../../src/interact/workbook-objects.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const setup = async (valueFields = ['Sales', 'Qty']) => {
  const wb = await WorkbookHandle.createDefault();
  const rows: (string | number)[][] = [
    ['Region', 'Product', 'Sales', 'Qty'],
    ['East', 'Desk', 12, 1],
    ['West', 'Chair', 5, 3],
  ];
  for (const [r, row] of rows.entries()) {
    for (const [c, value] of row.entries()) {
      const addr = { sheet: 0, row: r, col: c };
      if (typeof value === 'number') wb.setNumber(addr, value);
      else wb.setText(addr, value);
    }
  }
  const created = createPivotTableFromRange(wb, {
    source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 3 },
    destination: { sheet: 0, row: 6, col: 0 },
    rowField: 'Region',
    valueFields,
    valueField: 'Sales',
  });
  expect(created.ok).toBe(true);
  return wb;
};

const submit = (form: HTMLFormElement) =>
  form.dispatchEvent(new SubmitEvent('submit', { bubbles: true, cancelable: true }));

describe('PivotTable editors against the engine', () => {
  let host: HTMLElement;
  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  const open = (wb: WorkbookHandle) => {
    const handle = attachWorkbookObjectsPanel({
      store: createSpreadsheetStore(),
      host,
      wb,
      strings: en,
    });
    handle.open();
    Array.from(host.querySelectorAll<HTMLButtonElement>('button'))
      .find((el) => el.textContent === en.workbookObjects.editPivotTable)
      ?.click();
    const form = host.querySelector<HTMLFormElement>('.fc-objects__pivot-edit');
    if (!form) throw new Error('missing pivot edit form');
    return { handle, form };
  };

  const openFieldList = (wb: WorkbookHandle) => {
    const handle = attachWorkbookObjectsPanel({
      store: createSpreadsheetStore(),
      host,
      wb,
      strings: en,
    });
    expect(handle.openPivotFieldList(0, 0)).toBe(true);
    const tick = (field: string, checked: boolean): void => {
      const check = host.querySelector<HTMLInputElement>(
        `.fc-objects__pivot-pane [data-pivot-field-list-field="${field}"]`,
      );
      if (!check) throw new Error(`missing ${field} checkbox`);
      expect(check.checked).toBe(!checked);
      check.checked = checked;
      check.dispatchEvent(new Event('change', { bubbles: true }));
    };
    return { handle, tick };
  };

  it('unticks a value field and keeps the other data field', async () => {
    const wb = await setup();
    const { handle, tick } = openFieldList(wb);
    tick('Sales', false);

    const definition = wb.getPivotDefinition(0, 0);
    expect(definition?.fields[2]?.axis).toBeNull();
    expect(definition?.dataFields.map((d) => d.fieldIndex)).toEqual([3]);
    handle.detach();
    wb.dispose();
  });

  it('ticks an unused text field in atop Rows and an unused number field atop Values', async () => {
    const wb = await setup(['Sales']);
    const { handle, tick } = openFieldList(wb);
    tick('Product', true);
    tick('Qty', true);

    const definition = wb.getPivotDefinition(0, 0);
    expect(definition?.fields[1]?.axis).toBe(PivotAxis.Row);
    expect(definition?.fields[3]?.axis).toBe(PivotAxis.Value);
    expect(definition?.rowFieldOrder).toEqual([1, 0]);
    expect(definition?.dataFields.map((d) => d.fieldIndex)).toEqual([3, 2]);
    handle.detach();
    wb.dispose();
  });

  it('writes custom subtotal functions chosen in the field settings', async () => {
    const wb = await setup();
    const { handle, form } = open(wb);
    const mode = form.querySelector<HTMLSelectElement>(
      '[data-pivot-subtotal-mode-field-index="0"]',
    );
    if (!mode) throw new Error('missing subtotal mode');
    mode.value = 'custom';
    mode.dispatchEvent(new Event('change', { bubbles: true }));
    const sum = form.querySelector<HTMLInputElement>(
      `[data-pivot-subtotal-fns-field-index="0"] input[value="${PivotAggregation.Sum}"]`,
    );
    if (!sum) throw new Error('missing subtotal function');
    sum.checked = true;
    submit(form);

    expect(wb.getPivotField(0, 0, 0)).toMatchObject({
      defaultSubtotal: false,
      subtotalFns: [PivotAggregation.Sum],
    });
    handle.detach();
    wb.dispose();
  });

  it('keeps the dialog open when Custom is chosen with no function ticked', async () => {
    const wb = await setup();
    const clear = vi.spyOn(wb, 'clearPivotFieldSubtotalFns');
    const { handle, form } = open(wb);
    const mode = form.querySelector<HTMLSelectElement>(
      '[data-pivot-subtotal-mode-field-index="0"]',
    );
    if (!mode) throw new Error('missing subtotal mode');
    mode.value = 'custom';
    mode.dispatchEvent(new Event('change', { bubbles: true }));
    submit(form);

    expect(host.textContent).toContain('Select at least one function');
    expect(host.querySelector('.fc-objects__pivot-edit')).toBeTruthy();
    expect(clear).not.toHaveBeenCalled();
    expect(wb.getPivotField(0, 0, 0)).toMatchObject({ defaultSubtotal: true, subtotalFns: [] });
    handle.detach();
    wb.dispose();
  });

  it('writes Automatic for a field that had no subtotal', async () => {
    const wb = await setup();
    const real = wb.getPivotField.bind(wb);
    vi.spyOn(wb, 'getPivotField').mockImplementation((sheet, pivot, field) => {
      const definition = real(sheet, pivot, field);
      return definition && field === 0 ? { ...definition, defaultSubtotal: false } : definition;
    });
    const clear = vi.spyOn(wb, 'clearPivotFieldSubtotalFns');
    const { handle, form } = open(wb);
    const mode = form.querySelector<HTMLSelectElement>(
      '[data-pivot-subtotal-mode-field-index="0"]',
    );
    if (!mode) throw new Error('missing subtotal mode');
    expect(mode.value).toBe('none');
    mode.value = 'auto';
    mode.dispatchEvent(new Event('change', { bubbles: true }));
    submit(form);

    expect(clear).toHaveBeenCalledWith(0, 0, 0);
    handle.detach();
    wb.dispose();
  });
  it('applies and deletes from the editor as one undo step each', async () => {
    const wb = await setup(['Sales']);
    const history = new History();
    const handle = attachWorkbookObjectsPanel({
      store: createSpreadsheetStore(),
      host,
      wb,
      strings: en,
      history,
    });
    const editForm = (): HTMLFormElement => {
      handle.open();
      Array.from(host.querySelectorAll<HTMLButtonElement>('button'))
        .find((el) => el.textContent === en.workbookObjects.editPivotTable)
        ?.click();
      const form = host.querySelector<HTMLFormElement>('.fc-objects__pivot-edit');
      if (!form) throw new Error('missing pivot edit form');
      return form;
    };
    const rowTotals = () => wb.getPivotDefinition(0, 0)?.grandTotals.rows;

    const form = editForm();
    const check = Array.from(form.querySelectorAll('label')).find(
      (label) => label.textContent === en.workbookObjects.rowGrandTotals,
    )?.firstElementChild as HTMLInputElement;
    check.checked = false;
    submit(form);
    expect(rowTotals()).toBe(false);
    expect(history.undo()).toBe(true);
    expect(rowTotals()).toBe(true);
    expect(history.canUndo()).toBe(false);
    expect(history.redo()).toBe(true);
    expect(rowTotals()).toBe(false);

    Array.from(editForm().querySelectorAll<HTMLButtonElement>('button'))
      .find((el) => el.textContent === en.workbookObjects.deletePivotTable)
      ?.click();
    expect(wb.getPivotTables()).toHaveLength(0);
    expect(history.undo()).toBe(true);
    expect(wb.getPivotTables()).toHaveLength(1);
    expect(rowTotals()).toBe(false);
    expect(history.undo()).toBe(true);
    expect(rowTotals()).toBe(true);
    handle.detach();
    wb.dispose();
  });
});
