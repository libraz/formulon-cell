import { beforeEach, describe, expect, it, vi } from 'vitest';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import { PivotAggregation, PivotAxis } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings.js';
import { attachWorkbookObjectsPanel } from '../../../src/interact/workbook-objects.js';

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

describe('PivotTable Field List against the engine', () => {
  let host: HTMLElement;
  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  const open = (wb: WorkbookHandle) => {
    const handle = attachWorkbookObjectsPanel({ host, wb, strings: en });
    expect(handle.openPivotFieldList(0, 0)).toBe(true);
    const form = host.querySelector<HTMLFormElement>('.fc-objects__pivot-edit');
    if (!form) throw new Error('missing field list form');
    return { handle, form };
  };

  it('unticks a value field and keeps the other data field', async () => {
    const wb = await setup();
    const { handle, form } = open(wb);
    const check = form.querySelector<HTMLInputElement>('[data-pivot-field-list-field-index="2"]');
    if (!check) throw new Error('missing checkbox');
    expect(check.checked).toBe(true);
    check.checked = false;
    check.dispatchEvent(new Event('change', { bubbles: true }));
    submit(form);

    const definition = wb.getPivotDefinition(0, 0);
    expect(definition?.fields[2]?.axis).toBeNull();
    expect(definition?.dataFields.map((d) => d.fieldIndex)).toEqual([3]);
    handle.detach();
    wb.dispose();
  });

  it('ticks an unused text field in on Rows and an unused number field on Values', async () => {
    const wb = await setup(['Sales']);
    const { handle, form } = open(wb);
    for (const index of [1, 3]) {
      const check = form.querySelector<HTMLInputElement>(
        `[data-pivot-field-list-field-index="${index}"]`,
      );
      if (!check) throw new Error('missing checkbox');
      expect(check.checked).toBe(false);
      check.checked = true;
      check.dispatchEvent(new Event('change', { bubbles: true }));
    }
    submit(form);

    const definition = wb.getPivotDefinition(0, 0);
    expect(definition?.fields[1]?.axis).toBe(PivotAxis.Row);
    expect(definition?.fields[3]?.axis).toBe(PivotAxis.Value);
    expect(definition?.rowFieldOrder).toEqual([0, 1]);
    expect(definition?.dataFields.map((d) => d.fieldIndex)).toEqual([2, 3]);
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
});
