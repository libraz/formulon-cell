import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { PivotAggregation, PivotAxis } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings.js';
import { attachWorkbookObjectsPanel } from '../../../src/interact/workbook-objects.js';

const fixture = resolve(
  dirname(fileURLToPath(import.meta.url)),
  '../../fixtures/charts/chart_pivot.xlsx',
);

const openEditor = async (
  fieldList = false,
): Promise<{ wb: WorkbookHandle; form: HTMLFormElement; detach: () => void }> => {
  const wb = await WorkbookHandle.loadBytes(new Uint8Array(readFileSync(fixture)));
  const host = document.createElement('div');
  document.body.appendChild(host);
  const handle = attachWorkbookObjectsPanel({ host, wb, strings: en });
  handle.open();
  const label = fieldList ? en.workbookObjects.pivotFieldList : en.workbookObjects.editPivotTable;
  const button = Array.from(host.querySelectorAll<HTMLButtonElement>('button')).find(
    (el) => el.textContent === label,
  );
  button?.click();
  const form = host.querySelector<HTMLFormElement>('.fc-objects__pivot-edit');
  if (!form) throw new Error('missing pivot edit form');
  return {
    wb,
    form,
    detach: () => {
      handle.detach();
      host.remove();
    },
  };
};

describe('PivotTable editor on a report loaded from a file', () => {
  it('shows the stored definition rather than a guess', async () => {
    const { wb, form, detach } = await openEditor();
    const inputs = Array.from(form.querySelectorAll<HTMLInputElement>('input[type="text"]'));
    expect(inputs[0]?.value).toBe('PT1');
    const axisOf = (index: number) =>
      form.querySelector<HTMLSelectElement>(`[data-pivot-field-index="${index}"]`)?.value;
    expect(axisOf(0)).toBe(String(PivotAxis.Row));
    expect(axisOf(1)).toBe('none');
    expect(axisOf(2)).toBe(String(PivotAxis.Value));
    const aggregation = form.querySelector<HTMLSelectElement>(
      '[data-pivot-aggregation-field-index="2"]',
    );
    expect(aggregation?.value).toBe(String(PivotAggregation.Sum));
    const items = Array.from(
      form.querySelectorAll<HTMLInputElement>(
        '[data-pivot-filter-checklist-field-index="0"] input',
      ),
    );
    expect(items.map((item) => [item.value, item.checked])).toEqual([
      ['East', true],
      ['North', true],
      ['West', true],
    ]);
    detach();
    wb.dispose();
  });

  it('keeps the stored data-field name when an apply changes nothing', async () => {
    const { wb, form, detach } = await openEditor(true);
    form.dispatchEvent(new SubmitEvent('submit', { bubbles: true, cancelable: true }));
    const definition = wb.getPivotDefinition(1, 0);
    expect(definition?.dataFields).toMatchObject([{ name: '合計 / Sales', fieldIndex: 2 }]);
    expect(definition?.rowFieldOrder).toEqual([0]);
    expect(definition?.fields[1]?.axis).toBeNull();
    detach();
    wb.dispose();
  });

  it('renames a data field from the profile when its aggregation changes', async () => {
    const { wb, form, detach } = await openEditor(true);
    const aggregation = form.querySelector<HTMLSelectElement>(
      '[data-pivot-aggregation-field-index="2"]',
    );
    if (!aggregation) throw new Error('missing aggregation control');
    aggregation.value = String(PivotAggregation.Max);
    form.dispatchEvent(new SubmitEvent('submit', { bubbles: true, cancelable: true }));
    expect(wb.getPivotDefinition(1, 0)?.dataFields).toMatchObject([
      { name: 'Max of Sales', fieldIndex: 2, aggregation: PivotAggregation.Max },
    ]);
    detach();
    wb.dispose();
  });
});
