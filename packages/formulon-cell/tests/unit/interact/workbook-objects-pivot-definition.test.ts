import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  PivotAggregation,
  PivotAxis,
  type SpreadsheetProfileId,
} from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings.js';
import { attachWorkbookObjectsPanel } from '../../../src/interact/workbook-objects.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const fixture = resolve(
  dirname(fileURLToPath(import.meta.url)),
  '../../fixtures/charts/chart_pivot.xlsx',
);

const openEditor = async (
  fieldList = false,
  profile?: SpreadsheetProfileId,
): Promise<{ wb: WorkbookHandle; form: HTMLElement; detach: () => void }> => {
  const wb = await WorkbookHandle.loadBytes(new Uint8Array(readFileSync(fixture)));
  if (profile) wb.setSpreadsheetProfileId(profile);
  const host = document.createElement('div');
  document.body.appendChild(host);
  const handle = attachWorkbookObjectsPanel({
    store: createSpreadsheetStore(),
    host,
    wb,
    strings: en,
  });
  handle.open();
  const label = fieldList ? en.workbookObjects.pivotFieldList : en.workbookObjects.editPivotTable;
  const button = Array.from(host.querySelectorAll<HTMLButtonElement>('button')).find(
    (el) => el.textContent === label,
  );
  button?.click();
  const form = host.querySelector<HTMLElement>(
    fieldList ? '.fc-objects__pivot-pane' : '.fc-objects__pivot-edit',
  );
  if (!form) throw new Error('missing pivot editor');
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

  it('keeps the stored data-field name when a field list change leaves Values alone', async () => {
    const { wb, form, detach } = await openEditor(true);
    const unused = form.querySelectorAll<HTMLInputElement>('[data-pivot-field-list-field]')[1];
    if (!unused) throw new Error('missing field checkbox');
    expect(unused.checked).toBe(false);
    unused.checked = true;
    unused.dispatchEvent(new Event('change', { bubbles: true }));
    const definition = wb.getPivotDefinition(1, 0);
    expect(definition?.dataFields).toMatchObject([{ name: '合計 / Sales', fieldIndex: 2 }]);
    expect(definition?.rowFieldOrder).toEqual([1, 0]);
    detach();
    wb.dispose();
  });

  it('renames a default data-field caption after the new function', async () => {
    // The file was saved by the Japanese app, so its captions follow that profile.
    const { wb, form, detach } = await openEditor(true, 'mac-ja_JP');
    form.querySelector<HTMLElement>('[data-pivot-area="values"] [data-pivot-area-field]')?.click();
    const max = document.querySelector<HTMLElement>(
      `.fc-pvfsdlg [data-pvfs-aggregation] [data-pvfs-value="${PivotAggregation.Max}"]`,
    );
    if (!max) throw new Error('missing Max in the field dialog');
    max.click();
    document.querySelector<HTMLButtonElement>('.fc-pvfsdlg [data-pvfs-ok]')?.click();
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(wb.getPivotDefinition(1, 0)?.dataFields).toMatchObject([
      { name: '最大 / Sales', fieldIndex: 2, aggregation: PivotAggregation.Max },
    ]);
    detach();
    wb.dispose();
  });
});
