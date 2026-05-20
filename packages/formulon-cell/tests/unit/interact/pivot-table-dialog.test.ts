import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import type { CellValue } from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings.js';
import { attachPivotTableDialog } from '../../../src/interact/pivot-table-dialog.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const text = (value: string): CellValue => ({ kind: 'text', value });
const num = (value: number): CellValue => ({ kind: 'number', value });

const dispatchDragEvent = (
  target: Element,
  type: string,
  dataTransfer: {
    effectAllowed?: string;
    setData(type: string, value: string): void;
    getData(type: string): string;
  },
): void => {
  const event = new Event(type, { bubbles: true, cancelable: true }) as DragEvent;
  Object.defineProperty(event, 'dataTransfer', { value: dataTransfer });
  target.dispatchEvent(event);
};

const makeDataTransfer = () => {
  const data = new Map<string, string>();
  return {
    effectAllowed: 'move',
    setData: (type: string, value: string) => data.set(type, value),
    getData: (type: string) => data.get(type) ?? '',
  };
};

const makeWb = () => {
  const calls: string[] = [];
  const wb = {
    capabilities: { pivotTableMutate: true },
    sheetCount: 1,
    sheetName: () => 'Sheet1',
    getPivotTables: () => [],
    getValue: ({ row, col }: { row: number; col: number }) => {
      if (row === 0 && col === 0) return text('Region');
      if (row === 0 && col === 1) return text('Sales');
      if (row === 0 && col === 2) return text('Qty');
      if (row === 0 && col === 3) return text('Channel');
      if (row === 0 && col === 4) return text('Segment');
      if (row === 1 && col === 0) return text('East');
      if (row === 1 && col === 1) return num(10);
      if (row === 1 && col === 2) return num(2);
      if (row === 1 && col === 3) return text('Online');
      if (row === 1 && col === 4) return text('Consumer');
      if (row === 2 && col === 0) return text('West');
      if (row === 2 && col === 1) return num(20);
      if (row === 2 && col === 2) return num(4);
      if (row === 2 && col === 3) return text('Retail');
      if (row === 2 && col === 4) return text('Business');
      return { kind: 'blank' } as CellValue;
    },
    createPivotCache: () => 9,
    removePivotCache: () => true,
    addPivotCacheField: (_cache: number, name: string) => {
      calls.push(`field:${name}`);
      return calls.filter((c) => c.startsWith('field:')).length - 1;
    },
    addPivotCacheSharedItem: () => true,
    addPivotCacheRecord: () => 0,
    setPivotCacheRecordValue: () => true,
    createPivotTable: () => {
      calls.push('pivot');
      return 0;
    },
    removePivotTable: () => true,
    setPivotTableGrandTotals: (_sheet: number, _pivot: number, rows: boolean, cols: boolean) => {
      calls.push(`grand:${rows}:${cols}`);
      return true;
    },
    addPivotField: (_sheet: number, _pivot: number, spec: { sourceName: string; axis: number }) => {
      calls.push(`pivot-field:${spec.sourceName}:${spec.axis}`);
      return 0;
    },
    addPivotFieldItem: (
      _sheet: number,
      _pivot: number,
      _fieldIdx: number,
      itemName: string,
      visible: boolean,
    ) => {
      calls.push(`pivot-item:${itemName}:${visible}`);
      return true;
    },
    addPivotFilter: (
      _sheet: number,
      _pivot: number,
      spec: {
        fieldName: string;
        type: number;
        valueDouble?: number;
        valueHighDouble?: number;
        valueInt?: number;
      },
    ) => {
      calls.push(
        [
          'pivot-filter',
          spec.fieldName,
          String(spec.type),
          spec.valueDouble === undefined ? '' : String(spec.valueDouble),
          spec.valueHighDouble === undefined ? '' : String(spec.valueHighDouble),
          spec.valueInt === undefined ? '' : String(spec.valueInt),
        ].join(':'),
      );
      return true;
    },
    setPivotFieldSort: () => true,
    addPivotDataField: (_sheet: number, _pivot: number, spec: { name: string }) => {
      calls.push(`data-field:${spec.name}`);
      return 0;
    },
  } as unknown as WorkbookHandle;
  return { wb, calls };
};

describe('attachPivotTableDialog', () => {
  let host: HTMLElement;

  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    document.body.replaceChildren();
  });

  it('creates a PivotTable from the selected range', () => {
    const { wb, calls } = makeWb();
    const store = createSpreadsheetStore();
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 });
    const handle = attachPivotTableDialog({ host, store, wb, strings: en });

    handle.open();
    expect(document.body.textContent).toContain('Create PivotTable');
    expect(document.querySelector<HTMLInputElement>('.fc-pivotdlg__field input')?.value).toBe(
      'A1:B3',
    );
    expect(document.querySelectorAll('.fc-pivotdlg__section')).toHaveLength(4);
    expect(document.querySelector('.fc-pivotdlg__field-list')?.textContent).toContain(
      'PivotTable Fields',
    );
    expect(document.querySelectorAll('[data-pivot-field-list-field]')).toHaveLength(2);
    expect(document.querySelector('.fc-pivotdlg__area-grid')?.textContent).toContain('Region');
    expect(document.querySelectorAll('.fc-pivotdlg__checkgrid .fc-pivotdlg__check')).toHaveLength(
      4,
    );

    const form = document.querySelector('form');
    expect(form).toBeTruthy();
    if (!form) throw new Error('missing PivotTable form');
    form.dispatchEvent(new SubmitEvent('submit', { bubbles: true, cancelable: true }));

    expect(calls).toContain('field:Region');
    expect(calls).toContain('field:Sales');
    expect(calls).toContain('pivot');
    expect(calls).toContain('grand:true:true');
    expect(document.querySelector('.fc-pivotdlg')?.hasAttribute('hidden')).toBe(true);
    handle.detach();
  });

  it('lets the field list update the PivotTable value area', () => {
    const { wb } = makeWb();
    const store = createSpreadsheetStore();
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 });
    const handle = attachPivotTableDialog({ host, store, wb, strings: en });

    handle.open();
    const salesField = document.querySelector<HTMLInputElement>(
      '[data-pivot-field-list-field="Sales"]',
    );
    const valueSelect = Array.from(document.querySelectorAll<HTMLSelectElement>('select'))[3];
    expect(salesField?.checked).toBe(true);
    expect(valueSelect?.value).toBe('Sales');

    if (!salesField || !valueSelect) throw new Error('missing field list controls');
    salesField.checked = false;
    salesField.dispatchEvent(new Event('change', { bubbles: true }));
    expect(valueSelect.value).toBe('Qty');

    const nextSalesField = document.querySelector<HTMLInputElement>(
      '[data-pivot-field-list-field="Sales"]',
    );
    if (!nextSalesField) throw new Error('missing updated field list controls');
    nextSalesField.checked = true;
    nextSalesField.dispatchEvent(new Event('change', { bubbles: true }));
    expect(valueSelect.value).toBe('Qty');
    expect(document.querySelector('.fc-pivotdlg__area-grid')?.textContent).toContain('Sales');
    handle.detach();
  });

  it('lets the field list submit multiple Values area fields', () => {
    const { wb, calls } = makeWb();
    const store = createSpreadsheetStore();
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 });
    const handle = attachPivotTableDialog({ host, store, wb, strings: en });

    handle.open();
    const qtyField = document.querySelector<HTMLInputElement>(
      '[data-pivot-field-list-field="Qty"]',
    );
    if (!qtyField) throw new Error('missing Qty field checkbox');
    qtyField.checked = true;
    qtyField.dispatchEvent(new Event('change', { bubbles: true }));

    const areaText = document.querySelector('.fc-pivotdlg__area-grid')?.textContent ?? '';
    expect(areaText).toContain('Sales');
    expect(areaText).toContain('Qty');
    document
      .querySelector('form')
      ?.dispatchEvent(new SubmitEvent('submit', { bubbles: true, cancelable: true }));

    expect(calls).toContain('data-field:Sum of Sales');
    expect(calls).toContain('data-field:Sum of Qty');
    handle.detach();
  });

  it('opens a shared Field Settings entry from an assigned Values field', () => {
    const { wb } = makeWb();
    const store = createSpreadsheetStore();
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 });
    const handle = attachPivotTableDialog({ host, store, wb, strings: en });

    handle.open();
    const settings = document.querySelector<HTMLButtonElement>(
      '.fc-pivotdlg__area-settings[aria-label="Field Settings: Sales"]',
    );
    if (!settings) throw new Error('missing value field settings button');
    settings.click();

    const panel = document.querySelector<HTMLElement>('.fc-pivotdlg__area-settings-panel');
    expect(panel?.hidden).toBe(false);
    expect(panel?.textContent).toContain('Field Settings: Sales');
    expect(panel?.textContent).toContain('summarize-by');
    expect(document.activeElement).toBe(panel?.querySelector('select'));
    handle.detach();
  });

  it('lets a Values field settings panel change aggregation before submit', () => {
    const { wb, calls } = makeWb();
    const store = createSpreadsheetStore();
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 });
    const handle = attachPivotTableDialog({ host, store, wb, strings: en });

    handle.open();
    document
      .querySelector<HTMLButtonElement>(
        '.fc-pivotdlg__area-settings[aria-label="Field Settings: Sales"]',
      )
      ?.click();
    const panel = document.querySelector<HTMLElement>('.fc-pivotdlg__area-settings-panel');
    const aggregation = panel?.querySelector<HTMLSelectElement>('select');
    if (!aggregation) throw new Error('missing field settings aggregation');
    aggregation.value = '1';
    aggregation.dispatchEvent(new Event('change', { bubbles: true }));
    document
      .querySelector('form')
      ?.dispatchEvent(new SubmitEvent('submit', { bubbles: true, cancelable: true }));

    expect(calls).toContain('data-field:Count of Sales');
    handle.detach();
  });

  it('lets the field list show and submit a Filters area field', () => {
    const { wb, calls } = makeWb();
    const store = createSpreadsheetStore();
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 });
    const handle = attachPivotTableDialog({ host, store, wb, strings: en });

    handle.open();
    const filterSelect = Array.from(document.querySelectorAll<HTMLSelectElement>('select'))[0];
    if (!filterSelect) throw new Error('missing filter select');
    filterSelect.value = 'Qty';
    filterSelect.dispatchEvent(new Event('change', { bubbles: true }));

    expect(document.querySelector('.fc-pivotdlg__area-grid')?.textContent).toContain('Qty');
    document
      .querySelector('form')
      ?.dispatchEvent(new SubmitEvent('submit', { bubbles: true, cancelable: true }));

    expect(calls).toContain('pivot');
    handle.detach();
  });

  it('lets a Filters field settings panel hide source items before submit', () => {
    const { wb, calls } = makeWb();
    const store = createSpreadsheetStore();
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 4 });
    const handle = attachPivotTableDialog({ host, store, wb, strings: en });

    handle.open();
    const colSelect = Array.from(document.querySelectorAll<HTMLSelectElement>('select'))[2];
    if (!colSelect) throw new Error('missing column select');
    colSelect.value = 'Qty';
    colSelect.dispatchEvent(new Event('change', { bubbles: true }));
    const channelField = document.querySelector<HTMLInputElement>(
      '[data-pivot-field-list-field="Channel"]',
    );
    if (!channelField) throw new Error('missing Channel field checkbox');
    channelField.checked = true;
    channelField.dispatchEvent(new Event('change', { bubbles: true }));

    document
      .querySelector<HTMLButtonElement>(
        '.fc-pivotdlg__area-settings[aria-label="Field Settings: Channel"]',
      )
      ?.click();
    const panel = document.querySelector<HTMLElement>('.fc-pivotdlg__area-settings-panel');
    expect(panel?.textContent).toContain('Filter items');
    const retailItem = Array.from(
      panel?.querySelectorAll<HTMLInputElement>('.fc-pivotdlg__settings-item-grid input') ?? [],
    ).find((input) => input.parentElement?.textContent?.includes('Retail'));
    if (!retailItem) throw new Error('missing Retail filter item');
    retailItem.checked = false;
    retailItem.dispatchEvent(new Event('change', { bubbles: true }));

    document
      .querySelector('form')
      ?.dispatchEvent(new SubmitEvent('submit', { bubbles: true, cancelable: true }));

    expect(calls).toContain('pivot-field:Channel:3');
    expect(calls).toContain('pivot-item:Retail:false');
    handle.detach();
  });

  it('lets a Filters field settings panel add a label condition before submit', () => {
    const { wb, calls } = makeWb();
    const store = createSpreadsheetStore();
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 4 });
    const handle = attachPivotTableDialog({ host, store, wb, strings: en });

    handle.open();
    const colSelect = Array.from(document.querySelectorAll<HTMLSelectElement>('select'))[2];
    if (!colSelect) throw new Error('missing column select');
    colSelect.value = 'Qty';
    colSelect.dispatchEvent(new Event('change', { bubbles: true }));
    const channelField = document.querySelector<HTMLInputElement>(
      '[data-pivot-field-list-field="Channel"]',
    );
    if (!channelField) throw new Error('missing Channel field checkbox');
    channelField.checked = true;
    channelField.dispatchEvent(new Event('change', { bubbles: true }));

    document
      .querySelector<HTMLButtonElement>(
        '.fc-pivotdlg__area-settings[aria-label="Field Settings: Channel"]',
      )
      ?.click();
    const panel = document.querySelector<HTMLElement>('.fc-pivotdlg__area-settings-panel');
    const conditionSelect = Array.from(
      panel?.querySelectorAll<HTMLSelectElement>('select') ?? [],
    ).find((select) =>
      Array.from(select.options).some((option) => option.value === 'label-contains'),
    );
    if (!conditionSelect) throw new Error('missing filter condition controls');
    conditionSelect.value = 'label-contains';
    conditionSelect.dispatchEvent(new Event('change', { bubbles: true }));
    const nextConditionValue = Array.from(
      panel?.querySelectorAll<HTMLInputElement>('input') ?? [],
    ).find((input) => input.placeholder === 'Text');
    if (!nextConditionValue) throw new Error('missing text condition value');
    nextConditionValue.value = 'Online';
    nextConditionValue.dispatchEvent(new Event('input', { bubbles: true }));

    document
      .querySelector('form')
      ?.dispatchEvent(new SubmitEvent('submit', { bubbles: true, cancelable: true }));

    expect(calls.some((call) => call.startsWith('pivot-filter:Channel:3:'))).toBe(true);
    handle.detach();
  });

  it('lets a Filters field settings panel add a value-between condition before submit', () => {
    const { wb, calls } = makeWb();
    const store = createSpreadsheetStore();
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 4 });
    const handle = attachPivotTableDialog({ host, store, wb, strings: en });

    handle.open();
    const colSelect = Array.from(document.querySelectorAll<HTMLSelectElement>('select'))[2];
    if (!colSelect) throw new Error('missing column select');
    colSelect.value = 'Qty';
    colSelect.dispatchEvent(new Event('change', { bubbles: true }));
    const channelField = document.querySelector<HTMLInputElement>(
      '[data-pivot-field-list-field="Channel"]',
    );
    if (!channelField) throw new Error('missing Channel field checkbox');
    channelField.checked = true;
    channelField.dispatchEvent(new Event('change', { bubbles: true }));

    document
      .querySelector<HTMLButtonElement>(
        '.fc-pivotdlg__area-settings[aria-label="Field Settings: Channel"]',
      )
      ?.click();
    const panel = document.querySelector<HTMLElement>('.fc-pivotdlg__area-settings-panel');
    const categorySelect = panel?.querySelector<HTMLSelectElement>(
      'select[data-pivot-filter-category="true"]',
    );
    if (!categorySelect) throw new Error('missing filter category controls');
    expect(categorySelect.textContent).toContain('Label Filters');
    expect(categorySelect.textContent).toContain('Value Filters');
    expect(categorySelect.textContent).toContain('Date Filters');
    categorySelect.value = 'value';
    categorySelect.dispatchEvent(new Event('change', { bubbles: true }));
    const conditionSelect = Array.from(
      panel?.querySelectorAll<HTMLSelectElement>('select') ?? [],
    ).find((select) =>
      Array.from(select.options).some((option) => option.value === 'value-between'),
    );
    if (!conditionSelect) throw new Error('missing filter condition controls');
    conditionSelect.value = 'value-between';
    conditionSelect.dispatchEvent(new Event('change', { bubbles: true }));
    const rangeInputs = Array.from(
      panel?.querySelectorAll<HTMLInputElement>('.fc-pivotdlg__settings-condition-values input') ??
        [],
    );
    const [low, high] = rangeInputs;
    if (!low || !high) throw new Error('missing between condition inputs');
    low.value = '10';
    low.dispatchEvent(new Event('input', { bubbles: true }));
    high.value = '20';
    high.dispatchEvent(new Event('input', { bubbles: true }));

    document
      .querySelector('form')
      ?.dispatchEvent(new SubmitEvent('submit', { bubbles: true, cancelable: true }));

    expect(calls).toContain('pivot-filter:Channel:2:10:20:');
    handle.detach();
  });

  it('lets fields move into an area by drag and drop', () => {
    const { wb, calls } = makeWb();
    const store = createSpreadsheetStore();
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 4 });
    const handle = attachPivotTableDialog({ host, store, wb, strings: en });

    handle.open();
    const source = document
      .querySelector<HTMLInputElement>('[data-pivot-field-list-field="Channel"]')
      ?.closest('.fc-pivotdlg__field-chip');
    const filtersArea = document.querySelector<HTMLElement>('[data-pivot-area="filters"]');
    if (!source || !filtersArea) throw new Error('missing drag source or target');
    const dataTransfer = makeDataTransfer();
    dispatchDragEvent(source, 'dragstart', dataTransfer);
    dispatchDragEvent(filtersArea, 'dragover', dataTransfer);
    dispatchDragEvent(filtersArea, 'drop', dataTransfer);

    expect(document.querySelector('.fc-pivotdlg__area-grid')?.textContent).toContain('Channel');
    document
      .querySelector('form')
      ?.dispatchEvent(new SubmitEvent('submit', { bubbles: true, cancelable: true }));

    expect(calls).toContain('pivot-field:Channel:3');
    handle.detach();
  });

  it('lets the field list submit multiple Filters area fields', () => {
    const { wb, calls } = makeWb();
    const store = createSpreadsheetStore();
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 4 });
    const handle = attachPivotTableDialog({ host, store, wb, strings: en });

    handle.open();
    const colSelect = Array.from(document.querySelectorAll<HTMLSelectElement>('select'))[2];
    if (!colSelect) throw new Error('missing column select');
    colSelect.value = 'Qty';
    colSelect.dispatchEvent(new Event('change', { bubbles: true }));
    for (const field of ['Channel', 'Segment']) {
      const checkbox = document.querySelector<HTMLInputElement>(
        `[data-pivot-field-list-field="${field}"]`,
      );
      if (!checkbox) throw new Error(`missing ${field} field checkbox`);
      checkbox.checked = true;
      checkbox.dispatchEvent(new Event('change', { bubbles: true }));
    }

    const areaText = document.querySelector('.fc-pivotdlg__area-grid')?.textContent ?? '';
    expect(areaText).toContain('Channel');
    expect(areaText).toContain('Segment');
    document
      .querySelector('form')
      ?.dispatchEvent(new SubmitEvent('submit', { bubbles: true, cancelable: true }));

    expect(calls).toContain('pivot-field:Channel:3');
    expect(calls).toContain('pivot-field:Segment:3');
    handle.detach();
  });

  it('rebuilds fields from an edited source range', () => {
    const { wb, calls } = makeWb();
    const store = createSpreadsheetStore();
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    const handle = attachPivotTableDialog({ host, store, wb, strings: en });

    handle.open();
    expect(document.querySelector('.fc-fmtdlg__btn--primary')?.hasAttribute('disabled')).toBe(true);
    const source = document.querySelector<HTMLInputElement>('.fc-pivotdlg__field input');
    expect(source).toBeTruthy();
    if (!source) throw new Error('missing source input');
    source.value = 'A1:B3';
    source.dispatchEvent(new Event('input', { bubbles: true }));
    expect(document.querySelector('.fc-fmtdlg__btn--primary')?.hasAttribute('disabled')).toBe(
      false,
    );

    const form = document.querySelector('form');
    expect(form).toBeTruthy();
    if (!form) throw new Error('missing PivotTable form');
    form.dispatchEvent(new SubmitEvent('submit', { bubbles: true, cancelable: true }));

    expect(calls).toContain('field:Region');
    expect(calls).toContain('field:Sales');
    expect(calls).toContain('pivot');
    handle.detach();
  });

  it('treats Enter as the default OK action when creation is available', () => {
    const { wb, calls } = makeWb();
    const store = createSpreadsheetStore();
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 });
    const handle = attachPivotTableDialog({ host, store, wb, strings: en });

    handle.open();
    const overlay = document.querySelector<HTMLElement>('.fc-pivotdlg');
    overlay?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));

    expect(calls).toContain('pivot');
    expect(overlay?.hasAttribute('hidden')).toBe(true);
    handle.detach();
  });

  it('shows a disabled-state message when mutation is unavailable', () => {
    const { wb } = makeWb();
    Object.defineProperty(wb, 'capabilities', { value: { pivotTableMutate: false } });
    const store = createSpreadsheetStore();
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 });
    const handle = attachPivotTableDialog({ host, store, wb, strings: en });

    handle.open();
    expect(document.body.textContent).toContain('does not support PivotTable creation');
    expect(document.querySelector('.fc-fmtdlg__btn--primary')?.hasAttribute('disabled')).toBe(true);
    expect(document.activeElement?.textContent).toBe('Cancel');
    handle.detach();
  });
});
