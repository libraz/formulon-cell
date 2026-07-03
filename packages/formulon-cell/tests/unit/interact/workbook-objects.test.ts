import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { beforeEach, describe, expect, it, vi } from 'vitest';
import {
  PivotAggregation,
  PivotAxis,
  type PivotDataFieldSpec,
  PivotFilterType,
  PivotFilterValueKind,
  PivotReportLayout,
} from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en, ja } from '../../../src/i18n/strings.js';
import {
  attachWorkbookObjectsPanel,
  buildSpreadsheetCompatibilityReport,
} from '../../../src/interact/workbook-objects.js';
import type { SessionIllustration } from '../../../src/store/store.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

const wbWithObjects = () => {
  const calls: string[] = [];
  const wb = {
    capabilities: {
      cellFormatting: true,
      conditionalFormatMutate: true,
      dataValidation: true,
      freeze: true,
      pivotTables: true,
      pivotTableMutate: true,
      pivotReportLayout: true,
      externalLinks: true,
    },
    getPassthroughs: () => [
      { path: 'xl/charts/chart1.xml' },
      { path: 'xl/drawings/drawing1.xml' },
      { path: 'xl/pivotTables/pivotTable1.xml' },
    ],
    getTables: () => [
      {
        sheetIndex: 0,
        name: 'Table1',
        displayName: 'Sales',
        ref: 'A1:C5',
        columns: ['Region', 'Sales', 'Margin'],
      },
      {
        sheetIndex: 1,
        name: 'Table2',
        displayName: '',
        ref: 'B2:B9',
        columns: ['Name'],
      },
    ],
    getPivotTables: () => [
      {
        sheetIndex: 0,
        pivotIndex: 0,
        top: 4,
        left: 1,
        rows: 6,
        cols: 3,
        cells: 18,
        fields: ['Region', 'Sales'],
        fieldItems: {
          Region: ['East', 'West'],
          Sales: ['10', '20'],
        },
      },
    ],
    renamePivotTable: (sheet: number, pivot: number, name: string) => {
      calls.push(`rename:${sheet}:${pivot}:${name}`);
      return true;
    },
    setPivotTableAnchor: (
      sheet: number,
      pivot: number,
      anchor: { row: number; col: number; rows: number; cols: number },
    ) => {
      calls.push(
        `anchor:${sheet}:${pivot}:${anchor.row}:${anchor.col}:${anchor.rows}:${anchor.cols}`,
      );
      return true;
    },
    setPivotTableGrandTotals: (
      sheet: number,
      pivot: number,
      rowsEnabled: boolean,
      colsEnabled: boolean,
    ) => {
      calls.push(`totals:${sheet}:${pivot}:${rowsEnabled}:${colsEnabled}`);
      return true;
    },
    getPivotReportLayout: () => PivotReportLayout.Compact,
    setPivotReportLayout: (sheet: number, pivot: number, layout: PivotReportLayout) => {
      calls.push(`layout:${sheet}:${pivot}:${layout}`);
      return true;
    },
    removePivotTable: (sheet: number, pivot: number) => {
      calls.push(`remove:${sheet}:${pivot}`);
      return true;
    },
    setPivotFieldAxis: (sheet: number, pivot: number, field: number, axis: number) => {
      calls.push(`field-axis:${sheet}:${pivot}:${field}:${axis}`);
      return true;
    },
    clearPivotFieldAggregations: (sheet: number, pivot: number, field: number) => {
      calls.push(`clear-agg:${sheet}:${pivot}:${field}`);
      return true;
    },
    addPivotFieldAggregation: (sheet: number, pivot: number, field: number, agg: number) => {
      calls.push(`add-agg:${sheet}:${pivot}:${field}:${agg}`);
      return true;
    },
    pivotDataFieldCount: (sheet: number, pivot: number) => {
      calls.push(`data-count:${sheet}:${pivot}`);
      return 1;
    },
    addPivotDataField: (sheet: number, pivot: number, spec: PivotDataFieldSpec) => {
      calls.push(
        `data-add:${sheet}:${pivot}:${spec.fieldIndex}:${spec.aggregation}:${spec.numberFormat ?? ''}`,
      );
      return 1;
    },
    setPivotDataField: (
      sheet: number,
      pivot: number,
      dataField: number,
      spec: PivotDataFieldSpec,
    ) => {
      calls.push(
        `data-set:${sheet}:${pivot}:${dataField}:${spec.fieldIndex}:${spec.aggregation}:${spec.numberFormat ?? ''}`,
      );
      return true;
    },
    setPivotFieldNumberFormat: (sheet: number, pivot: number, field: number, format: string) => {
      calls.push(`format:${sheet}:${pivot}:${field}:${format}`);
      return true;
    },
    clearPivotFieldItems: (sheet: number, pivot: number, field: number) => {
      calls.push(`clear-items:${sheet}:${pivot}:${field}`);
      return true;
    },
    addPivotFieldItem: (
      sheet: number,
      pivot: number,
      field: number,
      item: string,
      visible: boolean,
    ) => {
      calls.push(`add-item:${sheet}:${pivot}:${field}:${item}:${visible}`);
      return true;
    },
    clearPivotFilters: (sheet: number, pivot: number) => {
      calls.push(`clear-filters:${sheet}:${pivot}`);
      return true;
    },
    addPivotFilter: (
      sheet: number,
      pivot: number,
      spec: {
        fieldName: string;
        type: number;
        valueText?: string;
        valueDouble?: number;
        valueHighDouble?: number;
        valueInt?: number;
      },
    ) => {
      calls.push(
        [
          'add-filter',
          String(sheet),
          String(pivot),
          spec.fieldName,
          String(spec.type),
          spec.valueText ?? '',
          spec.valueDouble === undefined ? '' : String(spec.valueDouble),
          spec.valueHighDouble === undefined ? '' : String(spec.valueHighDouble),
          spec.valueInt === undefined ? '' : String(spec.valueInt),
        ].join(':'),
      );
      return true;
    },
  } as unknown as WorkbookHandle;
  return { wb, calls };
};

const emptyWb = () =>
  ({
    capabilities: {},
    getPassthroughs: () => [],
    getTables: () => [],
    getPivotTables: () => [],
  }) as unknown as WorkbookHandle;

describe('attachWorkbookObjectsPanel', () => {
  let host: HTMLElement;

  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  it('lists preserved OOXML categories, paths, and table names', () => {
    const { wb } = wbWithObjects();
    const handle = attachWorkbookObjectsPanel({ host, wb, strings: en });
    handle.open();

    expect(host.textContent).toContain('Workbook Objects');
    expect(host.textContent).toContain('Charts');
    expect(host.textContent).toContain('Drawings');
    expect(host.textContent).toContain('PivotTables');
    expect(host.textContent).toContain('xl/charts/chart1.xml');
    expect(host.textContent).toContain('Sales');
    expect(host.textContent).toContain('Table2');
    expect(host.textContent).toContain('A1:C5');
    expect(host.textContent).toContain('3 columns');
    expect(host.textContent).toContain('B2:B9');
    expect(host.textContent).toContain('1 column');
    expect(host.textContent).toContain('Pivot 1');
    expect(host.textContent).toContain('R5C2');
    expect(host.textContent).toContain('6 x 3');
    expect(host.textContent).toContain('Region');
    expect(host.textContent).toContain('Sales');
    expect(host.textContent).toContain('Spreadsheet compatibility');
    expect(host.textContent).toContain('PivotTable authoring');
    expect(host.textContent).toContain('Writable');
    expect(host.textContent).toContain('Unsupported');
    handle.detach();
  });

  it('uses current compatibility details in workbook object reports', () => {
    const { wb } = wbWithObjects();
    const report = buildSpreadsheetCompatibilityReport(wb, en.workbookObjects);

    expect(report.some((entry) => entry.detail.includes('hidden dropdown-arrow state'))).toBe(true);
    expect(report.some((entry) => entry.detail.includes('sheet-scoped defined names'))).toBe(true);
    expect(report.some((entry) => entry.detail.includes('comments on blank cells'))).toBe(true);
  });

  it('lists session illustrations supplied by the host feature', () => {
    const handle = attachWorkbookObjectsPanel({
      host,
      wb: emptyWb(),
      strings: en,
      listSessionIllustrations: () => [
        {
          id: 'ribbon-image-0',
          kind: 'image',
          sheet: 0,
          src: 'https://example.test/picture.png',
          alt: 'picture.png',
        },
        {
          id: 'ribbon-shape-0',
          kind: 'shape',
          shape: 'arrow',
          sheet: 1,
        },
      ],
    });
    handle.open();

    expect(host.textContent).toContain('Illustrations');
    expect(host.textContent).toContain('Pictures · Sheet 1 · ribbon-image-0');
    expect(host.textContent).toContain('arrow · Sheet 2 · ribbon-shape-0');
    expect(host.textContent).not.toContain('No preserved charts');
    handle.detach();
  });

  it('updates editable session shape appearance from the illustrations section', () => {
    const onUpdateSessionIllustration = vi.fn();
    const handle = attachWorkbookObjectsPanel({
      host,
      wb: emptyWb(),
      strings: en,
      listSessionIllustrations: () => [
        {
          id: 'shape-a',
          kind: 'shape',
          shape: 'rounded-rectangle',
          sheet: 0,
          color: '#0f6cbd',
          radius: 8,
          lineWidth: 4,
          opacity: 0.25,
        },
      ],
      onUpdateSessionIllustration,
    });
    handle.open();

    const color = host.querySelector<HTMLInputElement>('input[type="color"]');
    const numberInputs = Array.from(
      host.querySelectorAll<HTMLInputElement>('input[type="number"]'),
    );
    const radius = numberInputs[0];
    const lineWidth = numberInputs[1];
    const opacity = host.querySelector<HTMLInputElement>('input[type="range"]');
    expect(color?.value).toBe('#0f6cbd');
    expect(radius?.value).toBe('8');
    expect(lineWidth?.value).toBe('4');
    expect(opacity?.value).toBe('0.25');
    expect(color).not.toBeNull();
    expect(radius).not.toBeNull();
    expect(lineWidth).not.toBeNull();
    expect(opacity).not.toBeNull();
    if (!color || !radius || !lineWidth || !opacity) {
      throw new Error('Expected editable shape controls');
    }
    color.value = '#ff0000';
    radius.value = '18';
    lineWidth.value = '7';
    opacity.value = '0.45';
    radius.form?.dispatchEvent(new Event('submit', { bubbles: true, cancelable: true }));

    expect(onUpdateSessionIllustration).toHaveBeenCalledWith('shape-a', {
      color: '#ff0000',
      radius: 18,
      lineWidth: 7,
      opacity: 0.45,
    });
    handle.detach();
  });

  it('routes session illustration select, duplicate, and delete actions from the list', () => {
    const onSelectSessionIllustration = vi.fn();
    const onDuplicateSessionIllustration = vi.fn();
    const onClearSessionIllustration = vi.fn();
    const handle = attachWorkbookObjectsPanel({
      host,
      wb: emptyWb(),
      strings: en,
      listSessionIllustrations: () => [
        {
          id: 'shape-a',
          kind: 'shape',
          shape: 'rectangle',
          sheet: 0,
        },
      ],
      onSelectSessionIllustration,
      onDuplicateSessionIllustration,
      onClearSessionIllustration,
    });
    handle.open();

    const buttons = Array.from(host.querySelectorAll<HTMLButtonElement>('.fc-objects__action'));
    buttons.find((button) => button.textContent === 'Select')?.click();
    buttons.find((button) => button.textContent === 'Duplicate')?.click();
    buttons.find((button) => button.textContent === 'Delete')?.click();

    expect(onSelectSessionIllustration).toHaveBeenCalledWith('shape-a');
    expect(onDuplicateSessionIllustration).toHaveBeenCalledWith('shape-a');
    expect(onClearSessionIllustration).toHaveBeenCalledWith('shape-a');
    handle.detach();
  });

  it('refreshes open session illustration details from the shared subscription', () => {
    let illustrations: SessionIllustration[] = [];
    const listeners = new Set<() => void>();
    const handle = attachWorkbookObjectsPanel({
      host,
      wb: emptyWb(),
      strings: en,
      listSessionIllustrations: () => illustrations,
      subscribeSessionObjects: (listener) => {
        listeners.add(listener);
        return () => listeners.delete(listener);
      },
    });
    handle.open();
    expect(host.textContent).toContain('No preserved charts');

    illustrations = [
      {
        id: 'shape-later',
        kind: 'shape',
        shape: 'oval',
        sheet: 0,
      },
    ];
    for (const listener of listeners) listener();
    expect(host.textContent).toContain('oval · Sheet 1 · shape-later');

    handle.detach();
    illustrations = [
      {
        id: 'image-after-detach',
        kind: 'image',
        sheet: 0,
        src: 'https://example.test/after.png',
      },
    ];
    for (const listener of listeners) listener();
    expect(host.textContent).not.toContain('image-after-detach');
    expect(listeners.size).toBe(0);
  });

  it('opens the PivotTable creation flow from pivot details when supplied', () => {
    let opened = 0;
    const handle = attachWorkbookObjectsPanel({
      host,
      wb: wbWithObjects().wb,
      strings: en,
      onOpenPivotTableDialog: () => {
        opened += 1;
      },
    });
    handle.open();

    const button = Array.from(host.querySelectorAll<HTMLButtonElement>('button')).find((el) =>
      el.textContent?.includes('Create PivotTable'),
    );
    expect(button).toBeTruthy();
    button?.click();
    expect(opened).toBe(1);
    handle.detach();
  });

  it('edits a projected PivotTable through workbook object details', () => {
    const { wb, calls } = wbWithObjects();
    const onAfterPivotEdit = vi.fn();
    const handle = attachWorkbookObjectsPanel({
      host,
      wb,
      strings: en,
      onAfterPivotEdit,
    });
    handle.open();

    const edit = Array.from(host.querySelectorAll<HTMLButtonElement>('button')).find((el) =>
      el.textContent?.includes('Edit'),
    );
    edit?.click();
    const form = host.querySelector<HTMLFormElement>('.fc-objects__pivot-edit');
    const inputs = Array.from(form?.querySelectorAll<HTMLInputElement>('input') ?? []);
    const name = inputs.find((input) => input.type === 'text' && input.value.includes('Pivot'));
    const anchor = inputs.find((input) => input.type === 'text' && input.value === 'B5');
    if (!form || !name || !anchor) throw new Error('missing pivot edit form');
    name.value = 'Sales Pivot';
    anchor.value = 'C7';
    form.dispatchEvent(new SubmitEvent('submit', { bubbles: true, cancelable: true }));

    expect(calls).toContain('rename:0:0:Sales Pivot');
    expect(calls).toContain('anchor:0:0:6:2:6:3');
    expect(calls).toContain('totals:0:0:true:true');
    expect(calls).toContain(`layout:0:0:${PivotReportLayout.Compact}`);
    expect(calls).toContain('field-axis:0:0:0:0');
    expect(calls).toContain('field-axis:0:0:1:2');
    expect(calls).toContain(`data-set:0:0:0:1:${PivotAggregation.Sum}:`);
    expect(calls).not.toContain('clear-agg:0:0:1');
    expect(calls).not.toContain('add-agg:0:0:1:0');
    expect(onAfterPivotEdit).toHaveBeenCalledTimes(1);
    handle.detach();
  });

  it('updates the PivotTable report layout from workbook object details', () => {
    const { wb, calls } = wbWithObjects();
    const handle = attachWorkbookObjectsPanel({ host, wb, strings: en });
    handle.open();

    const edit = Array.from(host.querySelectorAll<HTMLButtonElement>('button')).find((el) =>
      el.textContent?.includes('Edit'),
    );
    edit?.click();
    const form = host.querySelector<HTMLFormElement>('.fc-objects__pivot-edit');
    const layout = Array.from(form?.querySelectorAll<HTMLSelectElement>('select') ?? []).find(
      (select) =>
        Array.from(select.options).some((option) => option.textContent === 'Tabular form'),
    );
    if (!form || !layout) throw new Error('missing pivot layout controls');
    layout.value = String(PivotReportLayout.Tabular);
    form.dispatchEvent(new SubmitEvent('submit', { bubbles: true, cancelable: true }));

    expect(calls).toContain(`layout:0:0:${PivotReportLayout.Tabular}`);
    handle.detach();
  });

  it('edits existing PivotTable field axis assignments', () => {
    const { wb, calls } = wbWithObjects();
    const handle = attachWorkbookObjectsPanel({ host, wb, strings: en });
    handle.open();
    const edit = Array.from(host.querySelectorAll<HTMLButtonElement>('button')).find((el) =>
      el.textContent?.includes('Edit'),
    );
    edit?.click();
    const form = host.querySelector<HTMLFormElement>('.fc-objects__pivot-edit');
    const fieldSelects = Array.from(
      form?.querySelectorAll<HTMLSelectElement>('[data-pivot-field-index]') ?? [],
    );
    expect(fieldSelects).toHaveLength(2);
    const [regionField, salesField] = fieldSelects as [HTMLSelectElement, HTMLSelectElement];
    expect(regionField.classList.contains('fc-objects__input')).toBe(true);
    expect(Array.from(regionField.options, (option) => [option.value, option.textContent])).toEqual(
      [
        ['0', 'Rows'],
        ['1', 'Columns'],
        ['3', 'Filters'],
        ['2', 'Values'],
      ],
    );
    const salesAggregation = form?.querySelector<HTMLSelectElement>(
      '[data-pivot-aggregation-field-index="1"]',
    );
    expect(salesAggregation?.classList.contains('fc-objects__input')).toBe(true);
    expect(
      Array.from(salesAggregation?.options ?? [], (option) => [option.value, option.textContent]),
    ).toEqual([
      ['0', 'Sum'],
      ['1', 'Count'],
      ['2', 'Average'],
      ['3', 'Max'],
      ['4', 'Min'],
    ]);
    regionField.value = '3';
    salesField.value = '1';
    form?.dispatchEvent(new SubmitEvent('submit', { bubbles: true, cancelable: true }));

    expect(calls).toContain('field-axis:0:0:0:3');
    expect(calls).toContain('field-axis:0:0:1:1');
    expect(calls).toContain('clear-items:0:0:0');
    expect(calls).not.toContain('add-agg:0:0:0:0');
    expect(calls).not.toContain('add-agg:0:0:1:0');
    handle.detach();
  });

  it('opens an existing PivotTable Field List without the general edit controls', () => {
    const { wb, calls } = wbWithObjects();
    const handle = attachWorkbookObjectsPanel({ host, wb, strings: en });
    handle.open();
    const fieldList = Array.from(host.querySelectorAll<HTMLButtonElement>('button')).find((el) =>
      el.textContent?.includes('PivotTable Fields'),
    );
    fieldList?.click();
    const form = host.querySelector<HTMLFormElement>('.fc-objects__pivot-edit');
    if (!form) throw new Error('missing field list form');
    expect(host.querySelector('.fc-objects')?.classList.contains('fc-objects--taskpane')).toBe(
      true,
    );
    expect(host.textContent).toContain('Back to Workbook Objects');
    expect(form.getAttribute('aria-label')).toBe('PivotTable Fields');
    expect(form.textContent).toContain('Choose fields to add to report');
    expect(form.textContent).toContain('Region');
    expect(form.textContent).toContain('Sales');
    expect(form.textContent).toContain('East');
    expect(form.textContent).toContain('West');
    expect(form.textContent).not.toContain('Anchor cell');
    const fieldListChecks = Array.from(
      form.querySelectorAll<HTMLInputElement>(
        '.fc-objects__pivot-field-list > .fc-objects__pivot-field-list-item > input',
      ),
    );
    expect(fieldListChecks).toHaveLength(2);
    for (const check of fieldListChecks) {
      expect(check.disabled).toBe(true);
      expect(check.dataset.disabledReason).toBe(en.workbookObjects.pivotFieldListCheckboxReadOnly);
      expect(check.getAttribute('aria-description')).toBe(
        en.workbookObjects.pivotFieldListCheckboxReadOnly,
      );
      expect(check.title).toBe(en.workbookObjects.pivotFieldListCheckboxReadOnly);
    }

    const regionAxis = form.querySelector<HTMLSelectElement>('[data-pivot-field-index="0"]');
    if (!regionAxis) throw new Error('missing region field axis');
    regionAxis.value = '3';
    regionAxis.dispatchEvent(new Event('change', { bubbles: true }));
    const eastItem = Array.from(
      form.querySelectorAll<HTMLInputElement>(
        '[data-pivot-filter-checklist-field-index="0"] input',
      ),
    ).find((input) => input.value === 'East');
    if (!eastItem) throw new Error('missing East filter item');
    eastItem.checked = false;
    form.dispatchEvent(new SubmitEvent('submit', { bubbles: true, cancelable: true }));

    expect(calls).toContain('field-axis:0:0:0:3');
    expect(calls).toContain('clear-items:0:0:0');
    expect(calls).toContain('add-item:0:0:0:East:false');
    expect(calls).toContain('add-item:0:0:0:West:true');
    expect(calls.some((call) => call.startsWith('rename:'))).toBe(false);
    expect(calls.some((call) => call.startsWith('anchor:'))).toBe(false);
    handle.detach();
  });

  it('opens an existing PivotTable Field List through the shared panel handle', () => {
    const { wb } = wbWithObjects();
    const handle = attachWorkbookObjectsPanel({ host, wb, strings: en });

    expect(handle.openPivotFieldList(0, 0)).toBe(true);
    expect(host.querySelector('.fc-objects')?.classList.contains('fc-objects--taskpane')).toBe(
      true,
    );
    expect(host.querySelector<HTMLFormElement>('.fc-objects__pivot-edit')?.textContent).toContain(
      'Choose fields to add to report',
    );
    expect(handle.openPivotFieldList(2, 0)).toBe(false);
    handle.detach();
  });

  it('edits existing PivotTable value aggregation and number format', () => {
    const { wb, calls } = wbWithObjects();
    const handle = attachWorkbookObjectsPanel({ host, wb, strings: en });
    handle.open();
    const edit = Array.from(host.querySelectorAll<HTMLButtonElement>('button')).find((el) =>
      el.textContent?.includes('Edit'),
    );
    edit?.click();
    const form = host.querySelector<HTMLFormElement>('.fc-objects__pivot-edit');
    const salesAxis = form?.querySelector<HTMLSelectElement>('[data-pivot-field-index="1"]');
    const aggregation = form?.querySelector<HTMLSelectElement>(
      '[data-pivot-aggregation-field-index="1"]',
    );
    const format = form?.querySelector<HTMLInputElement>(
      '[data-pivot-number-format-field-index="1"]',
    );
    if (!salesAxis || !aggregation || !format) throw new Error('missing value controls');
    salesAxis.value = '2';
    aggregation.value = '2';
    format.value = '#,##0.00';
    form?.dispatchEvent(new SubmitEvent('submit', { bubbles: true, cancelable: true }));

    expect(calls).toContain(`data-set:0:0:0:1:${PivotAggregation.Average}:#,##0.00`);
    expect(calls).not.toContain('clear-agg:0:0:1');
    expect(calls).not.toContain('add-agg:0:0:1:2');
    expect(calls).not.toContain('format:0:0:1:#,##0.00');
    handle.detach();
  });

  it('edits existing PivotTable filter item visibility list', () => {
    const { wb, calls } = wbWithObjects();
    const handle = attachWorkbookObjectsPanel({ host, wb, strings: en });
    handle.open();
    const edit = Array.from(host.querySelectorAll<HTMLButtonElement>('button')).find((el) =>
      el.textContent?.includes('Edit'),
    );
    edit?.click();
    const form = host.querySelector<HTMLFormElement>('.fc-objects__pivot-edit');
    const regionAxis = form?.querySelector<HTMLSelectElement>('[data-pivot-field-index="0"]');
    if (!regionAxis || !form) throw new Error('missing filter item controls');
    regionAxis.value = '3';
    regionAxis.dispatchEvent(new Event('change', { bubbles: true }));
    const westItem = Array.from(
      form.querySelectorAll<HTMLInputElement>(
        '[data-pivot-filter-checklist-field-index="0"] input',
      ),
    ).find((input) => input.value === 'West');
    if (!westItem) throw new Error('missing West filter item');
    westItem.checked = false;
    form.dispatchEvent(new SubmitEvent('submit', { bubbles: true, cancelable: true }));

    expect(calls).toContain('clear-items:0:0:0');
    expect(calls).toContain('add-item:0:0:0:East:true');
    expect(calls).toContain('add-item:0:0:0:West:false');
    expect(calls).not.toContain('clear-filters:0:0');
    handle.detach();
  });

  it('edits existing PivotTable filter conditions with the shared condition model', () => {
    const { wb, calls } = wbWithObjects();
    const handle = attachWorkbookObjectsPanel({ host, wb, strings: en });
    handle.open();
    const edit = Array.from(host.querySelectorAll<HTMLButtonElement>('button')).find((el) =>
      el.textContent?.includes('Edit'),
    );
    edit?.click();
    const form = host.querySelector<HTMLFormElement>('.fc-objects__pivot-edit');
    const regionAxis = form?.querySelector<HTMLSelectElement>('[data-pivot-field-index="0"]');
    const category = form?.querySelector<HTMLSelectElement>(
      '[data-pivot-filter-category-field-index="0"]',
    );
    const condition = form?.querySelector<HTMLSelectElement>(
      '[data-pivot-filter-condition-field-index="0"]',
    );
    if (!form || !regionAxis || !category || !condition) {
      throw new Error('missing filter condition controls');
    }
    regionAxis.value = '3';
    regionAxis.dispatchEvent(new Event('change', { bubbles: true }));
    expect(category.textContent).toContain('Label Filters');
    expect(category.textContent).toContain('Value Filters');
    category.value = 'label';
    category.dispatchEvent(new Event('change', { bubbles: true }));
    condition.value = 'label-contains';
    condition.dispatchEvent(new Event('change', { bubbles: true }));
    const value = form.querySelector<HTMLInputElement>(
      '.fc-objects__pivot-filter-condition-values input',
    );
    if (!value) throw new Error('missing filter condition value');
    value.value = 'East';
    value.dispatchEvent(new Event('input', { bubbles: true }));
    form.dispatchEvent(new SubmitEvent('submit', { bubbles: true, cancelable: true }));

    expect(calls).toContain('clear-filters:0:0');
    expect(calls.some((call) => call.startsWith('add-filter:0:0:Region:3:East:'))).toBe(true);
    handle.detach();
  });

  it('opens the shared PivotTable filter dialog from existing PivotTable edit', async () => {
    const { wb } = wbWithObjects();
    const handle = attachWorkbookObjectsPanel({ host, wb, strings: en });
    handle.open();
    const edit = Array.from(host.querySelectorAll<HTMLButtonElement>('button')).find((el) =>
      el.textContent?.includes('Edit'),
    );
    edit?.click();
    const form = host.querySelector<HTMLFormElement>('.fc-objects__pivot-edit');
    const regionAxis = form?.querySelector<HTMLSelectElement>('[data-pivot-field-index="0"]');
    if (!form || !regionAxis) throw new Error('missing pivot edit controls');
    regionAxis.value = '3';
    regionAxis.dispatchEvent(new Event('change', { bubbles: true }));
    const filterDialogButton = Array.from(form.querySelectorAll<HTMLButtonElement>('button')).find(
      (button) => button.textContent === 'Filter...',
    );
    expect(filterDialogButton).toBeTruthy();
    filterDialogButton?.click();

    const dialog = document.body.querySelector<HTMLElement>(
      '.fc-pivotdlg[role="dialog"]:not([hidden])',
    );
    expect(dialog?.textContent).toContain('PivotTable Filter: Region');
    const condition = dialog?.querySelector<HTMLSelectElement>(
      'select[data-pivot-filter-condition="true"]',
    );
    expect(condition?.value).toBe('none');
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn')?.click();
    await Promise.resolve();
    expect(document.body.textContent).not.toContain('PivotTable Filter: Region');
    handle.detach();
  });

  it('hydrates existing PivotTable filters into the shared condition editor', () => {
    const { wb } = wbWithObjects();
    wb.getPivotTables = () => [
      {
        sheetIndex: 0,
        pivotIndex: 0,
        top: 4,
        left: 1,
        rows: 6,
        cols: 3,
        cells: 18,
        fields: ['Region', 'Sales'],
        fieldItems: {
          Region: ['East', 'West'],
          Sales: ['-10', '20'],
        },
        pivotFilters: [
          {
            axis: PivotAxis.Page,
            fieldName: 'Sales',
            type: PivotFilterType.ValueBetween,
            valueKind: PivotFilterValueKind.Double,
            valueDouble: -10,
            valueHighKind: PivotFilterValueKind.Double,
            valueHighDouble: 20,
          },
        ],
      },
    ];
    const handle = attachWorkbookObjectsPanel({ host, wb, strings: en });
    handle.open();
    const edit = Array.from(host.querySelectorAll<HTMLButtonElement>('button')).find((el) =>
      el.textContent?.includes('Edit'),
    );
    edit?.click();

    const axis = host.querySelector<HTMLSelectElement>('[data-pivot-field-index="1"]');
    const category = host.querySelector<HTMLSelectElement>(
      '[data-pivot-filter-category-field-index="1"]',
    );
    const condition = host.querySelector<HTMLSelectElement>(
      '[data-pivot-filter-condition-field-index="1"]',
    );
    const valueContainer = condition?.closest('label')?.nextElementSibling;
    const values = Array.from(valueContainer?.querySelectorAll<HTMLInputElement>('input') ?? []);
    expect(axis?.value).toBe(String(PivotAxis.Page));
    expect(category?.value).toBe('value');
    expect(condition?.value).toBe('value-between');
    expect(values.map((input) => input.value)).toEqual(['-10', '20']);
    handle.detach();
  });

  it('refreshes when locale or workbook changes while open', () => {
    const handle = attachWorkbookObjectsPanel({ host, wb: emptyWb(), strings: en });
    handle.open();
    expect(host.textContent).toContain('No preserved charts');

    handle.setStrings(ja);
    expect(host.textContent).toContain('保持されたグラフ');
    expect(host.textContent).toContain('スプレッドシート互換');

    handle.bindWorkbook(wbWithObjects().wb);
    expect(host.textContent).toContain('グラフ');
    expect(host.textContent).toContain('シート 1');
    expect(host.textContent).toContain('3 列');
    expect(host.textContent).toContain('ピボット 1');
    expect(host.textContent).toContain('xl/pivotTables/pivotTable1.xml');
    handle.detach();
  });

  it('Escape closes the panel and restores focus to the opener', () => {
    host.tabIndex = -1;
    const handle = attachWorkbookObjectsPanel({ host, wb: emptyWb(), strings: en });
    host.focus();
    handle.open();
    const root = host.querySelector<HTMLElement>('.fc-objects');
    expect(document.activeElement).toBe(root);

    root?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));

    expect(root?.hidden).toBe(true);
    expect(document.activeElement).toBe(host);
    handle.detach();
  });

  it('keeps workbook object action buttons on the shared dialog primitive', () => {
    const source = readFileSync(join(root, 'src/interact/workbook-objects.ts'), 'utf8');
    expect(source).toContain('function createWorkbookObjectsActionButton(');
    expect(source).toContain(
      "createDialogButton({\n    label,\n    baseClass: 'fc-objects__action'",
    );
    expect(source).not.toContain("label: '×'");
    expect(source).not.toContain("document.createElement('button')");
  });

  it('keeps workbook objects panel on compact desktop pane chrome', () => {
    const css = readFileSync(
      join(root, 'src/styles/core/app/overlays/workbook-objects.css'),
      'utf8',
    );

    expect(css).toMatch(/\.fc-objects\s*\{[\s\S]*?border-radius: 2px;/);
    expect(css).toMatch(/\.fc-objects__close\s*\{[\s\S]*?border-radius: 2px;/);
    expect(css).toMatch(
      /\.fc-objects__close::before,[\s\S]*?\.fc-objects__close::after\s*\{[\s\S]*?background: currentColor;[\s\S]*?content: "";/,
    );
    expect(css).toMatch(/\.fc-objects__close:hover\s*\{[\s\S]*?background: var\(--fc-bg-hover/);
    expect(css).not.toContain('background: var(--fc-accent-soft');
  });
});
