import { beforeEach, describe, expect, it } from 'vitest';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en, ja } from '../../../src/i18n/strings.js';
import { attachWorkbookObjectsPanel } from '../../../src/interact/workbook-objects.js';

const wbWithObjects = () =>
  ({
    capabilities: {
      cellFormatting: true,
      conditionalFormatMutate: true,
      dataValidation: true,
      freeze: true,
      pivotTables: true,
      pivotTableMutate: true,
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
      },
    ],
  }) as unknown as WorkbookHandle;

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
    const handle = attachWorkbookObjectsPanel({ host, wb: wbWithObjects(), strings: en });
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
    expect(host.textContent).toContain('Region, Sales');
    expect(host.textContent).toContain('Spreadsheet compatibility');
    expect(host.textContent).toContain('PivotTable authoring');
    expect(host.textContent).toContain('Writable');
    expect(host.textContent).toContain('Unsupported');
    handle.detach();
  });

  it('refreshes when locale or workbook changes while open', () => {
    const handle = attachWorkbookObjectsPanel({ host, wb: emptyWb(), strings: en });
    handle.open();
    expect(host.textContent).toContain('No preserved charts');

    handle.setStrings(ja);
    expect(host.textContent).toContain('保持されたグラフ');
    expect(host.textContent).toContain('スプレッドシート互換');

    handle.bindWorkbook(wbWithObjects());
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
});
