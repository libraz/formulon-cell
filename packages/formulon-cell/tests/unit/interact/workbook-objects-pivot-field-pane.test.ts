import { beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import { PivotAxis } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en, ja } from '../../../src/i18n/strings.js';
import { attachWorkbookObjectsPanel } from '../../../src/interact/workbook-objects.js';

const setup = async () => {
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
    valueFields: ['Sales'],
    valueField: 'Sales',
  });
  expect(created.ok).toBe(true);
  return wb;
};

const dragEvent = (type: string, data: Map<string, string>): DragEvent => {
  const event = new Event(type, { bubbles: true, cancelable: true }) as DragEvent;
  Object.defineProperty(event, 'dataTransfer', {
    value: {
      effectAllowed: 'move',
      dropEffect: 'none',
      setData: (key: string, value: string) => data.set(key, value),
      getData: (key: string) => data.get(key) ?? '',
    },
  });
  return event;
};

describe('PivotTable Fields pane', () => {
  let host: HTMLElement;
  beforeEach(() => {
    document.body.replaceChildren();
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  const open = (wb: WorkbookHandle, history?: History, strings = en) => {
    let edits = 0;
    const handle = attachWorkbookObjectsPanel({
      host,
      wb,
      strings,
      history: history ?? null,
      onAfterPivotEdit: () => {
        edits += 1;
      },
    });
    expect(handle.openPivotFieldList(0, 0)).toBe(true);
    return { handle, edits: () => edits };
  };
  const pane = (): HTMLElement => {
    const el = host.querySelector<HTMLElement>('.fc-objects__pivot-pane');
    if (!el) throw new Error('missing pane');
    return el;
  };
  const checkbox = (field: string): HTMLInputElement => {
    const el = pane().querySelector<HTMLInputElement>(`[data-pivot-field-list-field="${field}"]`);
    if (!el) throw new Error(`missing ${field} checkbox`);
    return el;
  };
  const area = (kind: string): HTMLElement => {
    const el = pane().querySelector<HTMLElement>(`[data-pivot-area="${kind}"]`);
    if (!el) throw new Error(`missing ${kind} area`);
    return el;
  };
  const chips = (kind: string): string[] =>
    Array.from(area(kind).querySelectorAll<HTMLElement>('[data-pivot-area-field]')).map(
      (chip) => chip.dataset.pivotAreaField ?? '',
    );
  const tick = (field: string, checked: boolean): void => {
    const box = checkbox(field);
    box.checked = checked;
    box.dispatchEvent(new Event('change', { bubbles: true }));
  };
  const drag = (source: Element, target: Element): void => {
    const data = new Map<string, string>();
    source.dispatchEvent(dragEvent('dragstart', data));
    target.dispatchEvent(dragEvent('dragover', data));
    target.dispatchEvent(dragEvent('drop', data));
    source.dispatchEvent(dragEvent('dragend', data));
  };

  it('shows the field list and four areas, with no meta line, selects or Apply', async () => {
    const wb = await setup();
    const { handle } = open(wb, undefined, ja);
    const root = host.querySelector<HTMLElement>('.fc-objects');
    expect(root?.classList.contains('fc-objects--taskpane')).toBe(true);
    expect(root?.querySelector('.fc-objects__title')?.textContent).toBe(
      'ピボットテーブルのフィールド',
    );
    expect(root?.querySelector('.fc-objects__heading')).toBeNull();
    expect(root?.textContent).not.toContain('R7C1');
    expect(pane().querySelector('select')).toBeNull();
    expect(pane().querySelector('form')).toBeNull();
    expect(root?.querySelector('[type="submit"]')).toBeNull();
    expect(pane().textContent).not.toContain('ピボットテーブルのフィールド');
    const labels = Array.from(pane().querySelectorAll('[data-pivot-area] > span')).map(
      (el) => el.textContent,
    );
    expect(labels).toEqual(['フィルター', '列', '行', '値']);
    expect(
      Array.from(pane().querySelectorAll<HTMLElement>('[data-pivot-area]')).map(
        (el) => el.dataset.pivotArea,
      ),
    ).toEqual(['filters', 'columns', 'rows', 'values']);
    const search = pane().querySelector<HTMLInputElement>('[data-pivot-field-search]');
    expect(search?.placeholder).toBe('フィールドの検索');
    expect(search?.closest('.fc-pivotdlg__field-search-row')?.textContent).toBe('フィールド名');
    expect(search?.parentElement?.querySelector('svg.fc-pivotdlg__field-search-icon')).toBeTruthy();
    expect(
      pane().querySelectorAll('[data-pivot-area] > span > svg.fc-pivotdlg__area-icon'),
    ).toHaveLength(4);
    expect(pane().textContent).not.toContain('レポートに追加するフィールドを選択');
    expect(pane().querySelector('.fc-pivotdlg__area-hint')?.textContent).toBe(
      'エリア間でフィールドをドラッグしてください',
    );
    expect(area('filters').querySelector('.fc-pivotdlg__area-fields')?.textContent).toBe('');
    expect(checkbox('Region').checked).toBe(true);
    expect(checkbox('Product').checked).toBe(false);
    expect(chips('rows')).toEqual(['Region']);
    expect(chips('values')).toEqual(['Sales']);
    handle.detach();
    wb.dispose();
  });

  it('puts a ticked text field at the top of Rows and a number field at the top of Values', async () => {
    const wb = await setup();
    const { handle, edits } = open(wb);
    tick('Product', true);
    expect(wb.getPivotDefinition(0, 0)?.rowFieldOrder).toEqual([1, 0]);
    expect(chips('rows')).toEqual(['Product', 'Region']);
    tick('Qty', true);
    expect(wb.getPivotDefinition(0, 0)?.dataFields.map((d) => d.fieldIndex)).toEqual([3, 2]);
    expect(chips('values')).toEqual(['Qty', 'Sales']);
    // Two value fields bring in the Values field, which the engine lays across the columns.
    expect(area('columns').querySelector('[data-pivot-area-fixed]')?.textContent).toBe('Values');
    expect(area('rows').querySelector('[data-pivot-area-fixed]')).toBeNull();
    tick('Sales', false);
    expect(area('columns').querySelector('[data-pivot-area-fixed]')).toBeNull();
    tick('Sales', true);
    expect(chips('values')).toEqual(['Sales', 'Qty']);
    expect(wb.getPivotDefinition(0, 0)?.dataFields.map((d) => d.fieldIndex)).toEqual([2, 3]);
    expect(edits()).toBe(4);
    tick('Product', false);
    expect(wb.getPivotDefinition(0, 0)?.fields[1]?.axis).toBeNull();
    handle.detach();
    wb.dispose();
  });

  it('moves a chip to Columns by drag and drop at once, and drops it out to remove it', async () => {
    const wb = await setup();
    const { handle } = open(wb);
    tick('Product', true);
    const product = area('rows').querySelector('[data-pivot-area-field="Product"]');
    if (!product) throw new Error('missing Product chip');
    drag(product, area('columns'));
    let definition = wb.getPivotDefinition(0, 0);
    expect(definition?.fields[1]?.axis).toBe(PivotAxis.Col);
    expect(definition?.colFieldOrder).toEqual([1]);
    expect(definition?.rowFieldOrder).toEqual([0]);
    expect(chips('columns')).toEqual(['Product']);

    const moved = area('columns').querySelector('[data-pivot-area-field="Product"]');
    if (!moved) throw new Error('missing moved chip');
    const data = new Map<string, string>();
    moved.dispatchEvent(dragEvent('dragstart', data));
    moved.dispatchEvent(dragEvent('dragend', data));
    definition = wb.getPivotDefinition(0, 0);
    expect(definition?.fields[1]?.axis).toBeNull();
    expect(chips('columns')).toEqual([]);
    expect(checkbox('Product').checked).toBe(false);
    handle.detach();
    wb.dispose();
  });

  it('reorders fields within an area by dropping on a chip', async () => {
    const wb = await setup();
    const { handle } = open(wb);
    tick('Product', true);
    const product = area('rows').querySelector('[data-pivot-area-field="Product"]');
    const region = area('rows').querySelector('[data-pivot-area-field="Region"]');
    if (!product || !region) throw new Error('missing row chips');
    drag(region, product);
    expect(wb.getPivotDefinition(0, 0)?.rowFieldOrder).toEqual([0, 1]);
    expect(chips('rows')).toEqual(['Region', 'Product']);

    tick('Qty', true);
    const qty = area('values').querySelector('[data-pivot-area-field="Qty"]');
    const sales = area('values').querySelector('[data-pivot-area-field="Sales"]');
    if (!qty || !sales) throw new Error('missing value chips');
    drag(sales, qty);
    expect(wb.getPivotDefinition(0, 0)?.dataFields.map((d) => d.fieldIndex)).toEqual([2, 3]);
    handle.detach();
    wb.dispose();
  });

  it('undoes and redoes each pane change as one step', async () => {
    const wb = await setup();
    const history = new History();
    const { handle } = open(wb, history);
    tick('Product', true);
    const product = area('rows').querySelector('[data-pivot-area-field="Product"]');
    if (!product) throw new Error('missing Product chip');
    drag(product, area('columns'));
    expect(wb.getPivotDefinition(0, 0)?.colFieldOrder).toEqual([1]);

    expect(history.undo()).toBe(true);
    expect(wb.getPivotDefinition(0, 0)?.rowFieldOrder).toEqual([1, 0]);
    expect(wb.getPivotDefinition(0, 0)?.colFieldOrder).toEqual([]);
    expect(chips('rows')).toEqual(['Product', 'Region']);
    expect(history.undo()).toBe(true);
    expect(wb.getPivotDefinition(0, 0)?.fields[1]?.axis).toBeNull();
    expect(wb.getPivotDefinition(0, 0)?.dataFields.map((d) => d.fieldIndex)).toEqual([2]);
    expect(checkbox('Product').checked).toBe(false);
    expect(history.redo()).toBe(true);
    expect(wb.getPivotDefinition(0, 0)?.rowFieldOrder).toEqual([1, 0]);
    handle.detach();
    wb.dispose();
  });

  it('opens the PivotTable Field dialog from a value or row chip, not an inline panel', async () => {
    const wb = await setup();
    const { handle } = open(wb);
    area('values').querySelector<HTMLElement>('[data-pivot-area-field="Sales"]')?.click();
    let dialog = document.querySelector<HTMLElement>('.fc-pvfsdlg');
    expect(dialog?.dataset.pvfsKind).toBe('value');
    expect(pane().querySelector('[data-pivot-aggregation-field-index]')).toBeNull();
    dialog?.querySelector<HTMLButtonElement>('[data-pvfs-cancel]')?.click();
    area('rows').querySelector<HTMLElement>('[data-pivot-area-field="Region"]')?.click();
    dialog = document.querySelector<HTMLElement>('.fc-pvfsdlg');
    expect(dialog?.dataset.pvfsKind).toBe('field');
    expect(pane().querySelector('[data-pivot-subtotal-mode-field-index]')).toBeNull();
    dialog?.querySelector<HTMLButtonElement>('[data-pvfs-cancel]')?.click();
    handle.detach();
    wb.dispose();
  });

  it('opens filter items from a Filters chip and hides an item at once', async () => {
    const wb = await setup();
    const { handle } = open(wb);
    const region = area('rows').querySelector('[data-pivot-area-field="Region"]');
    if (!region) throw new Error('missing Region chip');
    drag(region, area('filters'));
    expect(wb.getPivotDefinition(0, 0)?.fields[0]?.axis).toBe(PivotAxis.Page);
    area('filters').querySelector<HTMLElement>('[data-pivot-area-field="Region"]')?.click();
    const east = Array.from(
      pane().querySelectorAll<HTMLInputElement>(
        '[data-pivot-filter-checklist-field-index="0"] input',
      ),
    ).find((input) => input.value === 'East');
    if (!east) throw new Error('missing East item');
    east.checked = false;
    east.dispatchEvent(new Event('change', { bubbles: true }));
    const items = wb.getPivotDefinition(0, 0)?.fields[0]?.items ?? [];
    const eastIndex = Number(east.dataset.pivotItemCacheIndex);
    expect(items.find((item) => item.cacheIndex === eastIndex)?.visible).toBe(false);
    expect(items.find((item) => item.cacheIndex !== eastIndex)?.visible).toBe(true);
    handle.detach();
    wb.dispose();
  });

  it('filters the field list by the search text', async () => {
    const wb = await setup();
    const { handle } = open(wb);
    const search = pane().querySelector<HTMLInputElement>('[data-pivot-field-search]');
    if (!search) throw new Error('missing search box');
    search.value = 'pro';
    search.dispatchEvent(new Event('input', { bubbles: true }));
    const visible = Array.from(pane().querySelectorAll<HTMLElement>('.fc-pivotdlg__field-chip'))
      .filter((row) => !row.hidden)
      .map((row) => row.textContent);
    expect(visible).toEqual(['Product']);
    tick('Product', true);
    const kept = pane().querySelector<HTMLInputElement>('[data-pivot-field-search]');
    expect(kept?.value, 'search survives the re-render after a change').toBe('pro');
    expect(
      Array.from(pane().querySelectorAll<HTMLElement>('.fc-pivotdlg__field-chip'))
        .filter((row) => !row.hidden)
        .map((row) => row.textContent),
    ).toEqual(['Product']);
    handle.detach();
    wb.dispose();
  });

  it('closes on Escape', async () => {
    const wb = await setup();
    const { handle } = open(wb);
    host
      .querySelector('.fc-objects')
      ?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(handle.isPivotFieldListOpen()).toBe(false);
    handle.detach();
    wb.dispose();
  });

  it('docks in the host task-pane dock and leaves it on close', async () => {
    const wb = await setup();
    const dock = document.createElement('div');
    dock.hidden = true;
    host.appendChild(dock);
    const handle = attachWorkbookObjectsPanel({ host, wb, strings: en, dock });
    const root = host.querySelector<HTMLElement>('.fc-objects');
    handle.open();
    expect(root?.parentElement, 'the general panel floats over the host').toBe(host);
    expect(dock.hidden).toBe(true);
    expect(handle.openPivotFieldList(0, 0)).toBe(true);
    expect(root?.parentElement).toBe(dock);
    expect(dock.hidden).toBe(false);
    tick('Product', true);
    expect(root?.parentElement, 'a live edit keeps it docked').toBe(dock);
    handle.close();
    expect(root?.parentElement).toBe(host);
    expect(dock.hidden).toBe(true);
    handle.detach();
    wb.dispose();
  });

  it('resizes the field list from the splitter and keeps the size across edits', async () => {
    const wb = await setup();
    const { handle } = open(wb);
    const splitter = pane().querySelector<HTMLElement>('[data-pivot-pane-splitter]');
    if (!splitter) throw new Error('missing splitter');
    expect(splitter.getAttribute('role')).toBe('separator');
    expect(
      splitter.previousElementSibling?.classList.contains('fc-pivotdlg__field-list-available'),
    ).toBe(true);
    expect(splitter.nextElementSibling?.classList.contains('fc-pivotdlg__areas')).toBe(true);
    splitter.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    const grown = pane().style.getPropertyValue('--fc-pivot-pane-list-height');
    expect(grown).toMatch(/^\d+px$/);
    splitter.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    const taller = Number.parseInt(
      pane().style.getPropertyValue('--fc-pivot-pane-list-height'),
      10,
    );
    expect(taller).toBe(Number.parseInt(grown, 10) + 22);
    tick('Product', true);
    expect(pane().style.getPropertyValue('--fc-pivot-pane-list-height')).toBe(`${taller}px`);
    handle.detach();
    wb.dispose();
  });
});
