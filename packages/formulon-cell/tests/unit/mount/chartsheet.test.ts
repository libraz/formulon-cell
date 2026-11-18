import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { mutators } from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

const fixture = (): Uint8Array =>
  new Uint8Array(readFileSync(join(__dirname, '../../fixtures/charts/chart_chartsheet.xlsx')));

describe('an active chartsheet', () => {
  let sheet: MountedStubSheet | null = null;
  afterEach(() => {
    sheet?.dispose();
    sheet = null;
  });

  const mountFixture = async (): Promise<MountedStubSheet> => {
    const workbook = await WorkbookHandle.loadBytes(fixture());
    sheet = await mountStubSheet({ workbook });
    return sheet;
  };

  it('marks the host and draws its chart in place of the grid', async () => {
    const { host, instance } = await mountFixture();
    expect(instance.store.getState().data.sheetIndex).toBe(0);
    expect(host.dataset.fcSheetKind).toBe('chartsheet');
    const panels = host.querySelectorAll('.fc-charts .fc-chart');
    expect(panels).toHaveLength(1);
    // Its chart is fixed to the sheet: no sizing handles.
    expect(panels[0]?.classList.contains('fc-chart--fixed')).toBe(true);
  });

  it('draws the chart over the whole area of the grid', async () => {
    const { host, instance } = await mountFixture();
    const grid = host.querySelector<HTMLElement>('.fc-host__grid');
    if (!grid) throw new Error('missing grid');
    Object.defineProperty(grid, 'clientWidth', { value: 800, configurable: true });
    Object.defineProperty(grid, 'clientHeight', { value: 500, configurable: true });
    mutators.setSheetIndex(instance.store, 1);
    mutators.setSheetIndex(instance.store, 0);
    const panel = host.querySelector<HTMLElement>('.fc-charts .fc-chart');
    expect(panel?.style.width).toBe('800px');
    expect(panel?.style.height).toBe('500px');
    expect(panel?.style.left).toBe('0px');
    expect(panel?.style.top).toBe('0px');
  });

  it('shows nothing editable in the name box and the formula bar', async () => {
    const { host } = await mountFixture();
    const tag = host.querySelector<HTMLInputElement>('.fc-host__formulabar-tag');
    const fx = host.querySelector<HTMLTextAreaElement>('.fc-host__formulabar-input');
    expect(tag?.value).toBe('');
    expect(tag?.readOnly).toBe(true);
    expect(fx?.value).toBe('');
    expect(fx?.readOnly).toBe(true);
  });

  it('does not open the cell editor', async () => {
    const { instance } = await mountFixture();
    instance.cellEditor.begin('x');
    expect(instance.cellEditor.isActive()).toBe(false);
    mutators.setSheetIndex(instance.store, 1);
    instance.cellEditor.begin('x');
    expect(instance.cellEditor.isActive()).toBe(true);
  });

  it('is a grid again on the next worksheet tab', async () => {
    const { host, instance } = await mountFixture();
    mutators.setSheetIndex(instance.store, 1);
    expect(host.dataset.fcSheetKind).toBeUndefined();
    const tag = host.querySelector<HTMLInputElement>('.fc-host__formulabar-tag');
    expect(tag?.readOnly).toBe(false);
  });
});
