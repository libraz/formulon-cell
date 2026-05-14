import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  dispatchWorkbookObjectSummaries,
  hydrateActiveSheetFromEngine,
  hydrateWorkbookMetadataFromEngine,
} from '../../../src/mount/hydration.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

describe('mount/hydration', () => {
  let host: HTMLElement;
  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
  });
  afterEach(() => {
    host.remove();
  });

  it('hydrates the active sheet cells from the workbook into the store', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 42);
    wb.setNumber({ sheet: 0, row: 1, col: 0 }, 7);
    const store = createSpreadsheetStore();

    hydrateActiveSheetFromEngine(wb, store);

    const cells = store.getState().data.cells;
    const a1 = cells.get('0:0:0');
    const a2 = cells.get('0:1:0');
    expect(a1?.value).toEqual({ kind: 'number', value: 42 });
    expect(a2?.value).toEqual({ kind: 'number', value: 7 });
  });

  it('hydrates workbook-level metadata without touching cell data', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    const store = createSpreadsheetStore();
    const before = store.getState().data;

    hydrateWorkbookMetadataFromEngine(wb, store);

    const after = store.getState().data;
    expect(after.cells).toBe(before.cells); // untouched
  });

  it('dispatches passthrough + tables custom events on the host', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    const passthroughs = vi.fn();
    const tables = vi.fn();
    host.addEventListener('fc:passthroughs', passthroughs);
    host.addEventListener('fc:tables', tables);

    dispatchWorkbookObjectSummaries(host, wb);

    expect(passthroughs).toHaveBeenCalledTimes(1);
    expect(tables).toHaveBeenCalledTimes(1);
    const evt = passthroughs.mock.calls[0]?.[0] as CustomEvent;
    expect(evt.type).toBe('fc:passthroughs');
    expect(evt.detail).toBeDefined();
  });
});
