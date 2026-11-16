import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import {
  attachContextMenu,
  type ContextMenuHandle,
} from '../../../../src/interact/context-menu.js';
import { disposeOverlayPortal } from '../../../../src/interact/overlay-portal.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../src/store/store.js';
import { fireContextMenu, item, newWb, seed, setRange } from './fixtures.js';

describe('context menu over a table', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let detach: ContextMenuHandle | null;

  beforeEach(async () => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    wb = await newWb();
    detach = null;
    seed(store, wb, [{ row: 0, col: 0, value: 'h' }]);
    // Header row 0, data rows 1-4.
    mutators.upsertTableOverlay(store, {
      id: 't1',
      source: 'session',
      range: { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 2 },
      style: 'medium',
      showHeader: true,
      showTotal: false,
      banded: true,
    });
  });

  afterEach(() => {
    detach?.();
    disposeOverlayPortal(host);
    document.body.innerHTML = '';
    vi.restoreAllMocks();
  });

  it('greys out Delete on a row header that holds the table header row', () => {
    setRange(store, 0, 0, 0, 0);
    detach = attachContextMenu({ getLocale: () => 'en-US', host, store, wb });
    fireContextMenu(host, 10, 30); // row 0 header
    expect(item('rowDelete')?.disabled).toBe(true);
  });

  it('keeps Delete on for a data row header', () => {
    setRange(store, 1, 0, 1, 0);
    detach = attachContextMenu({ getLocale: () => 'en-US', host, store, wb });
    fireContextMenu(host, 10, 59); // row 1 header
    expect(item('rowDelete')).not.toBeNull();
    expect(item('rowDelete')?.disabled).toBe(false);
  });

  it('opens an Insert submenu instead of the flat Insert entry on a table cell', () => {
    setRange(store, 2, 1, 2, 1);
    detach = attachContextMenu({ getLocale: () => 'en-US', host, store, wb });
    fireContextMenu(host, 120, 90);
    expect(store.getState().selection.active).toMatchObject({ row: 2, col: 1 });
    expect(item('insertTableMenu')).not.toBeNull();
    expect(item('insertCells')).toBeNull();
  });

  it('keeps the flat Insert entry outside a table', () => {
    setRange(store, 9, 6, 9, 6);
    detach = attachContextMenu({ getLocale: () => 'en-US', host, store, wb });
    fireContextMenu(host, 120, 260);
    expect(item('insertCells')).not.toBeNull();
    expect(item('insertTableMenu')).toBeNull();
  });
});
