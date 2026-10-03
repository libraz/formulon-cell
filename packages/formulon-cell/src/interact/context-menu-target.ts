import { MAX_COL, MAX_ROW } from '../engine/address.js';
import type { Addr } from '../engine/types.js';
import { hitZone, layoutForView } from '../render/geometry.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import type { MenuKind } from './context-menu-spec.js';

export type MenuTarget = { kind: MenuKind; cell: Addr };

/** Resolve which menu flavour to show based on the click target. Header
 *  clicks promote the selection to the whole row/column so the action
 *  inherits a sensible band. */
export const resolveContextMenuTarget = (
  store: SpreadsheetStore,
  hitHost: HTMLElement,
  e: MouseEvent,
  updateCellSelection: boolean,
  canChangeSelection: () => boolean,
): MenuTarget => {
  const rect = hitHost.getBoundingClientRect();
  const x = e.clientX - rect.left;
  const y = e.clientY - rect.top;
  const s = store.getState();
  const zone = hitZone(layoutForView(s), s.viewport, x, y, null, { resizeHandles: false });
  const fallback = { kind: 'cell' as const, cell: { ...s.selection.active } };
  if (!zone) return fallback;
  const selectedRanges = [s.selection.range, ...(s.selection.extraRanges ?? [])];
  if (zone.kind === 'row-header' || zone.kind === 'row-resize') {
    const inSel = selectedRanges.some(
      (sel) => zone.row >= sel.r0 && zone.row <= sel.r1 && sel.c0 === 0 && sel.c1 >= MAX_COL,
    );
    if (!inSel && canChangeSelection()) mutators.selectRow(store, zone.row);
    return {
      kind: 'row',
      cell: { ...store.getState().selection.active },
    };
  }
  if (zone.kind === 'col-header' || zone.kind === 'col-resize') {
    const inSel = selectedRanges.some(
      (sel) => zone.col >= sel.c0 && zone.col <= sel.c1 && sel.r0 === 0 && sel.r1 >= MAX_ROW,
    );
    if (!inSel && canChangeSelection()) mutators.selectCol(store, zone.col);
    return {
      kind: 'col',
      cell: { ...store.getState().selection.active },
    };
  }
  if (zone.kind === 'cell') {
    const selected = selectedRanges.find(
      (sel) => zone.row >= sel.r0 && zone.row <= sel.r1 && zone.col >= sel.c0 && zone.col <= sel.c1,
    );
    if (selected?.c0 === 0 && selected.c1 >= MAX_COL) {
      return {
        kind: 'row',
        cell: { sheet: s.selection.active.sheet, row: zone.row, col: zone.col },
      };
    }
    if (selected?.r0 === 0 && selected.r1 >= MAX_ROW) {
      return {
        kind: 'col',
        cell: { sheet: s.selection.active.sheet, row: zone.row, col: zone.col },
      };
    }
    if (!selected && updateCellSelection && canChangeSelection()) {
      const cell = { sheet: s.selection.active.sheet, row: zone.row, col: zone.col };
      mutators.setActive(store, cell);
      return { kind: 'cell', cell };
    }
    return {
      kind: 'cell',
      cell: { sheet: s.selection.active.sheet, row: zone.row, col: zone.col },
    };
  }
  return fallback;
};
