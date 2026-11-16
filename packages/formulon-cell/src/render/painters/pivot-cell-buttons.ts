import type { PivotStyleLayer } from '../../engine/pivot-style.js';
import type { Rect } from '../geometry.js';
import { ruleWidth } from '../grid/chrome-context.js';
import { pivotFieldButtonRect, pivotFieldButtonSize } from './filter-button.js';
import { pivotCollapseButtonRect } from './pivot-outline-button.js';

/** What sizes a report cell's buttons: the platform, the zoom, screen px per point, the device
 *  pixel ratio and the sheet direction. */
export interface PivotCellButtonMetrics {
  mac: boolean;
  zoom: number;
  ptPx: number;
  dpr: number;
  rtl: boolean;
}

/** A button a report cell draws: a row label's expand/collapse button or a field caption's
 *  drop-down. */
export type PivotCellButton =
  | { kind: 'outline'; rect: Rect; collapsed: boolean }
  | { kind: 'field'; rect: Rect; filtered: boolean };

/** The buttons of the report cell at (row, col) whose box is `cell`, in paint order. Painting and
 *  hit-testing both read this, so a button is hit exactly where it is drawn. */
export function pivotCellButtons(
  layer: PivotStyleLayer,
  row: number,
  col: number,
  cell: Rect,
  m: PivotCellButtonMetrics,
): PivotCellButton[] {
  const buttons: PivotCellButton[] = [];
  const outline = layer.outlineAt(row, col);
  if (outline?.collapse) {
    buttons.push({
      kind: 'outline',
      rect: pivotCollapseButtonRect(cell, outline.steps, m.ptPx, m.rtl),
      collapsed: outline.collapsed,
    });
  }
  if (layer.hasButton(row, col)) {
    buttons.push({
      kind: 'field',
      rect: pivotFieldButtonRect(
        cell,
        pivotFieldButtonSize(m.mac, m.zoom),
        ruleWidth(m.dpr),
        m.rtl,
      ),
      filtered: layer.buttonAt(row, col)?.filtered === true,
    });
  }
  return buttons;
}

/** The button of the report cell at (row, col) under (x, y), or null. */
export function pivotCellButtonAt(
  layer: PivotStyleLayer,
  row: number,
  col: number,
  cell: Rect,
  m: PivotCellButtonMetrics,
  x: number,
  y: number,
): PivotCellButton | null {
  return (
    pivotCellButtons(layer, row, col, cell, m).find(
      ({ rect }) => x >= rect.x && x < rect.x + rect.w && y >= rect.y && y < rect.y + rect.h,
    ) ?? null
  );
}
