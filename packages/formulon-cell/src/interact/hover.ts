import { addrKey } from '../engine/workbook-handle.js';
import { hitTest } from '../render/geometry.js';
import type { SpreadsheetStore } from '../store/store.js';

export interface HoverDeps {
  /** The grid surface hosting pointer events. */
  grid: HTMLElement;
  store: SpreadsheetStore;
}

export interface HoverHandle {
  detach(): void;
}

/**
 * Owns:
 *  - Comment tooltip overlay — shown when hovering a cell with `format.comment`.
 *  - Hyperlink activation — Ctrl/Cmd-click opens `format.hyperlink` in a new tab.
 *  - Cursor adjustments — `pointer` cursor over hyperlink cells while modifier
 *    is held (Excel/Sheets convention).
 */
export function attachHover(deps: HoverDeps): HoverHandle {
  const { grid, store } = deps;

  const tip = document.createElement('div');
  tip.className = 'fc-hover-tip';
  tip.style.position = 'fixed';
  tip.style.pointerEvents = 'none';
  tip.style.zIndex = '900';
  tip.hidden = true;
  document.body.appendChild(tip);

  let modifier = false;

  const cellAt = (e: MouseEvent | PointerEvent): { row: number; col: number } | null => {
    const rect = grid.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;
    const s = store.getState();
    return hitTest(s.layout, s.viewport, x, y);
  };

  const cellFormatAt = (
    row: number,
    col: number,
  ): { hyperlink?: string; comment?: string } | null => {
    const s = store.getState();
    const fmt = s.format.formats.get(addrKey({ sheet: s.data.sheetIndex, row, col }));
    return fmt ?? null;
  };

  const onMove = (e: PointerEvent): void => {
    const at = cellAt(e);
    const fmt = at ? cellFormatAt(at.row, at.col) : null;
    if (fmt?.comment) {
      tip.textContent = fmt.comment;
      tip.style.left = `${e.clientX + 12}px`;
      tip.style.top = `${e.clientY + 14}px`;
      tip.hidden = false;
    } else {
      tip.hidden = true;
    }
    if (fmt?.hyperlink && modifier) grid.style.cursor = 'pointer';
  };

  const onLeave = (): void => {
    tip.hidden = true;
  };

  const onClick = (e: MouseEvent): void => {
    if (!(e.ctrlKey || e.metaKey)) return;
    const at = cellAt(e);
    if (!at) return;
    const fmt = cellFormatAt(at.row, at.col);
    if (!fmt?.hyperlink) return;
    e.preventDefault();
    e.stopPropagation();
    window.open(fmt.hyperlink, '_blank', 'noopener,noreferrer');
  };

  const onModifier = (e: KeyboardEvent): void => {
    modifier = e.ctrlKey || e.metaKey;
  };

  grid.addEventListener('pointermove', onMove);
  grid.addEventListener('pointerleave', onLeave);
  grid.addEventListener('click', onClick, true);
  window.addEventListener('keydown', onModifier);
  window.addEventListener('keyup', onModifier);

  return {
    detach() {
      grid.removeEventListener('pointermove', onMove);
      grid.removeEventListener('pointerleave', onLeave);
      grid.removeEventListener('click', onClick, true);
      window.removeEventListener('keydown', onModifier);
      window.removeEventListener('keyup', onModifier);
      tip.remove();
    },
  };
}
