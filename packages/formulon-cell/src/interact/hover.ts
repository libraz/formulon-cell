import { hyperlinkCoverAt, isFollowableHyperlinkTarget } from '../commands/hyperlinks.js';
import { addrKey } from '../engine/address.js';
import { hitTest, layoutForView } from '../render/geometry.js';
import type { SpreadsheetStore } from '../store/store.js';
import { overlayPortalFor } from './overlay-portal.js';
import { clampPanelToViewport } from './overlay-position.js';

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
 *  - Cursor adjustments — `pointer` cursor over hyperlink cells while modifier
 *    is held (spreadsheet convention).
 */
export function attachHover(deps: HoverDeps): HoverHandle {
  const { grid, store } = deps;
  const ownerDocument = grid.ownerDocument;
  const ownerWindow = ownerDocument.defaultView;

  const tip = ownerDocument.createElement('div');
  tip.className = 'fc-hover-tip';
  tip.setAttribute('role', 'tooltip');
  tip.style.position = 'fixed';
  tip.style.pointerEvents = 'none';
  tip.hidden = true;
  overlayPortalFor(grid).appendChild(tip);

  let modifier = false;

  const cellAt = (e: MouseEvent | PointerEvent): { row: number; col: number } | null => {
    const rect = grid.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;
    const s = store.getState();
    return hitTest(layoutForView(s), s.viewport, x, y);
  };

  const cellFormatAt = (
    row: number,
    col: number,
  ): {
    hyperlink?: string;
    hyperlinkTooltip?: string;
    comment?: string;
    commentVisible?: boolean;
  } | null => {
    const s = store.getState();
    const addr = { sheet: s.data.sheetIndex, row, col };
    const fmt = s.format.formats.get(addrKey(addr));
    // A cell inside a multi-cell link answers with the link's anchor.
    const link = hyperlinkCoverAt(s, addr)?.fmt;
    if (!link) return fmt ?? null;
    return {
      ...fmt,
      hyperlink: link.hyperlink || `#${link.hyperlinkLocation}`,
      hyperlinkTooltip: link.hyperlinkTooltip,
    };
  };

  const onMove = (e: PointerEvent): void => {
    const at = cellAt(e);
    const fmt = at ? cellFormatAt(at.row, at.col) : null;
    // A shown note is already on the canvas; hovering it adds no popup.
    const tooltip = (fmt?.commentVisible ? undefined : fmt?.comment) || fmt?.hyperlinkTooltip;
    if (tooltip) {
      tip.textContent = tooltip;
      tip.hidden = false;
      const position = clampPanelToViewport(tip, e.clientX + 12, e.clientY + 14, {
        pad: 8,
        fallbackWidth: 280,
        fallbackHeight: 32,
      });
      tip.style.left = `${Math.round(position.x)}px`;
      tip.style.top = `${Math.round(position.y)}px`;
    } else {
      tip.hidden = true;
    }
    grid.style.cursor =
      fmt?.hyperlink && modifier && isFollowableHyperlinkTarget(fmt.hyperlink) ? 'pointer' : '';
  };

  const onLeave = (): void => {
    tip.hidden = true;
    grid.style.cursor = '';
  };

  const onViewportChange = (): void => {
    tip.hidden = true;
    grid.style.cursor = '';
  };

  const onModifier = (e: KeyboardEvent): void => {
    modifier = e.ctrlKey || e.metaKey;
    if (!modifier) grid.style.cursor = '';
  };

  grid.addEventListener('pointermove', onMove);
  grid.addEventListener('pointerleave', onLeave);
  ownerWindow?.addEventListener('keydown', onModifier);
  ownerWindow?.addEventListener('keyup', onModifier);
  ownerDocument.addEventListener('scroll', onViewportChange, true);
  ownerWindow?.addEventListener('resize', onViewportChange);

  return {
    detach() {
      grid.removeEventListener('pointermove', onMove);
      grid.removeEventListener('pointerleave', onLeave);
      ownerWindow?.removeEventListener('keydown', onModifier);
      ownerWindow?.removeEventListener('keyup', onModifier);
      ownerDocument.removeEventListener('scroll', onViewportChange, true);
      ownerWindow?.removeEventListener('resize', onViewportChange);
      grid.style.cursor = '';
      tip.remove();
    },
  };
}
