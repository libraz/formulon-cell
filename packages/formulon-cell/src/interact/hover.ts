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

  let shownText: string | null = null;
  const hideTip = (): void => {
    shownText = null;
    tip.hidden = true;
  };

  let modifier = false;
  // The grid's cursor is shared with the pointer layer's resize and fill
  // cursors; only the link cursor this layer set is ever cleared here.
  let linkCursor = false;
  const setLinkCursor = (on: boolean): void => {
    if (on) grid.style.cursor = 'pointer';
    else if (linkCursor) grid.style.cursor = '';
    linkCursor = on;
  };

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

  // Format lookup for the last hovered cell, valid while the format state is unchanged.
  let last: {
    row: number;
    col: number;
    sheet: number;
    format: unknown;
    fmt: ReturnType<typeof cellFormatAt>;
  } | null = null;

  const onMove = (e: PointerEvent): void => {
    const at = cellAt(e);
    let fmt: ReturnType<typeof cellFormatAt> = null;
    if (at) {
      const s = store.getState();
      if (
        last &&
        last.row === at.row &&
        last.col === at.col &&
        last.sheet === s.data.sheetIndex &&
        last.format === s.format
      ) {
        fmt = last.fmt;
      } else {
        fmt = cellFormatAt(at.row, at.col);
        last = { row: at.row, col: at.col, sheet: s.data.sheetIndex, format: s.format, fmt };
      }
    } else {
      last = null;
    }
    // A shown note is already on the canvas; hovering it adds no popup.
    const tooltip = (fmt?.commentVisible ? undefined : fmt?.comment) || fmt?.hyperlinkTooltip;
    if (tooltip) {
      if (tooltip !== shownText) {
        tip.textContent = tooltip;
        tip.hidden = false;
        shownText = tooltip;
      }
      const position = clampPanelToViewport(tip, e.clientX + 12, e.clientY + 14, {
        pad: 8,
        fallbackWidth: 280,
        fallbackHeight: 32,
      });
      tip.style.left = `${Math.round(position.x)}px`;
      tip.style.top = `${Math.round(position.y)}px`;
    } else {
      hideTip();
    }
    setLinkCursor(!!fmt?.hyperlink && modifier && isFollowableHyperlinkTarget(fmt.hyperlink));
  };

  const onLeave = (): void => {
    hideTip();
    setLinkCursor(false);
  };

  const onModifier = (e: KeyboardEvent): void => {
    modifier = e.ctrlKey || e.metaKey;
    if (!modifier) setLinkCursor(false);
  };

  grid.addEventListener('pointermove', onMove);
  grid.addEventListener('pointerleave', onLeave);
  ownerWindow?.addEventListener('keydown', onModifier);
  ownerWindow?.addEventListener('keyup', onModifier);
  ownerDocument.addEventListener('scroll', onLeave, true);
  ownerWindow?.addEventListener('resize', onLeave);

  return {
    detach() {
      grid.removeEventListener('pointermove', onMove);
      grid.removeEventListener('pointerleave', onLeave);
      ownerWindow?.removeEventListener('keydown', onModifier);
      ownerWindow?.removeEventListener('keyup', onModifier);
      ownerDocument.removeEventListener('scroll', onLeave, true);
      ownerWindow?.removeEventListener('resize', onLeave);
      setLinkCursor(false);
      tip.remove();
    },
  };
}
