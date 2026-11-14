import { staleFormulaMarked } from '../commands/error-indicators.js';
import { hyperlinkCoverAt, isFollowableHyperlinkTarget } from '../commands/hyperlinks.js';
import { addrKey } from '../engine/address.js';
import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { isMacHost } from '../extensions/ui-options.js';
import type { Strings } from '../i18n/strings.js';
import { cellRect, hitTest, layoutForView, pointPx } from '../render/geometry.js';
import { gridHitState, triangleHitContains } from '../render/grid/hit-state.js';
import type { SpreadsheetStore } from '../store/store.js';
import { overlayPortalFor } from './overlay-portal.js';
import { placePanel } from './overlay-position.js';

/** Measured desktop note card size, in points. */
const NOTE_CARD_W_PT = 105;
const NOTE_CARD_H_PT = 77;

export interface HoverDeps {
  /** The grid surface hosting pointer events. */
  grid: HTMLElement;
  store: SpreadsheetStore;
  /** Hover tip on the error triangle of a stale formula value. */
  staleTip?: {
    strings: () => Strings;
    isStale: (addr: Addr) => boolean;
  };
}

/** The hover stale-value tip for a store, reading the workbook lazily. */
export function staleTipFor(
  store: SpreadsheetStore,
  getWb: () => WorkbookHandle | null,
  strings: () => Strings,
): NonNullable<HoverDeps['staleTip']> {
  return {
    strings,
    isStale: (addr) => {
      const state = store.getState();
      const cell = state.data.cells.get(addrKey(addr));
      return staleFormulaMarked(getWb(), state.ui.formatStaleValues, cell?.formula, addr);
    },
  };
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

  // The stale-value tip rides the cell's error triangle.
  const staleTipAt = (e: PointerEvent): string | undefined => {
    const stale = deps.staleTip;
    if (!stale) return undefined;
    const box = grid.getBoundingClientRect();
    const x = e.clientX - box.left;
    const y = e.clientY - box.top;
    const hit = gridHitState(grid).errorTriangles.find(
      (t) => t.kind === 'error' && triangleHitContains(t, x, y),
    );
    if (!hit || !stale.isStale(hit.addr)) return undefined;
    const m = stale.strings().mac.errorMenu;
    return `${m.staleHeading}\n${m.staleBody}`;
  };

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
    const note = fmt?.commentVisible || store.getState().ui.showAllNotes ? undefined : fmt?.comment;
    const tooltip = staleTipAt(e) || note || fmt?.hyperlinkTooltip;
    if (tooltip) {
      if (tooltip !== shownText) {
        tip.textContent = tooltip;
        tip.hidden = false;
        shownText = tooltip;
      }
      // The Mac note card hangs off the cell, so it stays put while the pointer moves inside it.
      const anchored = !!note && !!at && isMacHost(grid);
      if (tip.classList.contains('fc-hover-tip--note') !== anchored) {
        tip.classList.toggle('fc-hover-tip--note', anchored);
      }
      let px = e.clientX + 12;
      let py = e.clientY + 14;
      tip.style.width = '';
      tip.style.minHeight = '';
      tip.style.borderWidth = '';
      if (anchored && at) {
        const s = store.getState();
        const layout = layoutForView(s);
        const pt = pointPx(layout, s.viewport);
        tip.style.width = `${NOTE_CARD_W_PT * pt}px`;
        tip.style.minHeight = `${NOTE_CARD_H_PT * pt}px`;
        tip.style.borderWidth = `${pt}px`;
        const cell = cellRect(layout, s.viewport, at.row, at.col);
        const above = cellRect(layout, s.viewport, Math.max(0, at.row - 1), at.col);
        const gridBox = grid.getBoundingClientRect();
        px = gridBox.left + cell.x + cell.w;
        py = gridBox.top + above.y;
      }
      placePanel(tip, px, py, {
        pad: 8,
        fallbackWidth: 280,
        fallbackHeight: 32,
      });
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
