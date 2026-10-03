// In-grid mirror of the formula being typed in the formula bar. It paints the
// raw draft over the anchor cell, hiding itself whenever the cell is off-sheet,
// scrolled behind a frozen band, or outside the grid.

import type { Addr } from '../engine/types.js';
import { bodyBandOrigin, cellRectUnclamped, layoutForView } from '../render/geometry.js';
import { formatWithPending } from '../store/pending-format.js';
import type { SpreadsheetStore } from '../store/store.js';

export interface FormulaDraftMirrorDeps {
  grid: HTMLElement;
  store: SpreadsheetStore;
}

export interface FormulaDraftMirror {
  /** Show `raw` over `anchor`; `null` clears the draft. */
  project: (anchor: Addr, raw: string | null) => void;
  /** Re-place the mirror after a scroll, resize or store change. */
  refresh: () => void;
  detach: () => void;
}

export function attachFormulaDraftMirror(deps: FormulaDraftMirrorDeps): FormulaDraftMirror {
  const { grid, store } = deps;
  const el = document.createElement('div');
  el.className = 'fc-host__formula-draft-mirror';
  el.setAttribute('aria-hidden', 'true');
  el.hidden = true;
  grid.appendChild(el);
  let draft: { anchor: Addr; raw: string } | null = null;
  const hide = (): void => {
    el.hidden = true;
  };
  const refresh = (): void => {
    const mirrorState = draft;
    if (!mirrorState) {
      hide();
      return;
    }
    const state = store.getState();
    if (mirrorState.anchor.sheet !== state.data.sheetIndex) {
      hide();
      return;
    }
    const layout = layoutForView(state);
    const rect = cellRectUnclamped(
      layout,
      state.viewport,
      mirrorState.anchor.row,
      mirrorState.anchor.col,
    );
    const band = bodyBandOrigin(layout, state.viewport);
    const behindFrozenBand =
      (mirrorState.anchor.row >= state.layout.freezeRows && rect.y < band.y) ||
      (mirrorState.anchor.col >= state.layout.freezeCols &&
        (layout.rtl ? rect.x + rect.w > band.x : rect.x < band.x));
    const gridRect = grid.getBoundingClientRect();
    const width = grid.clientWidth || gridRect.width;
    const height = grid.clientHeight || gridRect.height;
    const outsideGrid =
      (width > 0 && (rect.x + rect.w <= 0 || rect.x >= width)) ||
      (height > 0 && (rect.y + rect.h <= 0 || rect.y >= height));
    if (behindFrozenBand || outsideGrid || rect.w <= 0 || rect.h <= 0) {
      hide();
      return;
    }
    const format = formatWithPending(state, mirrorState.anchor);
    el.textContent = mirrorState.raw;
    el.style.left = `${rect.x}px`;
    el.style.top = `${rect.y}px`;
    el.style.width = `${rect.w}px`;
    el.style.height = `${rect.h}px`;
    el.style.background = format?.fill ?? '';
    el.style.color = format?.color ?? '';
    el.style.fontFamily = format?.fontFamily ?? '';
    el.style.fontSize = format?.fontSize ? `${format.fontSize}px` : '';
    el.style.fontWeight = format?.bold ? 'bold' : '';
    el.style.fontStyle = format?.italic ? 'italic' : '';
    el.style.textDecoration = format?.underline ? 'underline' : '';
    el.style.textAlign = format?.align ?? '';
    el.style.direction = layout.rtl ? 'rtl' : 'ltr';
    el.hidden = false;
  };
  const project = (anchor: Addr, raw: string | null): void => {
    draft = raw === null ? null : { anchor: { ...anchor }, raw };
    refresh();
  };
  const unsubscribe = store.subscribe(() => refresh());
  return {
    project,
    refresh,
    detach: () => {
      unsubscribe();
      el.remove();
    },
  };
}
