import { writeInput } from '../commands/coerce-input.js';
import { resolveListValues } from '../commands/validate.js';
import { addrKey } from '../engine/address.js';
import { makeRangeResolver } from '../engine/range-resolver.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { getValidationChevron } from '../render/grid.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { inheritHostTokens } from './inherit-host-tokens.js';

export interface ValidationListDeps {
  /** The grid surface that paints the chevron and receives clicks. */
  grid: HTMLElement;
  store: SpreadsheetStore;
  wb: WorkbookHandle;
  /** Refresh cached cells after a write — same contract as InlineEditor. */
  onAfterCommit: () => void;
}

export interface ValidationListHandle {
  detach(): void;
}

/**
 * Owns the click-to-open list dropdown for cells with `validation.kind === 'list'`.
 * Hit-tests against the chevron rect surfaced by the renderer; clicking opens
 * a popover, picking writes the chosen value to the cell.
 */
export function attachValidationList(deps: ValidationListDeps): ValidationListHandle {
  const { grid, store, wb } = deps;
  let popover: HTMLDivElement | null = null;
  let restoreFocus: HTMLElement | null = null;

  const close = (shouldRestoreFocus = false): void => {
    if (!popover) return;
    popover.remove();
    popover = null;
    document.removeEventListener('mousedown', onDocMouseDown, true);
    document.removeEventListener('keydown', onDocKey, true);
    const focusTarget = restoreFocus;
    restoreFocus = null;
    if (shouldRestoreFocus) focusTarget?.focus({ preventScroll: true });
  };

  const onDocMouseDown = (e: MouseEvent): void => {
    if (!popover) return;
    if (popover.contains(e.target as Node)) return;
    close();
  };
  const onDocKey = (e: KeyboardEvent): void => {
    if (e.key === 'Escape') {
      e.preventDefault();
      close(true);
    }
  };

  const items = (): HTMLElement[] =>
    popover ? Array.from(popover.querySelectorAll<HTMLElement>('.fc-validation-list__item')) : [];

  const focusItem = (idx: number): void => {
    const options = items();
    if (options.length === 0) return;
    const next = (idx + options.length) % options.length;
    for (const [i, option] of options.entries()) {
      option.tabIndex = i === next ? 0 : -1;
      option.setAttribute('aria-selected', i === next ? 'true' : 'false');
    }
    options[next]?.focus({ preventScroll: true });
    options[next]?.scrollIntoView({ block: 'nearest' });
  };

  const commitValue = (row: number, col: number, value: string): void => {
    const sheet = store.getState().data.sheetIndex;
    try {
      writeInput(wb, { sheet, row, col }, value);
    } catch (err) {
      console.warn('formulon-cell: validation write failed', err);
    }
    deps.onAfterCommit();
    close();
  };

  const open = (row: number, col: number, list: string[]): void => {
    close();
    if (list.length === 0) return;
    const rect = grid.getBoundingClientRect();
    const chevron = getValidationChevron();
    if (!chevron) return;

    const div = document.createElement('div');
    div.className = 'fc-validation-list';
    div.style.position = 'fixed';
    div.style.left = `${rect.left + chevron.rect.x}px`;
    div.style.top = `${rect.top + chevron.rect.y + chevron.rect.h}px`;
    div.setAttribute('role', 'listbox');
    div.tabIndex = -1;
    restoreFocus =
      document.activeElement instanceof HTMLElement && document.activeElement !== document.body
        ? document.activeElement
        : grid;

    for (const [idx, v] of list.entries()) {
      const item = document.createElement('div');
      item.className = 'fc-validation-list__item';
      item.setAttribute('role', 'option');
      item.setAttribute('aria-selected', idx === 0 ? 'true' : 'false');
      item.tabIndex = idx === 0 ? 0 : -1;
      item.textContent = v;
      item.addEventListener('mousedown', (e) => {
        e.preventDefault();
        commitValue(row, col, v);
      });
      item.addEventListener('click', (e) => {
        e.preventDefault();
        commitValue(row, col, v);
      });
      div.appendChild(item);
    }
    div.addEventListener('keydown', (e) => {
      const options = items();
      const active = document.activeElement instanceof HTMLElement ? document.activeElement : null;
      const idx = active ? options.indexOf(active) : -1;
      if (e.key === 'ArrowDown') {
        e.preventDefault();
        focusItem(idx + 1);
      } else if (e.key === 'ArrowUp') {
        e.preventDefault();
        focusItem(idx - 1);
      } else if (e.key === 'Home') {
        e.preventDefault();
        focusItem(0);
      } else if (e.key === 'End') {
        e.preventDefault();
        focusItem(options.length - 1);
      } else if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        const value = active?.textContent ?? '';
        if (value) commitValue(row, col, value);
      } else if (e.key === 'Escape') {
        e.preventDefault();
        close(true);
      }
    });

    inheritHostTokens(grid, div);
    document.body.appendChild(div);
    popover = div;
    focusItem(0);
    document.addEventListener('mousedown', onDocMouseDown, true);
    document.addEventListener('keydown', onDocKey, true);
  };

  const onDown = (e: PointerEvent): void => {
    if (e.button !== 0) return;
    const chevron = getValidationChevron();
    if (!chevron) return;
    const rect = grid.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;
    if (
      x < chevron.rect.x ||
      x > chevron.rect.x + chevron.rect.w ||
      y < chevron.rect.y ||
      y > chevron.rect.y + chevron.rect.h
    ) {
      return;
    }
    const s = store.getState();
    const fmt = s.format.formats.get(
      addrKey({ sheet: s.data.sheetIndex, row: chevron.row, col: chevron.col }),
    );
    if (fmt?.validation?.kind !== 'list') return;
    e.preventDefault();
    e.stopPropagation();
    // Re-anchor the active cell to the chevron's cell so subsequent picks
    //  hit the same target.
    mutators.setActive(store, { sheet: s.data.sheetIndex, row: chevron.row, col: chevron.col });
    const values = resolveListValues(fmt.validation, makeRangeResolver(wb, s.data.sheetIndex));
    open(chevron.row, chevron.col, values);
  };

  // Capture phase so we beat the regular pointer.ts handler.
  grid.addEventListener('pointerdown', onDown, true);

  return {
    detach() {
      close();
      grid.removeEventListener('pointerdown', onDown, true);
    },
  };
}
