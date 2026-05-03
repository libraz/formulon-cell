import { writeInput } from '../commands/coerce-input.js';
import { addrKey } from '../engine/workbook-handle.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { getValidationChevron } from '../render/grid.js';
import { type SpreadsheetStore, mutators } from '../store/store.js';

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

  const close = (): void => {
    if (!popover) return;
    popover.remove();
    popover = null;
    document.removeEventListener('mousedown', onDocMouseDown, true);
    document.removeEventListener('keydown', onDocKey, true);
  };

  const onDocMouseDown = (e: MouseEvent): void => {
    if (!popover) return;
    if (popover.contains(e.target as Node)) return;
    close();
  };
  const onDocKey = (e: KeyboardEvent): void => {
    if (e.key === 'Escape') {
      e.preventDefault();
      close();
    }
  };

  const open = (row: number, col: number, list: string[]): void => {
    close();
    const rect = grid.getBoundingClientRect();
    const chevron = getValidationChevron();
    if (!chevron) return;

    const div = document.createElement('div');
    div.className = 'fc-validation-list';
    div.style.position = 'fixed';
    div.style.left = `${rect.left + chevron.rect.x}px`;
    div.style.top = `${rect.top + chevron.rect.y + chevron.rect.h}px`;
    div.setAttribute('role', 'listbox');

    for (const v of list) {
      const item = document.createElement('div');
      item.className = 'fc-validation-list__item';
      item.setAttribute('role', 'option');
      item.textContent = v;
      item.addEventListener('mousedown', (e) => {
        e.preventDefault();
        const sheet = store.getState().data.sheetIndex;
        try {
          writeInput(wb, { sheet, row, col }, v);
        } catch (err) {
          console.warn('formulon-cell: validation write failed', err);
        }
        deps.onAfterCommit();
        close();
      });
      div.appendChild(item);
    }

    document.body.appendChild(div);
    popover = div;
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
    open(chevron.row, chevron.col, fmt.validation.source);
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
