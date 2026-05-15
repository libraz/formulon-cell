import { applyCellStyle, CELL_STYLES, type CellStyleId } from '../commands/cell-styles.js';
import type { History } from '../commands/history.js';
import { flushFormatToEngine } from '../engine/cell-format-sync.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SpreadsheetStore } from '../store/store.js';
import { createDialogShell } from './dialog-shell.js';

export interface CellStylesGalleryDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  /** Workbook accessor — called at apply-time so the gallery survives a
   *  workbook swap. */
  getWb?: () => WorkbookHandle | null;
  history?: History | null;
  /** Optional label override per style id. Chrome that wants a localized
   *  gallery passes `(id) => translatedLabel(id)`; otherwise the default
   *  English label from `CELL_STYLES` is used. */
  labelFor?: (id: CellStyleId) => string;
}

export interface CellStylesGalleryHandle {
  /** Open the gallery as a centered modal, anchored on a recent click. */
  open(): void;
  close(): void;
  detach(): void;
}

/**
 * Compact named-style picker — a grid of preset chips that apply on click.
 * Hangs off `host` so dismissal naturally co-exists with the cell editor and
 * grid surface. The Apply path goes through `applyCellStyle` so each click is
 * one undoable history entry.
 */
export function attachCellStylesGallery(deps: CellStylesGalleryDeps): CellStylesGalleryHandle {
  const { host, store } = deps;
  const history = deps.history ?? null;
  const getWb = deps.getWb ?? ((): WorkbookHandle | null => null);
  const labelFor = deps.labelFor ?? ((id) => CELL_STYLES.find((s) => s.id === id)?.label ?? id);

  const shell = createDialogShell({
    host,
    className: 'fc-stylegallery',
    ariaLabel: 'Cell styles',
    onDismiss: () => close(),
  });
  const { overlay, panel } = shell;

  const grid = document.createElement('div');
  grid.className = 'fc-stylegallery__grid';
  panel.appendChild(grid);

  for (const style of CELL_STYLES) {
    const chip = document.createElement('button');
    chip.type = 'button';
    chip.className = 'fc-stylegallery__chip';
    chip.dataset.fcStyle = style.id;
    if (style.format.bold) chip.style.fontWeight = '700';
    if (style.format.italic) chip.style.fontStyle = 'italic';
    if (style.format.underline) chip.style.textDecoration = 'underline';
    if (style.format.color) chip.style.color = style.format.color;
    if (style.format.fill) chip.style.background = style.format.fill;
    if (style.format.fontSize) chip.style.fontSize = `${style.format.fontSize}px`;
    chip.textContent = labelFor(style.id);
    grid.appendChild(chip);
  }

  const close = (): void => {
    shell.close();
  };

  const apply = (id: CellStyleId): void => {
    const range = store.getState().selection.range;
    applyCellStyle(store, history, range, id);
    const wb = getWb();
    if (wb) flushFormatToEngine(wb, store, range.sheet);
    close();
  };

  const onClick = (e: MouseEvent): void => {
    const target = e.target as HTMLElement;
    if (target === overlay) {
      close();
      return;
    }
    const chip = target.closest('.fc-stylegallery__chip') as HTMLElement | null;
    if (!chip) return;
    const id = chip.dataset.fcStyle as CellStyleId | undefined;
    if (!id) return;
    apply(id);
  };

  const onKey = (e: KeyboardEvent): void => {
    if (e.key === 'Escape') {
      e.preventDefault();
      close();
    }
  };

  shell.on(overlay, 'click', onClick as EventListener);
  shell.on(overlay, 'keydown', onKey as EventListener);

  return {
    open(): void {
      shell.open();
      requestAnimationFrame(() => {
        const first = grid.querySelector<HTMLElement>('.fc-stylegallery__chip');
        first?.focus();
      });
    },
    close,
    detach(): void {
      shell.dispose();
    },
  };
}
