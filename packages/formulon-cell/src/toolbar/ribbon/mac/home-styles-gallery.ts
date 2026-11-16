import {
  activeCellStyleId,
  applyCellStyleToSelection,
  CELL_STYLES,
  type CellStyleId,
  cellStyleFallbackProfileForPlatform,
  resolveCellStyleFormat,
} from '../../../commands/cell-styles.js';
import { runAfterAllowedRangeUnlock, selectionEditRanges } from '../../../commands/protection.js';
import { dictionaries } from '../../../i18n/strings.js';
import { createInteractionChipButton } from '../../../interact/chip-button.js';
import { focusSheet } from '../../../interact/dialog-focus.js';
import type { SpreadsheetInstance } from '../../../mount/types.js';
import type { ToolbarLang } from '../../ribbon-model.js';
import { paintCellStyleChip } from '../menus/styles.js';

/** Styles the Mac Home tab shows in its gallery strip, row by row. */
const STRIP_STYLE_IDS: readonly CellStyleId[] = [
  'normal',
  'neutral',
  'bad',
  'good',
  'checkCell',
  'note',
  'linkedCell',
  'calculation',
  'warning',
  'outputCell',
];
const STRIP_COLUMNS = 5;

export interface MacCellStylesGalleryOptions {
  lang: ToolbarLang;
  getInst: () => SpreadsheetInstance | null;
  /** The ribbon's own cell-styles button; it opens the full menu. */
  arrow: HTMLButtonElement;
  /** The full cell-styles menu, placed right after the arrow. */
  menu: HTMLElement | null;
}

/**
 * The Home tab's live cell-styles strip: the leading styles as swatches that
 * apply on click, with the cell-styles button as the more arrow.
 */
export const createMacCellStylesGallery = (opts: MacCellStylesGalleryOptions): HTMLDivElement => {
  const { lang, getInst, arrow, menu } = opts;
  const labels = dictionaries[lang].cellStylesGallery.styles;
  const wrap = document.createElement('div');
  wrap.className = 'fc-tb__cellstyle-gallery';
  const grid = document.createElement('div');
  grid.className = 'fc-tb__cellstyle-gallery-grid';
  grid.style.gridTemplateColumns = `repeat(${STRIP_COLUMNS}, 1fr)`;
  grid.setAttribute('role', 'group');

  const profile = (inst: SpreadsheetInstance) =>
    cellStyleFallbackProfileForPlatform(inst.host.dataset.fcPlatform);
  const buildInst = getInst();
  const buildProfile = buildInst ? profile(buildInst) : null;
  const chips = new Map<string, HTMLButtonElement>();
  for (const id of STRIP_STYLE_IDS) {
    const label = labels[id] ?? CELL_STYLES.find((s) => s.id === id)?.label ?? id;
    const chip = createInteractionChipButton({
      className: 'fc-tb__cellstyle-gallery-chip',
      label,
      dataset: { cellStyle: id },
      pressed: false,
    });
    chip.title = label;
    chip.setAttribute('aria-label', label);
    paintCellStyleChip(
      chip,
      buildInst && buildProfile
        ? resolveCellStyleFormat(id, buildInst.workbook ?? null, buildProfile)
        : {},
    );
    chip.addEventListener('click', () => {
      const current = getInst();
      if (!current) return;
      runAfterAllowedRangeUnlock(
        current.store,
        selectionEditRanges(current.store.getState()),
        () => {
          applyCellStyleToSelection(current.store, current.history, id, {
            origin: 'ribbon',
            commandId: 'cellStyles',
            getWorkbook: () => current.workbook,
            getFallbackProfile: () => profile(current),
          });
          focusSheet(current.host);
        },
      );
    });
    chips.set(id, chip);
    grid.appendChild(chip);
  }

  const syncActive = (): void => {
    const current = getInst();
    if (!current) return;
    const active = activeCellStyleId(current.store.getState());
    for (const [id, chip] of chips) chip.setAttribute('aria-pressed', String(id === active));
  };
  syncActive();
  const inst = getInst();
  if (inst) {
    // The ribbon re-renders often; drop the subscription once this strip has left the page.
    // The strip is attached by its caller right after creation, so the first check waits a tick.
    let settled = false;
    const unsubscribe = inst.store.subscribe(() => {
      if (wrap.isConnected) syncActive();
      else if (settled) unsubscribe();
    });
    queueMicrotask(() => {
      settled = true;
      if (!wrap.isConnected) unsubscribe();
    });
  }

  arrow.classList.add('fc-tb__cellstyle-gallery-arrow');
  wrap.append(grid, arrow);
  if (menu) wrap.appendChild(menu);
  return wrap;
};
