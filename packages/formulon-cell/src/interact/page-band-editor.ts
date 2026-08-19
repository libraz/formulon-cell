// Inline editing of the header / footer slots drawn in Page Layout view.
//
// The bands sit in the page margins, which hold no cells, so a click there is
// unambiguous: the pointer layer emits `fc:editpageband` and this opens a text
// field over the slot. The value is written straight into the sheet's page
// setup, which is the same field the Page Setup dialog edits — the two views
// are just different ways at the same six strings.

import type { History } from '../commands/history.js';
import { setPageSetup } from '../commands/page-setup.js';
import type { PageBandHit } from '../render/grid/page-view.js';
import { getPageSetup, type PageSetup, type SpreadsheetStore } from '../store/store.js';

export interface PageBandEditorDeps {
  /** Element the bands are painted in — the editor is positioned inside it. */
  grid: HTMLElement;
  store: SpreadsheetStore;
  history?: History | null;
}

export interface PageBandEditorHandle {
  /** True while a slot is open for editing. */
  isOpen(): boolean;
  detach(): void;
}

const FIELD_BY_SLOT: Record<PageBandHit['kind'], Record<PageBandHit['slot'], keyof PageSetup>> = {
  header: { left: 'headerLeft', center: 'headerCenter', right: 'headerRight' },
  footer: { left: 'footerLeft', center: 'footerCenter', right: 'footerRight' },
};

export function attachPageBandEditor(deps: PageBandEditorDeps): PageBandEditorHandle {
  const { grid, store } = deps;
  let input: HTMLInputElement | null = null;

  const close = (commit: boolean): void => {
    const field = input;
    if (!field) return;
    input = null;
    const key = field.dataset.fcField as keyof PageSetup | undefined;
    const sheet = Number(field.dataset.fcSheet ?? '0');
    field.remove();
    if (!commit || !key) return;
    const next = field.value;
    const current = getPageSetup(store.getState(), sheet)[key];
    if ((current ?? '') === next) return;
    // An emptied slot drops the field rather than storing "", so the print
    // document does not emit a blank margin box for it.
    setPageSetup(store, sheet, { [key]: next || undefined }, deps.history ?? null);
  };

  const onEditBand = (event: Event): void => {
    const detail = (event as CustomEvent<PageBandHit>).detail;
    if (!detail) return;
    close(true);
    const sheet = store.getState().data.sheetIndex;
    const key = FIELD_BY_SLOT[detail.kind][detail.slot];
    const field = document.createElement('input');
    field.type = 'text';
    field.className = 'fc-pageband-editor';
    field.dataset.fcField = key;
    field.dataset.fcSheet = String(sheet);
    field.dataset.fcSlot = detail.slot;
    field.dataset.fcBand = detail.kind;
    field.value = String(getPageSetup(store.getState(), sheet)[key] ?? '');
    field.style.left = `${detail.rect.x}px`;
    field.style.top = `${detail.rect.y}px`;
    field.style.width = `${detail.rect.w}px`;
    field.style.height = `${detail.rect.h}px`;
    field.style.textAlign = detail.slot;
    field.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        close(true);
        grid.focus();
      } else if (e.key === 'Escape') {
        e.preventDefault();
        close(false);
        grid.focus();
      }
      e.stopPropagation();
    });
    field.addEventListener('blur', () => close(true));
    input = field;
    grid.appendChild(field);
    field.focus();
    field.select();
  };

  grid.addEventListener('fc:editpageband', onEditBand);

  return {
    isOpen: () => input !== null,
    detach() {
      close(false);
      grid.removeEventListener('fc:editpageband', onEditBand);
    },
  };
}
