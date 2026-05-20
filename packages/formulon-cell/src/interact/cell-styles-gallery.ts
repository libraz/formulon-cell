import {
  applyCellStyleByName,
  CELL_STYLE_GROUPS,
  CELL_STYLES,
  type CellStyleGroupId,
  type CellStyleId,
  listCustomCellStyles,
} from '../commands/cell-styles.js';
import type { History } from '../commands/history.js';
import { flushFormatToEngine } from '../engine/cell-format-sync.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
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
  strings?: Strings;
}

export interface CellStylesGalleryHandle {
  /** Open the gallery as a centered modal, anchored on a recent click. */
  open(): void;
  close(): void;
  setStrings(strings: Strings): void;
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
  let strings = deps.strings ?? defaultStrings;
  const labelFor =
    deps.labelFor ??
    ((id: CellStyleId): string =>
      strings.cellStylesGallery.styles[id] ?? CELL_STYLES.find((s) => s.id === id)?.label ?? id);

  const shell = createDialogShell({
    host,
    className: 'fc-stylegallery',
    ariaLabel: strings.cellStylesGallery.title,
    onDismiss: () => close(),
  });
  const { overlay, panel } = shell;

  const grid = document.createElement('div');
  grid.className = 'fc-stylegallery__body';
  panel.appendChild(grid);

  let activeChipIndex = 0;
  const chips: HTMLButtonElement[] = [];
  const chipById = new Map<string, HTMLButtonElement>();
  const headings = new Map<CellStyleGroupId, HTMLElement>();
  const styleById = new Map(CELL_STYLES.map((style) => [style.id, style]));
  const focusChip = (idx: number): void => {
    if (chips.length === 0) return;
    activeChipIndex = (idx + chips.length) % chips.length;
    for (const [chipIndex, chip] of chips.entries()) {
      chip.tabIndex = chipIndex === activeChipIndex ? 0 : -1;
    }
    chips[activeChipIndex]?.focus({ preventScroll: true });
  };
  const gridColumnCount = (): number => {
    const activeChip = chips[activeChipIndex];
    const activeGrid = activeChip?.closest('.fc-stylegallery__grid') ?? grid;
    const columns = getComputedStyle(activeGrid).gridTemplateColumns;
    const count = columns ? columns.split(' ').filter(Boolean).length : 0;
    return Math.max(1, count || 3);
  };

  const createChipFromDef = (style: {
    id: string;
    label: string;
    format: {
      bold?: boolean;
      italic?: boolean;
      underline?: boolean;
      color?: string;
      fill?: string;
      fontSize?: number;
    };
  }): HTMLButtonElement => {
    const chip = document.createElement('button');
    chip.type = 'button';
    chip.className = 'fc-stylegallery__chip';
    chip.dataset.fcStyle = style.id;
    chip.tabIndex = chips.length === 0 ? 0 : -1;
    if (style.format.bold) chip.style.fontWeight = '700';
    if (style.format.italic) chip.style.fontStyle = 'italic';
    if (style.format.underline) chip.style.textDecoration = 'underline';
    if (style.format.color) chip.style.color = style.format.color;
    if (style.format.fill) chip.style.background = style.format.fill;
    if (style.format.fontSize) chip.style.fontSize = `${style.format.fontSize}px`;
    chip.textContent = style.label;
    chips.push(chip);
    chipById.set(style.id, chip);
    return chip;
  };

  const createChip = (id: CellStyleId): HTMLButtonElement | null => {
    const style = styleById.get(id);
    if (!style) return null;
    return createChipFromDef({ id: style.id, label: labelFor(style.id), format: style.format });
  };

  const renderGroups = (): void => {
    grid.textContent = '';
    chips.length = 0;
    chipById.clear();
    headings.clear();
    for (const group of CELL_STYLE_GROUPS) {
      const section = document.createElement('section');
      section.className = 'fc-stylegallery__section';

      const heading = document.createElement('div');
      heading.className = 'fc-stylegallery__heading';
      heading.textContent = strings.cellStylesGallery.groups[group.id];
      headings.set(group.id, heading);
      section.appendChild(heading);

      const groupGrid = document.createElement('div');
      groupGrid.className = 'fc-stylegallery__grid';
      groupGrid.setAttribute('role', 'toolbar');
      groupGrid.setAttribute('aria-label', strings.cellStylesGallery.groups[group.id]);
      for (const id of group.styleIds) {
        const chip = createChip(id);
        if (chip) groupGrid.appendChild(chip);
      }
      section.appendChild(groupGrid);
      grid.appendChild(section);
    }
    const customStyles = listCustomCellStyles(store.getState());
    if (customStyles.length > 0) {
      const section = document.createElement('section');
      section.className = 'fc-stylegallery__section';
      const heading = document.createElement('div');
      heading.className = 'fc-stylegallery__heading';
      heading.textContent = 'Custom';
      section.appendChild(heading);
      const groupGrid = document.createElement('div');
      groupGrid.className = 'fc-stylegallery__grid';
      groupGrid.setAttribute('role', 'toolbar');
      groupGrid.setAttribute('aria-label', heading.textContent ?? 'Custom');
      for (const style of customStyles) groupGrid.appendChild(createChipFromDef(style));
      section.appendChild(groupGrid);
      grid.appendChild(section);
    }
  };

  renderGroups();

  const close = (): void => {
    shell.close();
  };

  const apply = (id: string): void => {
    const range = store.getState().selection.range;
    applyCellStyleByName(store, history, range, id);
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
    const id = chip.dataset.fcStyle;
    if (!id) return;
    apply(id);
  };

  const onKey = (e: KeyboardEvent): void => {
    if (e.key === 'Escape') {
      e.preventDefault();
      close();
    } else if (e.key === 'ArrowRight') {
      e.preventDefault();
      focusChip(activeChipIndex + 1);
    } else if (e.key === 'ArrowLeft') {
      e.preventDefault();
      focusChip(activeChipIndex - 1);
    } else if (e.key === 'ArrowDown') {
      e.preventDefault();
      focusChip(activeChipIndex + gridColumnCount());
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      focusChip(activeChipIndex - gridColumnCount());
    } else if (e.key === 'Home') {
      e.preventDefault();
      focusChip(0);
    } else if (e.key === 'End') {
      e.preventDefault();
      focusChip(chips.length - 1);
    }
  };

  shell.on(overlay, 'click', onClick as EventListener);
  shell.on(overlay, 'keydown', onKey as EventListener);

  return {
    open(): void {
      renderGroups();
      shell.open();
      requestAnimationFrame(() => {
        focusChip(activeChipIndex);
      });
    },
    close,
    setStrings(next: Strings): void {
      strings = next;
      shell.setAriaLabel(strings.cellStylesGallery.title);
      for (const [id, heading] of headings) {
        heading.textContent = strings.cellStylesGallery.groups[id];
        const groupGrid = heading.nextElementSibling;
        if (groupGrid) {
          groupGrid.setAttribute('aria-label', strings.cellStylesGallery.groups[id]);
        }
      }
      for (const style of CELL_STYLES) {
        const chip = chipById.get(style.id);
        if (chip) chip.textContent = labelFor(style.id);
      }
    },
    detach(): void {
      shell.dispose();
    },
  };
}
