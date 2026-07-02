import { copy } from '../commands/clipboard/copy.js';
import { cut } from '../commands/clipboard/cut.js';
import { insertCopiedCellsFromTSV } from '../commands/clipboard/insert-copied-cells.js';
import { pasteTSV } from '../commands/clipboard/paste.js';
import { type PasteWhat, pasteSpecial } from '../commands/clipboard/paste-special.js';
import type { ClipboardSnapshot } from '../commands/clipboard/snapshot.js';
import { clearComment } from '../commands/comment.js';
import {
  applyValueFilter,
  clearFilter,
  distinctValues,
  filterValueKey,
  inferAutoFilterRange,
  reapplyFilters,
  recordFilterChange,
} from '../commands/filter.js';
import {
  cycleBorders,
  setAlign,
  toggleBold,
  toggleItalic,
  toggleUnderline,
} from '../commands/format.js';
import { type History, recordFormatChange } from '../commands/history.js';
import { groupCols, groupRows, ungroupCols, ungroupRows } from '../commands/outline.js';
import { inferSortHasHeader, sortRange } from '../commands/sort.js';
import {
  deleteCols,
  deleteRows,
  hiddenInSelection,
  hideCols,
  hideRows,
  insertCols,
  insertRows,
  showCols,
  showRows,
} from '../commands/structure.js';
import { addrKey } from '../engine/address.js';
import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import { hitZone } from '../render/geometry.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import { createInteractionButton } from './chip-button.js';
import {
  buildCellEntries,
  buildColEntries,
  buildRowEntries,
  compactMenuEntries,
  type ItemId,
  type MenuEntry,
  type MenuKind,
  PASTE_QUICK_IDS,
  SUBMENU_ICON_ACTION,
} from './context-menu-spec.js';
import { inheritHostTokens } from './inherit-host-tokens.js';
import { openInsertCopiedCellsDialog } from './insert-copied-cells-dialog.js';
import { clampPanelToViewport, panelSize, viewportSize } from './overlay-position.js';

export interface ContextMenuDeps {
  host: HTMLElement;
  /** Element whose coordinate space matches grid hit-testing. Defaults to host
   *  for standalone tests/legacy embedders. */
  grid?: HTMLElement;
  store: SpreadsheetStore;
  wb: WorkbookHandle;
  /** UI string dictionary. Falls back to the package default (ja) if omitted. */
  strings?: Strings;
  /** Shared history. When provided, format-mutating menu actions push entries
   *  so Cmd+Z reverts them. */
  history?: History | null;
  /** Called after cut/paste/clear so caller can refresh cached cells from engine. */
  onAfterCommit?: () => void;
  /** Called when the user clicks the "Format Cells…" menu entry. */
  onFormatDialog?: () => void;
  /** Called when the user clicks "Paste Special…". */
  onPasteSpecial?: () => void;
  /** Called when the user clicks "Define Name…". When omitted the entry is
   *  hidden. */
  onDefineName?: () => void;
  /** Returns the structured clipboard snapshot used by the Paste Special
   *  submenu's quick-paste actions. When it returns null those entries are
   *  disabled and only the "Paste Special…" dialog entry stays usable. */
  getClipboardSnapshot?: () => ClipboardSnapshot | null;
  /** Optional shared clipboard command path. When provided, context-menu
   *  copy/cut/paste uses it so structured snapshots and Paste Options stay
   *  consistent with keyboard shortcuts. */
  onClipboardShortcut?: (kind: 'copy' | 'cut' | 'paste') => void;
  /** Called when the user clicks "Edit comment…". When omitted the menu
   *  entry is hidden — the action requires the comment dialog feature to
   *  be wired up. */
  onEditComment?: (addr: Addr) => void;
  /** Called when the user clicks "Insert hyperlink…". When omitted the menu
   *  entry is hidden. */
  onInsertHyperlink?: () => void;
  /** Called when the user clicks the Add/Remove Watch entry. When omitted the
   *  menu entry is hidden. */
  onToggleWatch?: (addr: Addr) => void;
  /** Returns true when the active cell is currently watched. */
  isWatched?: (addr: Addr) => boolean;
}

const VIEWPORT_PAD = 4;

const appendContextMenuLabel = (
  button: HTMLButtonElement,
  labelText: string,
  hintText?: string,
): void => {
  const label = document.createElement('span');
  label.className = 'fc-ctxmenu__label';
  label.textContent = labelText;
  button.appendChild(label);
  if (hintText !== undefined) {
    const hint = document.createElement('span');
    hint.className = 'fc-ctxmenu__hint';
    hint.textContent = hintText;
    button.appendChild(hint);
  }
};

const createContextMenuItemButton = (
  entry: Extract<MenuEntry, { kind: 'item' }>,
): HTMLButtonElement => {
  const button = createInteractionButton({
    className: 'fc-ctxmenu__item',
    dataset: { fcAction: entry.id },
    role: 'menuitem',
    tabIndex: -1,
  });
  appendContextMenuLabel(button, entry.label, entry.hint ?? '');
  return button;
};

const createContextSubmenuButton = (
  entry: Extract<MenuEntry, { kind: 'submenu' }>,
): HTMLButtonElement => {
  const button = createInteractionButton({
    className: 'fc-ctxmenu__item fc-ctxmenu__item--parent',
    dataset: {
      fcSubmenu: entry.id,
      fcAction: SUBMENU_ICON_ACTION[entry.id] ?? entry.id,
    },
    role: 'menuitem',
    tabIndex: -1,
  });
  button.setAttribute('aria-haspopup', 'menu');
  button.setAttribute('aria-expanded', 'false');
  appendContextMenuLabel(button, entry.label);
  const arrow = document.createElement('span');
  arrow.className = 'fc-ctxmenu__arrow';
  arrow.textContent = '›';
  arrow.setAttribute('aria-hidden', 'true');
  button.appendChild(arrow);
  return button;
};

const createContextMiniToolbarButton = (item: {
  id: ItemId;
  label: string;
  text?: string;
}): HTMLButtonElement => {
  return createInteractionButton({
    className: 'fc-ctxmenu__mini-btn',
    dataset: { fcAction: item.id },
    ariaLabel: item.label,
    tabIndex: -1,
    text: item.text,
  });
};

/** Detacher returned by `attachContextMenu`. Also exposes `setStrings` so the
 *  active dictionary can be swapped after attach. */
export interface ContextMenuHandle {
  (): void;
  /** Swap the active dictionary; takes effect on next open. */
  setStrings(next: Strings): void;
}

export function attachContextMenu(deps: ContextMenuDeps): ContextMenuHandle {
  const { host, store, wb } = deps;
  const hitHost = deps.grid ?? host;
  const history = deps.history ?? null;
  if (history) wb.attachHistory(history);
  let strings = deps.strings ?? defaultStrings;
  const wrapFmt = (fn: () => void): void => recordFormatChange(history, store, fn);

  const root = document.createElement('div');
  root.className = 'fc-ctxmenu';
  root.setAttribute('role', 'menu');
  root.setAttribute('aria-label', strings.contextMenu.title);
  root.style.display = 'none';
  root.tabIndex = -1;
  document.body.appendChild(root);

  // Single reusable child panel — the context menu is one level deep.
  const sub = document.createElement('div');
  sub.className = 'fc-ctxmenu fc-ctxmenu__sub';
  sub.setAttribute('role', 'menu');
  sub.style.display = 'none';
  sub.tabIndex = -1;
  document.body.appendChild(sub);

  let visible = false;
  let pasteBtnRef: HTMLButtonElement | null = null;
  let activeIndex = -1;
  let focusPanel: 'root' | 'sub' = 'root';
  let restoreFocusEl: HTMLElement | null = null;
  const submenuChildren = new Map<string, MenuEntry[]>();
  let openSub: { id: string; parentBtn: HTMLButtonElement } | null = null;
  let subCloseTimer: ReturnType<typeof setTimeout> | null = null;

  const cancelSubClose = (): void => {
    if (subCloseTimer != null) {
      clearTimeout(subCloseTimer);
      subCloseTimer = null;
    }
  };

  const closeSubmenu = (): void => {
    cancelSubClose();
    if (!openSub) return;
    openSub.parentBtn.setAttribute('aria-expanded', 'false');
    openSub.parentBtn.classList.remove('fc-ctxmenu__item--open');
    openSub = null;
    sub.style.display = 'none';
    sub.replaceChildren();
    if (focusPanel === 'sub') focusPanel = 'root';
  };

  /** Close the submenu after a short grace period — lets the pointer travel
   *  diagonally from a parent row into the child panel without it snapping
   *  shut as it crosses sibling rows. */
  const scheduleSubClose = (): void => {
    if (!openSub) return;
    cancelSubClose();
    subCloseTimer = setTimeout(() => {
      closeSubmenu();
    }, 260);
  };

  const hide = (restoreFocus = false): void => {
    if (!visible) return;
    visible = false;
    closeSubmenu();
    root.style.display = 'none';
    activeIndex = -1;
    focusPanel = 'root';
    const focusTarget = restoreFocusEl;
    restoreFocusEl = null;
    if (restoreFocus) {
      (focusTarget ?? host).focus({ preventScroll: true });
    }
  };

  const activePanel = (): HTMLElement => (focusPanel === 'sub' ? sub : root);

  const panelItems = (panel: HTMLElement): HTMLButtonElement[] =>
    Array.from(panel.querySelectorAll<HTMLButtonElement>('.fc-ctxmenu__item')).filter(
      (btn) => !btn.disabled && btn.getAttribute('aria-disabled') !== 'true',
    );

  const focusMenuItem = (idx: number): void => {
    const items = panelItems(activePanel());
    if (items.length === 0) return;
    activeIndex = (idx + items.length) % items.length;
    items[activeIndex]?.focus();
  };

  const openSubmenu = (id: string, parentBtn: HTMLButtonElement): void => {
    cancelSubClose();
    if (openSub?.id === id) return;
    closeSubmenu();
    const children = submenuChildren.get(id);
    if (!children) return;
    const disabled = new Set<ItemId>();
    if (id === 'pasteSpecialMenu' && !deps.getClipboardSnapshot?.()) {
      for (const d of PASTE_QUICK_IDS) disabled.add(d);
    }
    sub.replaceChildren();
    for (const child of children) appendEntry('sub', sub, child, disabled);
    inheritHostTokens(host, sub);
    sub.style.display = 'block';
    sub.style.left = '-9999px';
    sub.style.top = '-9999px';
    const r = parentBtn.getBoundingClientRect();
    const { width: sw } = panelSize(sub);
    const viewport = viewportSize();
    let x = r.right - 2;
    if (x + sw > viewport.width - VIEWPORT_PAD) x = r.left - sw + 2;
    const pos = clampPanelToViewport(sub, x, r.top - 4, { pad: VIEWPORT_PAD });
    sub.style.left = `${pos.x}px`;
    sub.style.top = `${pos.y}px`;
    openSub = { id, parentBtn };
    parentBtn.setAttribute('aria-expanded', 'true');
    parentBtn.classList.add('fc-ctxmenu__item--open');
  };

  const appendEntry = (
    panel: 'root' | 'sub',
    container: HTMLElement,
    entry: MenuEntry,
    disabledIds: Set<ItemId>,
  ): void => {
    if (entry.kind === 'sep') {
      const sep = document.createElement('hr');
      sep.className = 'fc-ctxmenu__sep';
      container.appendChild(sep);
      return;
    }
    if (entry.kind === 'submenu') {
      submenuChildren.set(entry.id, entry.children);
      const btn = createContextSubmenuButton(entry);
      btn.addEventListener('mouseenter', () => {
        openSubmenu(entry.id, btn);
      });
      btn.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        openSubmenu(entry.id, btn);
        focusPanel = 'sub';
        focusMenuItem(0);
      });
      container.appendChild(btn);
      return;
    }
    const btn = createContextMenuItemButton(entry);
    if (disabledIds.has(entry.id)) {
      setContextMenuItemDisabled(btn, true, strings.contextMenu.pasteSpecialRequiresCopiedCells);
    }
    btn.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      if (btn.disabled) return;
      run(entry.id);
      hide(false);
    });
    btn.addEventListener('mouseenter', () => {
      if (panel === 'root') {
        if (openSub && openSub.parentBtn !== btn) scheduleSubClose();
      } else {
        cancelSubClose();
      }
    });
    container.appendChild(btn);
    if (panel === 'root' && entry.id === 'paste') pasteBtnRef = btn;
  };

  const buildMiniToolbar = (): HTMLElement => {
    const toolbar = document.createElement('div');
    toolbar.className = 'fc-ctxmenu__mini';
    toolbar.setAttribute('role', 'toolbar');
    toolbar.setAttribute('aria-label', strings.contextMenu.title);

    const buttons: readonly { id: ItemId; label: string; text?: string }[] = [
      { id: 'bold', label: strings.contextMenu.bold, text: 'B' },
      { id: 'italic', label: strings.contextMenu.italic, text: 'I' },
      { id: 'underline', label: strings.contextMenu.underline, text: 'U' },
      { id: 'alignLeft', label: strings.contextMenu.alignLeft },
      { id: 'alignCenter', label: strings.contextMenu.alignCenter },
      { id: 'alignRight', label: strings.contextMenu.alignRight },
      { id: 'borders', label: strings.contextMenu.borders },
      { id: 'formatCells', label: strings.contextMenu.formatCells },
    ];

    for (const item of buttons) {
      const btn = createContextMiniToolbarButton(item);
      btn.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        run(item.id);
        hide(false);
      });
      toolbar.appendChild(btn);
    }

    return toolbar;
  };

  const buildMenu = (kind: MenuKind): void => {
    root.replaceChildren();
    submenuChildren.clear();
    pasteBtnRef = null;
    if (kind === 'cell') root.appendChild(buildMiniToolbar());
    const raw =
      kind === 'row'
        ? buildRowEntries(strings)
        : kind === 'col'
          ? buildColEntries(strings)
          : buildCellEntries(strings);
    const activeAddr = store.getState().selection.active;
    const hasCopiedCells = !!store.getState().ui.copyRange;
    const watched = !!deps.isWatched?.(activeAddr);
    const entries = compactMenuEntries(
      raw
        .filter(
          (e) => !(e.kind === 'item' && e.id === 'insertHyperlink' && !deps.onInsertHyperlink),
        )
        .filter((e) => !(e.kind === 'item' && e.id === 'insertCopiedCells' && !hasCopiedCells))
        .filter((e) => !(e.kind === 'item' && e.id === 'insertComment' && !deps.onEditComment))
        .filter((e) => !(e.kind === 'item' && e.id === 'toggleWatch' && !deps.onToggleWatch))
        .filter((e) => !(e.kind === 'item' && e.id === 'defineName' && !deps.onDefineName))
        .map((e) => {
          if (e.kind === 'item' && e.id === 'toggleWatch') {
            return {
              ...e,
              label: watched ? strings.contextMenu.removeWatch : strings.contextMenu.addWatch,
            };
          }
          return e;
        }),
    );
    const noDisabled = new Set<ItemId>();
    for (const entry of entries) appendEntry('root', root, entry, noDisabled);

    const s = store.getState();
    const rowHidden =
      kind === 'row' &&
      hiddenInSelection(s.layout, 'row', s.selection.range.r0, s.selection.range.r1).length > 0;
    const colHidden =
      kind === 'col' &&
      hiddenInSelection(s.layout, 'col', s.selection.range.c0, s.selection.range.c1).length > 0;
    const rowUnhide = root.querySelector<HTMLButtonElement>('[data-fc-action="rowUnhide"]');
    const colUnhide = root.querySelector<HTMLButtonElement>('[data-fc-action="colUnhide"]');
    if (rowUnhide) {
      setContextMenuItemDisabled(rowUnhide, !rowHidden, strings.contextMenu.noHiddenRows);
    }
    if (colUnhide) {
      setContextMenuItemDisabled(colUnhide, !colHidden, strings.contextMenu.noHiddenColumns);
    }
  };

  const show = (clientX: number, clientY: number, kind: MenuKind): void => {
    inheritHostTokens(host, root);
    restoreFocusEl = document.activeElement instanceof HTMLElement ? document.activeElement : host;
    root.setAttribute('aria-label', strings.contextMenu.title);
    focusPanel = 'root';
    buildMenu(kind);
    if (pasteBtnRef) {
      const canPaste = canReadClipboard();
      setContextMenuItemDisabled(pasteBtnRef, !canPaste, strings.contextMenu.clipboardUnavailable);
    }
    root.style.display = 'block';
    root.style.left = '-9999px';
    root.style.top = '-9999px';
    visible = true;
    const { x, y } = clampPanelToViewport(root, clientX, clientY, { pad: VIEWPORT_PAD });
    root.style.left = `${x}px`;
    root.style.top = `${y}px`;
    focusMenuItem(0);
  };

  /** Resolve which menu flavour to show based on the click target. Header
   *  clicks promote the selection to the whole row/column so the action
   *  inherits a sensible band. */
  const resolveMenuKind = (e: MouseEvent): MenuKind => {
    const rect = hitHost.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;
    const s = store.getState();
    const zone = hitZone(s.layout, s.viewport, x, y, null, { resizeHandles: false });
    if (!zone) return 'cell';
    const selectedRanges = [s.selection.range, ...(s.selection.extraRanges ?? [])];
    if (zone.kind === 'row-header' || zone.kind === 'row-resize') {
      const inSel = selectedRanges.some(
        (sel) => zone.row >= sel.r0 && zone.row <= sel.r1 && sel.c0 === 0 && sel.c1 >= 16383,
      );
      if (!inSel) mutators.selectRow(store, zone.row);
      return 'row';
    }
    if (zone.kind === 'col-header' || zone.kind === 'col-resize') {
      const inSel = selectedRanges.some(
        (sel) => zone.col >= sel.c0 && zone.col <= sel.c1 && sel.r0 === 0 && sel.r1 >= 1048575,
      );
      if (!inSel) mutators.selectCol(store, zone.col);
      return 'col';
    }
    if (zone.kind === 'cell') {
      const selected = selectedRanges.find(
        (sel) =>
          zone.row >= sel.r0 && zone.row <= sel.r1 && zone.col >= sel.c0 && zone.col <= sel.c1,
      );
      if (selected?.c0 === 0 && selected.c1 >= 16383) return 'row';
      if (selected?.r0 === 0 && selected.r1 >= 1048575) return 'col';
    }
    return 'cell';
  };

  const isOwnChromeContextTarget = (target: EventTarget | null): boolean =>
    target instanceof Element &&
    !!target.closest('.fc-host__formulabar, .fc-host__sheetbar, .fc-sheetmenu');

  const onContextMenu = (e: MouseEvent): void => {
    if (isOwnChromeContextTarget(e.target)) return;
    e.preventDefault();
    const kind = resolveMenuKind(e);
    show(e.clientX, e.clientY, kind);
  };

  const insideMenu = (target: EventTarget | null): boolean =>
    target instanceof Node && (root.contains(target) || sub.contains(target));

  const onDocPointerDown = (e: MouseEvent): void => {
    if (!visible) return;
    if (insideMenu(e.target)) return;
    hide(false);
  };

  const onDocContextMenu = (e: MouseEvent): void => {
    if (!visible) return;
    if (insideMenu(e.target)) return;
    hide(false);
  };

  const onDocKey = (e: KeyboardEvent): void => {
    if (!visible) return;
    if (e.key === 'Escape') {
      e.preventDefault();
      if (openSub) {
        const parent = openSub.parentBtn;
        closeSubmenu();
        focusPanel = 'root';
        parent.focus();
        activeIndex = panelItems(root).indexOf(parent);
      } else {
        hide(true);
      }
    } else if (e.key === 'ArrowDown') {
      e.preventDefault();
      focusMenuItem(activeIndex + 1);
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      focusMenuItem(activeIndex - 1);
    } else if (e.key === 'Home') {
      e.preventDefault();
      focusMenuItem(0);
    } else if (e.key === 'End') {
      e.preventDefault();
      focusMenuItem(panelItems(activePanel()).length - 1);
    } else if (e.key === 'ArrowRight') {
      const el = document.activeElement;
      if (el instanceof HTMLButtonElement && el.dataset.fcSubmenu) {
        e.preventDefault();
        openSubmenu(el.dataset.fcSubmenu, el);
        focusPanel = 'sub';
        focusMenuItem(0);
      }
    } else if (e.key === 'ArrowLeft') {
      if (openSub && focusPanel === 'sub') {
        e.preventDefault();
        const parent = openSub.parentBtn;
        closeSubmenu();
        focusPanel = 'root';
        parent.focus();
        activeIndex = panelItems(root).indexOf(parent);
      }
    } else if (e.key === 'Enter' || e.key === ' ') {
      const target = document.activeElement;
      if (target instanceof HTMLButtonElement && insideMenu(target)) {
        e.preventDefault();
        target.click();
      }
    }
  };

  const onScroll = (): void => hide(false);

  sub.addEventListener('mouseenter', cancelSubClose);
  sub.addEventListener('mouseleave', scheduleSubClose);

  /** Clamp a selection's row span to the populated region. A whole-row /
   *  whole-column band selection spans ~1M rows; sorting or filtering that
   *  raw range would iterate (and rewrite) the entire sheet and freeze the
   *  UI, so bound it to the last populated row in the relevant columns. */
  const boundRowsToData = (range: Range): Range => {
    if (range.r1 - range.r0 < 50_000) return range;
    let maxRow = range.r0;
    const state = store.getState();
    const visitKey = (key: string): void => {
      const parts = key.split(':');
      if (parts.length !== 3 || Number(parts[0]) !== range.sheet) return;
      const row = Number(parts[1]);
      const col = Number(parts[2]);
      if (row < range.r0 || row > range.r1) return;
      if (col < range.c0 || col > range.c1) return;
      if (row > maxRow) maxRow = row;
    };
    for (const key of state.data.cells.keys()) visitKey(key);
    for (const [key, format] of state.format.formats) {
      if (Object.keys(format).length === 0) continue;
      visitKey(key);
    }
    return { ...range, r1: maxRow };
  };

  function runPasteSpecial(what: PasteWhat, transpose: boolean): void {
    const snap = deps.getClipboardSnapshot?.();
    if (!snap) return;
    if (history) history.begin();
    try {
      pasteSpecial(store.getState(), store, wb, snap, {
        what,
        operation: 'none',
        skipBlanks: false,
        transpose,
      });
    } catch (err) {
      console.warn('formulon-cell: paste special failed', err);
    } finally {
      if (history) history.end();
    }
    deps.onAfterCommit?.();
  }

  const hasPastePayload = (text: string, snap: ClipboardSnapshot | null | undefined): boolean =>
    text.length > 0 || snap != null;

  function run(id: ItemId): void {
    const state = store.getState();
    switch (id) {
      case 'copy': {
        if (deps.onClipboardShortcut) {
          deps.onClipboardShortcut('copy');
          return;
        }
        const r = copy(state);
        if (r) {
          if (r.ranges) mutators.setCopyRanges(store, r.ranges);
          else mutators.setCopyRange(store, r.range);
          void writeClipboard(r.tsv);
        } else {
          mutators.setCopyRange(store, null);
        }
        return;
      }
      case 'cut': {
        if (deps.onClipboardShortcut) {
          deps.onClipboardShortcut('cut');
          return;
        }
        if (history) history.begin();
        let r: ReturnType<typeof cut> = null;
        try {
          r = cut(state, wb);
        } finally {
          if (history) history.end();
        }
        if (r) {
          mutators.setCopyRange(store, r.range);
          void writeClipboard(r.tsv);
        }
        deps.onAfterCommit?.();
        return;
      }
      case 'paste': {
        if (deps.onClipboardShortcut) {
          deps.onClipboardShortcut('paste');
          return;
        }
        void readClipboard().then((text) => {
          const snap = deps.getClipboardSnapshot?.();
          if (!hasPastePayload(text, snap)) return;
          if (history) history.begin();
          let r: ReturnType<typeof pasteTSV> | ReturnType<typeof pasteSpecial> = null;
          try {
            r = text
              ? pasteTSV(store.getState(), wb, text)
              : snap
                ? pasteSpecial(store.getState(), store, wb, snap, {
                    what: 'all',
                    operation: 'none',
                    skipBlanks: false,
                    transpose: false,
                  })
                : null;
          } finally {
            if (history) history.end();
          }
          if (r) {
            mutators.setCopyRange(store, null);
            mutators.setRange(store, r.writtenRange);
          }
          deps.onAfterCommit?.();
        });
        return;
      }
      case 'pasteSpecial': {
        deps.onPasteSpecial?.();
        return;
      }
      case 'pasteAll':
        runPasteSpecial('all', false);
        return;
      case 'pasteFormulas':
        runPasteSpecial('formulas', false);
        return;
      case 'pasteFormulasNumFmt':
        runPasteSpecial('formulas-and-numfmt', false);
        return;
      case 'pasteValues':
        runPasteSpecial('values', false);
        return;
      case 'pasteValuesNumFmt':
        runPasteSpecial('values-and-numfmt', false);
        return;
      case 'pasteFormatsOnly':
        runPasteSpecial('formats', false);
        return;
      case 'pasteTranspose':
        runPasteSpecial('all', true);
        return;
      case 'insertCopiedCells': {
        openInsertCopiedCellsDialog({
          strings,
          onSubmit: (direction) => {
            void readClipboard().then((text) => {
              const snap = deps.getClipboardSnapshot?.();
              if (!hasPastePayload(text, snap)) return;
              const r = insertCopiedCellsFromTSV(store, wb, history, text, direction, snap);
              if (r) {
                mutators.setCopyRange(store, null);
                mutators.setRange(store, r.writtenRange);
                deps.onAfterCommit?.();
              }
            });
          },
        });
        return;
      }
      case 'clear': {
        const range = state.selection.range;
        const sheet = range.sheet;
        for (const key of state.data.cells.keys()) {
          const parts = key.split(':');
          if (parts.length !== 3) continue;
          if (Number(parts[0]) !== sheet) continue;
          const row = Number(parts[1]);
          const col = Number(parts[2]);
          if (row < range.r0 || row > range.r1) continue;
          if (col < range.c0 || col > range.c1) continue;
          wb.setBlank({ sheet, row, col });
        }
        deps.onAfterCommit?.();
        return;
      }
      case 'bold': {
        wrapFmt(() => toggleBold(state, store));
        return;
      }
      case 'italic': {
        wrapFmt(() => toggleItalic(state, store));
        return;
      }
      case 'underline': {
        wrapFmt(() => toggleUnderline(state, store));
        return;
      }
      case 'alignLeft': {
        wrapFmt(() => setAlign(state, store, 'left'));
        return;
      }
      case 'alignCenter': {
        wrapFmt(() => setAlign(state, store, 'center'));
        return;
      }
      case 'alignRight': {
        wrapFmt(() => setAlign(state, store, 'right'));
        return;
      }
      case 'borders': {
        wrapFmt(() => cycleBorders(state, store));
        return;
      }
      case 'formatCells': {
        deps.onFormatDialog?.();
        return;
      }
      case 'defineName': {
        deps.onDefineName?.();
        return;
      }
      case 'filterClear': {
        const range = boundRowsToData(state.ui.filterRange ?? inferAutoFilterRange(state));
        recordFilterChange(history, store, () => clearFilter(store.getState(), store, range));
        deps.onAfterCommit?.();
        return;
      }
      case 'filterReapply': {
        recordFilterChange(history, store, () => reapplyFilters(store.getState(), store));
        deps.onAfterCommit?.();
        return;
      }
      case 'filterByValue': {
        const range = boundRowsToData(state.ui.filterRange ?? inferAutoFilterRange(state));
        const byCol = state.selection.active.col;
        const keep = filterValueKey(state.data.cells.get(addrKey(state.selection.active))?.value);
        const hidden = distinctValues(state, range, byCol).filter((k) => k !== keep);
        recordFilterChange(history, store, () =>
          applyValueFilter(store.getState(), store, range, byCol, hidden),
        );
        deps.onAfterCommit?.();
        return;
      }
      case 'sortAsc':
      case 'sortDesc': {
        const range = boundRowsToData(inferAutoFilterRange(state, state.selection.range));
        sortRange(state, store, wb, range, {
          byCol: state.selection.active.col,
          direction: id === 'sortAsc' ? 'asc' : 'desc',
          hasHeader: inferSortHasHeader(state, range),
        });
        deps.onAfterCommit?.();
        return;
      }
      case 'rowHeight':
      case 'colWidth': {
        return;
      }
      case 'selectAll': {
        mutators.selectAll(store);
        return;
      }
      case 'rowInsertAbove': {
        const r = state.selection.range;
        insertRows(store, wb, history, r.r0, r.r1 - r.r0 + 1);
        deps.onAfterCommit?.();
        return;
      }
      case 'rowInsertBelow': {
        const r = state.selection.range;
        insertRows(store, wb, history, r.r1 + 1, r.r1 - r.r0 + 1);
        deps.onAfterCommit?.();
        return;
      }
      case 'rowDelete': {
        const r = state.selection.range;
        deleteRows(store, wb, history, r.r0, r.r1 - r.r0 + 1);
        deps.onAfterCommit?.();
        return;
      }
      case 'rowHide': {
        const r = state.selection.range;
        hideRows(store, history, r.r0, r.r1);
        return;
      }
      case 'rowUnhide': {
        const r = state.selection.range;
        const targets = hiddenInSelection(state.layout, 'row', r.r0, r.r1);
        const first = targets[0];
        const last = targets[targets.length - 1];
        if (first === undefined || last === undefined) return;
        showRows(store, history, first, last);
        return;
      }
      case 'rowGroup': {
        const r = state.selection.range;
        groupRows(store, history, r.r0, r.r1);
        return;
      }
      case 'rowUngroup': {
        const r = state.selection.range;
        ungroupRows(store, history, r.r0, r.r1);
        return;
      }
      case 'colInsertLeft': {
        const r = state.selection.range;
        insertCols(store, wb, history, r.c0, r.c1 - r.c0 + 1);
        deps.onAfterCommit?.();
        return;
      }
      case 'colInsertRight': {
        const r = state.selection.range;
        insertCols(store, wb, history, r.c1 + 1, r.c1 - r.c0 + 1);
        deps.onAfterCommit?.();
        return;
      }
      case 'colDelete': {
        const r = state.selection.range;
        deleteCols(store, wb, history, r.c0, r.c1 - r.c0 + 1);
        deps.onAfterCommit?.();
        return;
      }
      case 'colHide': {
        const r = state.selection.range;
        hideCols(store, history, r.c0, r.c1);
        return;
      }
      case 'colUnhide': {
        const r = state.selection.range;
        const targets = hiddenInSelection(state.layout, 'col', r.c0, r.c1);
        const first = targets[0];
        const last = targets[targets.length - 1];
        if (first === undefined || last === undefined) return;
        showCols(store, history, first, last);
        return;
      }
      case 'colGroup': {
        const r = state.selection.range;
        groupCols(store, history, r.c0, r.c1);
        return;
      }
      case 'colUngroup': {
        const r = state.selection.range;
        ungroupCols(store, history, r.c0, r.c1);
        return;
      }
      case 'insertComment': {
        deps.onEditComment?.(state.selection.active);
        return;
      }
      case 'deleteComment': {
        const addr = state.selection.active;
        wrapFmt(() => clearComment(store, addr, wb));
        return;
      }
      case 'insertHyperlink': {
        deps.onInsertHyperlink?.();
        return;
      }
      case 'toggleWatch': {
        deps.onToggleWatch?.(state.selection.active);
        return;
      }
    }
  }

  host.addEventListener('contextmenu', onContextMenu);
  document.addEventListener('contextmenu', onDocContextMenu, true);
  document.addEventListener('mousedown', onDocPointerDown, true);
  document.addEventListener('keydown', onDocKey, true);
  window.addEventListener('scroll', onScroll, true);

  const detach = ((): void => {
    host.removeEventListener('contextmenu', onContextMenu);
    document.removeEventListener('contextmenu', onDocContextMenu, true);
    document.removeEventListener('mousedown', onDocPointerDown, true);
    document.removeEventListener('keydown', onDocKey, true);
    window.removeEventListener('scroll', onScroll, true);
    cancelSubClose();
    root.remove();
    sub.remove();
  }) as ContextMenuHandle;
  detach.setStrings = (next: Strings): void => {
    strings = next;
    hide();
  };
  return detach;
}

function canReadClipboard(): boolean {
  return typeof navigator !== 'undefined' && typeof navigator.clipboard?.readText === 'function';
}

function setContextMenuItemDisabled(
  button: HTMLButtonElement,
  disabled: boolean,
  reason: string | null,
): void {
  projectDisabledState(button, disabled, reason, {
    datasetKey: 'disabledReason',
  });
}

async function writeClipboard(text: string): Promise<void> {
  if (typeof navigator === 'undefined' || typeof navigator.clipboard?.writeText !== 'function') {
    return;
  }
  try {
    await navigator.clipboard.writeText(text);
  } catch (err) {
    console.warn('formulon-cell: clipboard write failed', err);
  }
}

async function readClipboard(): Promise<string> {
  if (!canReadClipboard()) return '';
  try {
    return await navigator.clipboard.readText();
  } catch (err) {
    console.warn('formulon-cell: clipboard read failed', err);
    return '';
  }
}
