import { copy } from '../commands/clipboard/copy.js';
import { cut } from '../commands/clipboard/cut.js';
import { insertCopiedCellsFromTSV } from '../commands/clipboard/insert-copied-cells.js';
import { pasteTSV } from '../commands/clipboard/paste.js';
import { clearComment } from '../commands/comment.js';
import {
  clearFormat,
  cycleBorders,
  setAlign,
  toggleBold,
  toggleItalic,
  toggleUnderline,
} from '../commands/format.js';
import { type History, recordFormatChange } from '../commands/history.js';
import { groupCols, groupRows, ungroupCols, ungroupRows } from '../commands/outline.js';
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
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import { hitZone } from '../render/geometry.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { inheritHostTokens } from './inherit-host-tokens.js';
import { openInsertCopiedCellsDialog } from './insert-copied-cells-dialog.js';

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
  /** Called when the user clicks "Edit comment…". When omitted the menu
   *  entry is hidden — the action requires the comment dialog feature to
   *  be wired up. */
  onEditComment?: (addr: import('../engine/types.js').Addr) => void;
  /** Called when the user clicks "Insert hyperlink…". When omitted the menu
   *  entry is hidden (the action is purely UX sugar — Format Cells > More
   *  also covers it). */
  onInsertHyperlink?: () => void;
  /** Called when the user clicks the Add/Remove Watch entry. Receives the
   *  active cell address; the host decides whether to add or remove based
   *  on its own watch list. When omitted the menu entry is hidden. */
  onToggleWatch?: (addr: import('../engine/types.js').Addr) => void;
  /** Returns true when the active cell is currently watched. Used to flip
   *  the menu label between "Add Watch" and "Remove Watch". */
  isWatched?: (addr: import('../engine/types.js').Addr) => boolean;
}

type ItemId =
  | 'copy'
  | 'cut'
  | 'paste'
  | 'pasteSpecial'
  | 'insertCopiedCells'
  | 'clear'
  | 'bold'
  | 'italic'
  | 'underline'
  | 'alignLeft'
  | 'alignCenter'
  | 'alignRight'
  | 'borders'
  | 'clearFormat'
  | 'formatCells'
  | 'selectAll'
  | 'rowHeight'
  | 'colWidth'
  | 'rowInsertAbove'
  | 'rowInsertBelow'
  | 'rowDelete'
  | 'rowHide'
  | 'rowUnhide'
  | 'colInsertLeft'
  | 'colInsertRight'
  | 'colDelete'
  | 'colHide'
  | 'colUnhide'
  | 'rowGroup'
  | 'rowUngroup'
  | 'colGroup'
  | 'colUngroup'
  | 'insertComment'
  | 'deleteComment'
  | 'insertHyperlink'
  | 'toggleWatch';

type MenuKind = 'cell' | 'row' | 'col';

type MenuEntry =
  | { kind: 'item'; id: ItemId; label: string; hint?: string }
  | { kind: 'sep'; id: string };

function buildCellEntries(s: Strings): MenuEntry[] {
  const t = s.contextMenu;
  return [
    { kind: 'item', id: 'copy', label: t.copy, hint: '⌘C' },
    { kind: 'item', id: 'cut', label: t.cut, hint: '⌘X' },
    { kind: 'item', id: 'paste', label: t.paste, hint: '⌘V' },
    { kind: 'item', id: 'pasteSpecial', label: t.pasteSpecial, hint: '⌘⇧V' },
    { kind: 'item', id: 'insertCopiedCells', label: t.insertCopiedCells },
    { kind: 'item', id: 'clear', label: t.clear, hint: 'Del' },
    { kind: 'sep', id: 'sep1' },
    { kind: 'item', id: 'bold', label: t.bold, hint: '⌘B' },
    { kind: 'item', id: 'italic', label: t.italic, hint: '⌘I' },
    { kind: 'item', id: 'underline', label: t.underline, hint: '⌘U' },
    { kind: 'sep', id: 'sep2' },
    { kind: 'item', id: 'alignLeft', label: t.alignLeft },
    { kind: 'item', id: 'alignCenter', label: t.alignCenter },
    { kind: 'item', id: 'alignRight', label: t.alignRight },
    { kind: 'sep', id: 'sep3' },
    { kind: 'item', id: 'borders', label: t.borders },
    { kind: 'item', id: 'clearFormat', label: t.clearFormat },
    { kind: 'sep', id: 'sep4' },
    { kind: 'item', id: 'formatCells', label: t.formatCells, hint: '⌘1' },
    { kind: 'sep', id: 'sep5' },
    { kind: 'item', id: 'insertComment', label: t.insertComment, hint: '⇧F2' },
    { kind: 'item', id: 'deleteComment', label: t.deleteComment },
    { kind: 'item', id: 'insertHyperlink', label: t.insertHyperlink, hint: '⌘K' },
    { kind: 'sep', id: 'sep6' },
    { kind: 'item', id: 'toggleWatch', label: t.addWatch },
    { kind: 'sep', id: 'sep7' },
    { kind: 'item', id: 'selectAll', label: t.selectAll, hint: '⌘A' },
  ];
}

function buildRowEntries(s: Strings): MenuEntry[] {
  const t = s.contextMenu;
  return [
    { kind: 'item', id: 'cut', label: t.cut, hint: '⌘X' },
    { kind: 'item', id: 'copy', label: t.copy, hint: '⌘C' },
    { kind: 'item', id: 'paste', label: t.paste, hint: '⌘V' },
    { kind: 'sep', id: 'sepR1' },
    { kind: 'item', id: 'rowInsertAbove', label: t.insert },
    { kind: 'item', id: 'rowInsertBelow', label: t.rowInsertBelow },
    { kind: 'item', id: 'rowDelete', label: t.delete },
    { kind: 'item', id: 'clear', label: t.clear, hint: 'Del' },
    { kind: 'sep', id: 'sepR2' },
    { kind: 'item', id: 'formatCells', label: t.formatCells, hint: '⌘1' },
    { kind: 'item', id: 'rowHeight', label: t.rowHeight },
    { kind: 'item', id: 'rowHide', label: t.rowHide },
    { kind: 'item', id: 'rowUnhide', label: t.rowUnhide },
    { kind: 'sep', id: 'sepR3' },
    { kind: 'item', id: 'selectAll', label: t.selectAll, hint: '⌘A' },
  ];
}

function buildColEntries(s: Strings): MenuEntry[] {
  const t = s.contextMenu;
  return [
    { kind: 'item', id: 'cut', label: t.cut, hint: '⌘X' },
    { kind: 'item', id: 'copy', label: t.copy, hint: '⌘C' },
    { kind: 'item', id: 'paste', label: t.paste, hint: '⌘V' },
    { kind: 'item', id: 'pasteSpecial', label: t.pasteSpecial, hint: '⌘⇧V' },
    { kind: 'sep', id: 'sepC1' },
    { kind: 'item', id: 'colInsertLeft', label: t.insert },
    { kind: 'item', id: 'colInsertRight', label: t.colInsertRight },
    { kind: 'item', id: 'colDelete', label: t.delete },
    { kind: 'item', id: 'clear', label: t.clear, hint: 'Del' },
    { kind: 'sep', id: 'sepC2' },
    { kind: 'item', id: 'formatCells', label: t.formatCells, hint: '⌘1' },
    { kind: 'item', id: 'colWidth', label: t.colWidth },
    { kind: 'item', id: 'colHide', label: t.colHide },
    { kind: 'item', id: 'colUnhide', label: t.colUnhide },
    { kind: 'sep', id: 'sepC3' },
    { kind: 'item', id: 'selectAll', label: t.selectAll, hint: '⌘A' },
  ];
}

function compactMenuEntries(entries: MenuEntry[]): MenuEntry[] {
  const out: MenuEntry[] = [];
  for (const entry of entries) {
    if (entry.kind === 'sep') {
      const prev = out[out.length - 1];
      if (!prev || prev.kind === 'sep') continue;
      out.push(entry);
      continue;
    }
    out.push(entry);
  }
  while (out[out.length - 1]?.kind === 'sep') out.pop();
  return out;
}

const VIEWPORT_PAD = 4;

/** Detacher returned by `attachContextMenu`. Also exposes `setStrings` so the
 *  active dictionary can be swapped after attach. The function form is kept
 *  for backwards-compat with callers that just want `detach()`. */
export interface ContextMenuHandle {
  (): void;
  /** Swap the active dictionary; takes effect on next open. */
  setStrings(next: Strings): void;
}

export function attachContextMenu(deps: ContextMenuDeps): ContextMenuHandle {
  const { host, store, wb } = deps;
  const hitHost = deps.grid ?? host;
  const history = deps.history ?? null;
  let strings = deps.strings ?? defaultStrings;
  const wrapFmt = (fn: () => void): void => recordFormatChange(history, store, fn);

  const root = document.createElement('div');
  root.className = 'fc-ctxmenu';
  root.setAttribute('role', 'menu');
  root.setAttribute('aria-label', strings.contextMenu.title);
  root.style.display = 'none';
  root.tabIndex = -1;
  document.body.appendChild(root);

  // Theme-token bridge — see ./inherit-host-tokens.ts.

  let visible = false;
  let pasteBtnRef: HTMLButtonElement | null = null;
  let activeIndex = -1;
  let restoreFocusEl: HTMLElement | null = null;

  const hide = (restoreFocus = false): void => {
    if (!visible) return;
    visible = false;
    root.style.display = 'none';
    activeIndex = -1;
    const focusTarget = restoreFocusEl;
    restoreFocusEl = null;
    if (restoreFocus) {
      (focusTarget ?? host).focus({ preventScroll: true });
    }
  };

  const menuItems = (): HTMLButtonElement[] =>
    Array.from(root.querySelectorAll<HTMLButtonElement>('.fc-ctxmenu__item')).filter(
      (btn) => !btn.disabled && btn.getAttribute('aria-disabled') !== 'true',
    );

  const focusMenuItem = (idx: number): void => {
    const items = menuItems();
    if (items.length === 0) return;
    activeIndex = (idx + items.length) % items.length;
    items[activeIndex]?.focus();
  };

  const buildMenu = (kind: MenuKind): void => {
    root.replaceChildren();
    pasteBtnRef = null;
    const raw =
      kind === 'row'
        ? buildRowEntries(strings)
        : kind === 'col'
          ? buildColEntries(strings)
          : buildCellEntries(strings);
    // Hide entries the host has not opted into. `insertHyperlink` and
    // `toggleWatch` are optional — the rest are always available.
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
    for (const entry of entries) {
      if (entry.kind === 'sep') {
        const sep = document.createElement('hr');
        sep.className = 'fc-ctxmenu__sep';
        root.appendChild(sep);
        continue;
      }
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'fc-ctxmenu__item';
      btn.dataset.fcAction = entry.id;
      btn.setAttribute('role', 'menuitem');
      btn.tabIndex = -1;
      const label = document.createElement('span');
      label.className = 'fc-ctxmenu__label';
      label.textContent = entry.label;
      const hint = document.createElement('span');
      hint.className = 'fc-ctxmenu__hint';
      hint.textContent = entry.hint ?? '';
      btn.append(label, hint);
      btn.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        if (btn.disabled) return;
        run(entry.id);
        hide(false);
      });
      root.appendChild(btn);
      if (entry.id === 'paste') pasteBtnRef = btn;
    }

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
      rowUnhide.disabled = !rowHidden;
      rowUnhide.setAttribute('aria-disabled', rowHidden ? 'false' : 'true');
    }
    if (colUnhide) {
      colUnhide.disabled = !colHidden;
      colUnhide.setAttribute('aria-disabled', colHidden ? 'false' : 'true');
    }
  };

  const clampToViewport = (x: number, y: number): { x: number; y: number } => {
    const w = root.offsetWidth;
    const h = root.offsetHeight;
    const vw = window.innerWidth;
    const vh = window.innerHeight;
    const cx = Math.max(VIEWPORT_PAD, Math.min(x, vw - w - VIEWPORT_PAD));
    const cy = Math.max(VIEWPORT_PAD, Math.min(y, vh - h - VIEWPORT_PAD));
    return { x: cx, y: cy };
  };

  const show = (clientX: number, clientY: number, kind: MenuKind): void => {
    inheritHostTokens(host, root);
    restoreFocusEl = document.activeElement instanceof HTMLElement ? document.activeElement : host;
    root.setAttribute('aria-label', strings.contextMenu.title);
    buildMenu(kind);
    if (pasteBtnRef) {
      const canPaste = canReadClipboard();
      pasteBtnRef.disabled = !canPaste;
      if (!canPaste) pasteBtnRef.setAttribute('aria-disabled', 'true');
      else pasteBtnRef.removeAttribute('aria-disabled');
    }
    root.style.display = 'block';
    root.style.left = '-9999px';
    root.style.top = '-9999px';
    visible = true;
    const { x, y } = clampToViewport(clientX, clientY);
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
      // Promote selection to the row (preserving multi-row drags).
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

  const onDocPointerDown = (e: MouseEvent): void => {
    if (!visible) return;
    if (e.target instanceof Node && root.contains(e.target)) return;
    hide(false);
  };

  const onDocContextMenu = (e: MouseEvent): void => {
    if (!visible) return;
    if (e.target instanceof Node && root.contains(e.target)) return;
    hide(false);
  };

  const onDocKey = (e: KeyboardEvent): void => {
    if (!visible) return;
    if (e.key === 'Escape') {
      e.preventDefault();
      hide(true);
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
      focusMenuItem(menuItems().length - 1);
    } else if (e.key === 'Enter' || e.key === ' ') {
      const target = document.activeElement;
      if (target instanceof HTMLButtonElement && root.contains(target)) {
        e.preventDefault();
        target.click();
      }
    }
  };

  const onScroll = (): void => hide(false);

  function run(id: ItemId): void {
    const state = store.getState();
    switch (id) {
      case 'copy': {
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
        const r = cut(state, wb);
        if (r) {
          mutators.setCopyRange(store, r.range);
          void writeClipboard(r.tsv);
        }
        deps.onAfterCommit?.();
        return;
      }
      case 'paste': {
        void readClipboard().then((text) => {
          if (!text) return;
          const r = pasteTSV(store.getState(), wb, text);
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
      case 'insertCopiedCells': {
        openInsertCopiedCellsDialog({
          strings,
          onSubmit: (direction) => {
            void readClipboard().then((text) => {
              if (!text) return;
              const r = insertCopiedCellsFromTSV(store, wb, history, text, direction);
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
      case 'clearFormat': {
        wrapFmt(() => clearFormat(state, store));
        return;
      }
      case 'formatCells': {
        deps.onFormatDialog?.();
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
        // Desktop spreadsheets: select rows flanking a hidden band, then unhide. We just
        // unhide every hidden row inside the active selection.
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
    root.remove();
  }) as ContextMenuHandle;
  detach.setStrings = (next: Strings): void => {
    strings = next;
    // Menu is rebuilt on each show via buildMenu(kind), which reads
    // `strings` lazily — closing here ensures any open menu doesn't
    // keep stale labels.
    hide();
  };
  return detach;
}

function canReadClipboard(): boolean {
  return typeof navigator !== 'undefined' && typeof navigator.clipboard?.readText === 'function';
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
