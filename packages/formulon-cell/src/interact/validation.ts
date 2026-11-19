import { type CellInputAlert, commitCellInput } from '../commands/cell-input-commit.js';
import { blockDataTableCells } from '../commands/data-table-guard.js';
import type { RefusalLike } from '../commands/permission-reasons.js';
import {
  deferForAllowedRangeUnlock,
  isCellWritable,
  warnProtected,
} from '../commands/protection.js';
import { resolveListValues } from '../commands/validate.js';
import { addrKey } from '../engine/address.js';
import { fontPointsToCssPx } from '../engine/font-units.js';
import { makeRangeResolver } from '../engine/range-resolver.js';
import { formatBoolean } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { isMacHost } from '../extensions/ui-options.js';
import { cellRect, layoutForView } from '../render/geometry.js';
import { gridHitState } from '../render/grid.js';
import { validationAt } from '../store/range-validations.js';
import { renderScale } from '../store/render-scale.js';
import { mutators, type SpreadsheetStore, type State } from '../store/store.js';
import { pickFromListValues } from './autocomplete.js';
import { appendDialogButton, createDialogShell, returnFocus } from './dialog-shell.js';
import { onOutsidePress } from './outside-press.js';
import { overlayPortalFor } from './overlay-portal.js';

export interface ValidationListDeps {
  /** The grid surface that paints the chevron and receives clicks. */
  grid: HTMLElement;
  store: SpreadsheetStore;
  wb: WorkbookHandle;
  /** Refresh cached cells after a write — same contract as InlineEditor. */
  onAfterCommit: () => void;
}

export interface ValidationListHandle {
  /** Opens Pick From Drop-down List on the active cell; false when the column offers nothing. */
  openPickList(): boolean;
  detach(): void;
}

export interface ValidationPromptDeps {
  grid: HTMLElement;
  store: SpreadsheetStore;
}

export interface ValidationPromptHandle {
  refresh(): void;
  detach(): void;
}

export interface ValidationAlertLabels {
  ok: string;
  stop: string;
  warning: string;
  information: string;
  /** Wording of the `protectedCell` alert. */
  protectedCell: string;
  /** Localized text of an alert's refusal. */
  reason: (refusal: RefusalLike) => string;
}

export interface ValidationAlertDeps {
  host: HTMLElement;
  labels: ValidationAlertLabels;
}

export interface ValidationAlertMessage {
  severity: 'stop' | 'warning' | 'information';
  title?: string;
  message: string;
  code?: CellInputAlert['code'];
  refusal?: CellInputAlert['refusal'];
}

export interface ValidationAlertHandle {
  show(message: ValidationAlertMessage): void;
  detach(): void;
}

/** The active value of a cell as the list shows it selected. */
const cellText = (s: State, row: number, col: number): string => {
  const current = s.data.cells.get(addrKey({ sheet: s.data.sheetIndex, row, col }))?.value;
  if (current?.kind === 'text') return current.value;
  if (current?.kind === 'number') return String(current.value);
  if (current?.kind === 'bool') return formatBoolean(current.value);
  return '';
};

/**
 * Owns the click-to-open list dropdown for cells with `validation.kind === 'list'`.
 * Hit-tests against the chevron rect surfaced by the renderer; clicking opens
 * a popover, picking writes the chosen value to the cell.
 */

export function attachValidationList(deps: ValidationListDeps): ValidationListHandle {
  const { grid, store, wb } = deps;
  let popover: HTMLDivElement | null = null;
  let restoreFocus: HTMLElement | null = null;
  let detachOutside: (() => void) | null = null;

  const close = (shouldRestoreFocus = false): void => {
    if (!popover) return;
    const popoverEl = popover;
    popover.remove();
    popover = null;
    detachOutside?.();
    detachOutside = null;
    document.removeEventListener('keydown', onDocKey, true);
    const opener = restoreFocus;
    restoreFocus = null;
    if (!shouldRestoreFocus) return;
    returnFocus(opener, popoverEl, grid.closest<HTMLElement>('.fc-host') ?? grid);
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
    const addr = { sheet, row, col };
    const cell = { sheet, r0: row, c0: col, r1: row, c1: col };
    if (deferForAllowedRangeUnlock(store, [cell], () => commitValue(row, col, value))) {
      close();
      return;
    }
    if (!isCellWritable(store.getState(), addr)) {
      warnProtected(addr);
      close(true);
      return;
    }
    if (blockDataTableCells(wb, [addr])) {
      close();
      return;
    }
    // The pick is a typed entry: one authorized, undoable cell write.
    const result = commitCellInput({ store, wb, addr, raw: value, origin: 'editor' });
    if (result.status === 'failed') console.warn('formulon-cell: list pick failed', result.error);
    deps.onAfterCommit();
    close(true);
  };

  const open = (row: number, col: number, list: string[], currentValue = ''): void => {
    close();
    if (list.length === 0) return;
    const rect = grid.getBoundingClientRect();
    // A list validation hangs from its chevron; Pick From Drop-down List from the cell itself.
    const painted = gridHitState(grid).validationChevron;
    const chevron = painted && painted.row === row && painted.col === col ? painted : null;
    const state = store.getState();
    const cell = chevron?.cell ?? cellRect(layoutForView(state), state.viewport, row, col);

    const div = document.createElement('div');
    div.className = 'fc-validation-list';
    div.style.position = 'fixed';
    // The Mac desktop hangs the list from the cell, 4 pt wider, in 19 pt rows.
    const macPt = isMacHost(grid)
      ? fontPointsToCssPx(1) * renderScale(store.getState().layout, store.getState().viewport)
      : undefined;
    if (macPt !== undefined || !chevron) {
      div.style.left = `${rect.left + cell.x}px`;
      div.style.top = `${rect.top + cell.y + cell.h}px`;
      if (macPt !== undefined) div.style.width = `${cell.w + 4 * macPt}px`;
    } else {
      div.style.left = `${rect.left + chevron.rect.x}px`;
      div.style.top = `${rect.top + chevron.rect.y + chevron.rect.h}px`;
    }
    div.setAttribute('role', 'listbox');
    div.tabIndex = -1;
    restoreFocus =
      document.activeElement instanceof HTMLElement && document.activeElement !== document.body
        ? document.activeElement
        : grid;

    const selectedIndex = Math.max(0, list.indexOf(currentValue));
    for (const [idx, v] of list.entries()) {
      const item = document.createElement('div');
      item.className = 'fc-validation-list__item';
      item.setAttribute('role', 'option');
      item.setAttribute('aria-selected', idx === selectedIndex ? 'true' : 'false');
      item.tabIndex = idx === selectedIndex ? 0 : -1;
      item.textContent = v;
      if (macPt !== undefined) item.style.height = `${19 * macPt}px`;
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

    overlayPortalFor(grid).appendChild(div);
    popover = div;
    focusItem(selectedIndex);
    detachOutside = onOutsidePress(
      document,
      () => [popover],
      () => {
        if (!popover) return;
        close();
      },
      { event: 'mousedown' },
    );
    document.addEventListener('keydown', onDocKey, true);
  };

  const openForCell = (row: number, col: number): boolean => {
    const s = store.getState();
    const validation = validationAt(s.format, { sheet: s.data.sheetIndex, row, col });
    if (validation?.kind !== 'list') return false;
    const values = resolveListValues(validation, makeRangeResolver(wb, s.data.sheetIndex));
    open(row, col, values, cellText(s, row, col));
    return values.length > 0;
  };

  const onDown = (e: PointerEvent): void => {
    if (e.button !== 0) return;
    const chevron = gridHitState(grid).validationChevron;
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
    const validation = validationAt(s.format, {
      sheet: s.data.sheetIndex,
      row: chevron.row,
      col: chevron.col,
    });
    if (validation?.kind !== 'list') return;
    e.preventDefault();
    e.stopPropagation();
    // Re-anchor the active cell to the chevron's cell so subsequent picks
    //  hit the same target.
    mutators.setActive(store, { sheet: s.data.sheetIndex, row: chevron.row, col: chevron.col });
    openForCell(chevron.row, chevron.col);
  };

  /** Pick From Drop-down List: the column's entries around the active cell. */
  const openPickList = (): boolean => {
    const s = store.getState();
    const { row, col } = s.selection.active;
    const values = pickFromListValues(wb.cells(s.data.sheetIndex), col, row);
    open(row, col, values, cellText(s, row, col));
    return values.length > 0;
  };

  const onGridKeyDown = (e: KeyboardEvent): void => {
    if (!e.altKey || e.key !== 'ArrowDown') return;
    const active = store.getState().selection.active;
    // A list validation opens its own list; any other cell opens Pick From Drop-down List.
    if (!openForCell(active.row, active.col) && !openPickList()) return;
    e.preventDefault();
    e.stopPropagation();
  };

  // Capture phase so we beat the regular pointer.ts handler.
  grid.addEventListener('pointerdown', onDown, true);
  grid.addEventListener('keydown', onGridKeyDown, true);

  return {
    openPickList,
    detach() {
      close();
      grid.removeEventListener('pointerdown', onDown, true);
      grid.removeEventListener('keydown', onGridKeyDown, true);
    },
  };
}

/** Shows Excel-style data-validation input messages when the active cell has
 * prompt metadata. The prompt follows selection/viewport changes and stays
 * passive so it never steals keyboard focus from the sheet. */
export function attachValidationPrompt(deps: ValidationPromptDeps): ValidationPromptHandle {
  const { grid, store } = deps;
  let prompt: HTMLDivElement | null = null;

  const ensurePrompt = (): HTMLDivElement => {
    if (prompt) return prompt;
    const div = document.createElement('div');
    div.className = 'fc-validation-prompt';
    div.setAttribute('role', 'tooltip');
    div.hidden = true;
    const title = document.createElement('div');
    title.className = 'fc-validation-prompt__title';
    const body = document.createElement('div');
    body.className = 'fc-validation-prompt__body';
    div.append(title, body);
    overlayPortalFor(grid).appendChild(div);
    prompt = div;
    return div;
  };

  const hide = (): void => {
    if (prompt) prompt.hidden = true;
  };

  const refresh = (): void => {
    const state = store.getState();
    const addr = state.selection.active;
    const validation = validationAt(state.format, addr);
    const title = validation?.promptTitle?.trim() ?? '';
    const message = validation?.promptMessage?.trim() ?? '';
    if (!validation || validation.showInputMessage === false || (!title && !message)) {
      hide();
      return;
    }

    const gridBounds = grid.getBoundingClientRect();
    const rect = cellRect(layoutForView(state), state.viewport, addr.row, addr.col);
    // The card hangs from the cell's horizontal centre, 6px below its bottom edge.
    const left = gridBounds.left + rect.x + rect.w / 2;
    const top = gridBounds.top + rect.y + rect.h + 6;
    if (
      left < gridBounds.left ||
      left > gridBounds.right ||
      top < gridBounds.top ||
      top > gridBounds.bottom + 24
    ) {
      hide();
      return;
    }

    const div = ensurePrompt();
    const titleEl = div.querySelector<HTMLElement>('.fc-validation-prompt__title');
    const bodyEl = div.querySelector<HTMLElement>('.fc-validation-prompt__body');
    if (titleEl) {
      titleEl.textContent = title;
      titleEl.hidden = !title;
    }
    if (bodyEl) {
      bodyEl.textContent = message;
      bodyEl.hidden = !message;
    }
    div.style.left = `${Math.max(gridBounds.left, left)}px`;
    div.style.top = `${top}px`;
    div.hidden = false;
  };

  const unsub = store.subscribe(refresh);
  refresh();

  return {
    refresh,
    detach() {
      unsub();
      prompt?.remove();
      prompt = null;
    },
  };
}

export function attachValidationAlert(deps: ValidationAlertDeps): ValidationAlertHandle {
  const { host, labels } = deps;
  const shell = createDialogShell({
    host,
    className: 'fc-valdlg',
    ariaLabel: labels.stop,
    onDismiss: () => shell.close(),
  });
  shell.overlay.classList.add('fc-fmtdlg');
  shell.panel.classList.add('fc-fmtdlg__panel', 'fc-valdlg__panel');

  const header = document.createElement('div');
  header.className = 'fc-fmtdlg__header';
  shell.panel.appendChild(header);

  const body = document.createElement('div');
  body.className = 'fc-fmtdlg__body fc-tb__dlg__body';
  const messageEl = document.createElement('p');
  messageEl.className = 'fc-tb__dlg__message';
  body.appendChild(messageEl);
  shell.panel.appendChild(body);

  const footer = document.createElement('div');
  footer.className = 'fc-fmtdlg__footer';
  shell.panel.appendChild(footer);

  const okBtn = appendDialogButton(footer, {
    label: labels.ok,
    variant: 'primary',
  });

  const defaultTitle = (severity: ValidationAlertMessage['severity']): string => {
    if (severity === 'warning') return labels.warning;
    if (severity === 'information') return labels.information;
    return labels.stop;
  };

  const close = (): void => shell.close();
  shell.on(okBtn, 'click', close);

  return {
    show(message) {
      const protectedCell = message.code === 'protectedCell';
      const text = protectedCell
        ? labels.protectedCell
        : message.refusal
          ? labels.reason(message.refusal)
          : message.message;
      const title = message.title?.trim() || defaultTitle(message.severity);
      // The Mac desktop's protected-cell panel carries no title.
      const untitled = protectedCell && isMacHost(host);
      header.hidden = untitled;
      header.textContent = untitled ? '' : title;
      messageEl.textContent = text;
      shell.setAriaLabel(untitled ? text : title);
      shell.open();
      okBtn.focus({ preventScroll: true });
    },
    detach() {
      shell.dispose();
    },
  };
}
