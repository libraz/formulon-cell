import { writeInput } from '../commands/coerce-input.js';
import { isCellWritable, warnProtected } from '../commands/protection.js';
import { resolveListValues } from '../commands/validate.js';
import { addrKey } from '../engine/address.js';
import { makeRangeResolver } from '../engine/range-resolver.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { cellRect } from '../render/geometry.js';
import { getValidationChevron } from '../render/grid.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { createDialogShell } from './dialog-shell.js';
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
}

export interface ValidationAlertDeps {
  host: HTMLElement;
  labels: ValidationAlertLabels;
}

export interface ValidationAlertMessage {
  severity: 'stop' | 'warning' | 'information';
  title?: string;
  message: string;
}

export interface ValidationAlertHandle {
  show(message: ValidationAlertMessage): void;
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
    const addr = { sheet, row, col };
    if (!isCellWritable(store.getState(), addr)) {
      warnProtected(addr);
      close();
      return;
    }
    try {
      writeInput(wb, addr, value, store);
    } catch (err) {
      console.warn('formulon-cell: validation write failed', err);
    }
    deps.onAfterCommit();
    close();
  };

  const open = (row: number, col: number, list: string[], currentValue = ''): void => {
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

    const selectedIndex = Math.max(0, list.indexOf(currentValue));
    for (const [idx, v] of list.entries()) {
      const item = document.createElement('div');
      item.className = 'fc-validation-list__item';
      item.setAttribute('role', 'option');
      item.setAttribute('aria-selected', idx === selectedIndex ? 'true' : 'false');
      item.tabIndex = idx === selectedIndex ? 0 : -1;
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
    focusItem(selectedIndex);
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
    const current = s.data.cells.get(
      addrKey({ sheet: s.data.sheetIndex, row: chevron.row, col: chevron.col }),
    )?.value;
    const currentValue =
      current?.kind === 'text'
        ? current.value
        : current?.kind === 'number'
          ? String(current.value)
          : current?.kind === 'bool'
            ? current.value
              ? 'TRUE'
              : 'FALSE'
            : '';
    const values = resolveListValues(fmt.validation, makeRangeResolver(wb, s.data.sheetIndex));
    open(chevron.row, chevron.col, values, currentValue);
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
    inheritHostTokens(grid, div);
    document.body.appendChild(div);
    prompt = div;
    return div;
  };

  const hide = (): void => {
    if (prompt) prompt.hidden = true;
  };

  const refresh = (): void => {
    const state = store.getState();
    const addr = state.selection.active;
    const validation = state.format.formats.get(addrKey(addr))?.validation;
    const title = validation?.promptTitle?.trim() ?? '';
    const message = validation?.promptMessage?.trim() ?? '';
    if (!validation || validation.showInputMessage === false || (!title && !message)) {
      hide();
      return;
    }

    const gridBounds = grid.getBoundingClientRect();
    const rect = cellRect(state.layout, state.viewport, addr.row, addr.col);
    const left = gridBounds.left + rect.x;
    const top = gridBounds.top + rect.y + rect.h + 4;
    if (
      left + Math.min(rect.w, 24) < gridBounds.left ||
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
  body.className = 'fc-fmtdlg__body app__dlg__body';
  const messageEl = document.createElement('p');
  messageEl.className = 'app__dlg__message';
  body.appendChild(messageEl);
  shell.panel.appendChild(body);

  const footer = document.createElement('div');
  footer.className = 'fc-fmtdlg__footer';
  shell.panel.appendChild(footer);

  const okBtn = document.createElement('button');
  okBtn.type = 'button';
  okBtn.className = 'fc-fmtdlg__btn fc-fmtdlg__btn--primary';
  okBtn.textContent = labels.ok;
  footer.appendChild(okBtn);

  const defaultTitle = (severity: ValidationAlertMessage['severity']): string => {
    if (severity === 'warning') return labels.warning;
    if (severity === 'information') return labels.information;
    return labels.stop;
  };

  const close = (): void => shell.close();
  shell.on(okBtn, 'click', close);

  return {
    show(message) {
      const title = message.title?.trim() || defaultTitle(message.severity);
      header.textContent = title;
      messageEl.textContent = message.message;
      shell.setAriaLabel(title);
      shell.open();
      okBtn.focus({ preventScroll: true });
    },
    detach() {
      shell.dispose();
    },
  };
}
