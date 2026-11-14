import { inputLocaleForWorkbook } from '../commands/coerce-input.js';
import {
  convertTextToNumber,
  numberStoredAsText,
  staleValuesMarked,
} from '../commands/error-indicators.js';
import { addrKey } from '../engine/address.js';
import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { focusMenuItem, handleMenuKeydown, projectDisabledState } from '../toolbar/menu-a11y.js';
import { createInteractionButton } from './chip-button.js';
import { onOutsidePress } from './outside-press.js';
import { overlayPortalFor } from './overlay-portal.js';
import { placePanel } from './overlay-position.js';

export type ErrorMenuKind = 'error' | 'validation';

export interface ErrorMenuDeps {
  /** Element the popover is appended to. The popover uses `position: fixed`
   *  with viewport-relative coordinates, so the host only matters for
   *  scoping/teardown. */
  host: HTMLElement;
  store: SpreadsheetStore;
  /** Workbook getter — lazy so the menu stays in lockstep with `setWorkbook`
   *  swaps. Used to read the cell formula for the "Edit cell" entry. */
  getWb: () => WorkbookHandle | null;
  strings?: Strings;
  /** Hook to focus the formula bar with the cell's current formula/value.
   *  When omitted, "Edit cell" is a no-op. The mount layer wires this to
   *  the formulabar input. */
  onEditCell?: (addr: Addr) => void;
  /** Hook to draw trace arrows for the error cell. The mount layer wires
   *  this to same-sheet precedent tracing. */
  onTraceError?: (addr: Addr, kind: ErrorMenuKind) => void;
  /** Hook for the stale rule's Calculate Now and Switch to Automatic Calculation. */
  onCalculate?: (action: 'calculate-now' | 'auto') => void;
}

export interface ErrorMenuHandle {
  open(addr: Addr, screenX: number, screenY: number, kind: ErrorMenuKind): void;
  close(): void;
  detach(): void;
}

export function attachErrorMenu(deps: ErrorMenuDeps): ErrorMenuHandle {
  const { host, store, getWb } = deps;
  const strings = deps.strings ?? defaultStrings;

  const root = document.createElement('div');
  root.className = 'fc-errmenu';
  root.setAttribute('role', 'menu');
  root.style.display = 'none';
  root.tabIndex = -1;
  overlayPortalFor(host).appendChild(root);

  let visible = false;
  let detachOutside: (() => void) | null = null;
  let restoreFocusEl: HTMLElement | null = null;

  const close = (restoreFocus = false): void => {
    if (!visible) return;
    visible = false;
    root.style.display = 'none';
    root.replaceChildren();
    const focusTarget = restoreFocusEl;
    restoreFocusEl = null;
    detachOutside?.();
    detachOutside = null;
    document.removeEventListener('keydown', onDocKey, true);
    window.removeEventListener('scroll', onScroll, true);
    if (restoreFocus) {
      (focusTarget ?? host).focus({ preventScroll: true });
    }
  };

  const onDocKey = (e: KeyboardEvent): void => {
    if (!visible) return;
    handleMenuKeydown(e, root, { close, items: '.fc-errmenu__item' });
  };

  const onScroll = (): void => close(false);

  type ErrorMenuAction =
    | 'showInfo'
    | 'editCell'
    | 'traceError'
    | 'ignore'
    | 'convert'
    | 'calculateNow'
    | 'switchToAutomatic'
    | 'disableFormatStale'
    | 'options';
  type ErrorMenuEntry = { id: ErrorMenuAction; label: string; unavailable?: string };

  const fcHost = host.closest<HTMLElement>('.fc-host') ?? host;
  const mac = fcHost.dataset.fcPlatform === 'mac';

  const createErrorMenuItemButton = (entry: ErrorMenuEntry): HTMLButtonElement => {
    const button = createInteractionButton({
      className: 'fc-errmenu__item',
      dataset: { fcAction: entry.id },
      role: 'menuitem',
      text: entry.label,
    });
    if (entry.unavailable) {
      projectDisabledState(button, true, entry.unavailable, { titlePrefix: entry.label });
    }
    return button;
  };

  /** The error-check rule a cell breaks, for the rules the Mac menu words itself. */
  const macRuleOf = (addr: Addr): 'divZero' | 'numberAsText' | 'stale' | null => {
    const state = store.getState();
    const cell = state.data.cells.get(addrKey(addr));
    if (!cell) return null;
    const wb = getWb();
    if (
      cell.formula &&
      staleValuesMarked(wb, state.ui.formatStaleValues) &&
      wb?.isCellStale(addr)
    ) {
      return 'stale';
    }
    if (cell.value.kind === 'error' && cell.value.text === '#DIV/0!') return 'divZero';
    let locale: string | undefined;
    try {
      locale = inputLocaleForWorkbook(getWb() as WorkbookHandle);
    } catch {
      locale = undefined;
    }
    return numberStoredAsText(cell.value, cell.formula, locale) ? 'numberAsText' : null;
  };

  /** Format a short "{code} — {value}" preview for the heading row. Falls
   *  back to the raw cell value when the cell isn't currently an error
   *  (validation case). */
  const describeCell = (addr: Addr): string => {
    const s = store.getState();
    const cell = s.data.cells.get(addrKey(addr));
    if (!cell) return '';
    const v = cell.value;
    switch (v.kind) {
      case 'error':
        return v.text;
      case 'number':
        return String(v.value);
      case 'bool':
        return v.value ? 'TRUE' : 'FALSE';
      case 'text':
        return v.value;
      default:
        return '';
    }
  };

  const buildMenu = (addr: Addr, kind: ErrorMenuKind): void => {
    const t = strings.errorMenu;
    root.replaceChildren();

    const heading = document.createElement('div');
    heading.className = 'fc-errmenu__heading';
    const m = strings.mac.errorMenu;
    const rule = mac && kind === 'error' ? macRuleOf(addr) : null;
    const ruleHeadings = {
      divZero: m.divZeroHeading,
      numberAsText: m.numberAsTextHeading,
      stale: m.staleHeading,
    };
    const label = rule
      ? ruleHeadings[rule]
      : kind === 'validation'
        ? t.validationHeading
        : t.errorHeading;
    root.setAttribute('aria-label', label);
    const detail = rule ? '' : describeCell(addr);
    heading.textContent = detail ? `${label} — ${detail}` : label;
    heading.classList.toggle('fc-errmenu__heading--rule', rule !== null);
    root.appendChild(heading);

    const entries: ErrorMenuEntry[] =
      rule === 'stale'
        ? [
            { id: 'showInfo', label: m.help },
            { id: 'calculateNow', label: m.calculateNow },
            { id: 'switchToAutomatic', label: m.switchToAutomatic },
            { id: 'disableFormatStale', label: m.disableFormatStale },
            { id: 'ignore', label: m.ignore },
            { id: 'editCell', label: m.editInFormulaBar },
            { id: 'options', label: m.options, unavailable: m.optionsUnavailable },
          ]
        : rule
          ? [
              ...(rule === 'numberAsText'
                ? [{ id: 'convert' as const, label: m.convertToNumber }]
                : []),
              { id: 'showInfo', label: m.help },
              ...(rule === 'divZero' ? [{ id: 'traceError' as const, label: m.trace }] : []),
              { id: 'ignore', label: m.ignore },
              { id: 'editCell', label: m.editInFormulaBar },
              { id: 'options', label: m.options, unavailable: m.optionsUnavailable },
            ]
          : [
              { id: 'showInfo', label: t.showInfo },
              { id: 'editCell', label: t.editCell },
              { id: 'traceError', label: t.traceError },
              { id: 'ignore', label: t.ignore },
            ];

    for (const entry of entries) {
      const btn = createErrorMenuItemButton(entry);
      btn.addEventListener('click', (ev) => {
        ev.preventDefault();
        ev.stopPropagation();
        run(entry.id, addr, kind);
        close(false);
      });
      root.appendChild(btn);
    }
  };

  const run = (id: ErrorMenuAction, addr: Addr, kind: ErrorMenuKind): void => {
    switch (id) {
      case 'showInfo': {
        // Surface the error/value detail through aria-live for non-blocking
        // feedback. The chrome layer mirrors selection updates into the
        // host's `.fc-host__a11y-live` element; we re-use it implicitly by just
        // updating the heading and emitting a CustomEvent that consumers
        // can hook for richer surfaces. The default in v1 is to keep the
        // menu close + emit the event; no modal popup.
        const detail = describeCell(addr);
        host.dispatchEvent(
          new CustomEvent('fc:errorinfo', {
            bubbles: true,
            detail: { addr, kind, message: detail },
          }),
        );
        return;
      }
      case 'editCell': {
        if (deps.onEditCell) {
          deps.onEditCell(addr);
        } else {
          // Best-effort fallback — focus the cell so the user can press F2.
          mutators.setActive(store, addr);
        }
        return;
      }
      case 'traceError': {
        deps.onTraceError?.(addr, kind);
        host.dispatchEvent(
          new CustomEvent('fc:traceerror', { bubbles: true, detail: { addr, kind } }),
        );
        return;
      }
      case 'ignore': {
        mutators.ignoreError(store, addr);
        return;
      }
      case 'convert': {
        const wb = getWb();
        if (wb) convertTextToNumber(store, wb, addr);
        return;
      }
      case 'calculateNow':
        deps.onCalculate?.('calculate-now');
        return;
      case 'switchToAutomatic':
        deps.onCalculate?.('auto');
        return;
      case 'disableFormatStale':
        mutators.setFormatStaleValues(store, false);
        return;
      case 'options':
        return;
    }
  };

  const api: ErrorMenuHandle = {
    open(addr: Addr, screenX: number, screenY: number, kind: ErrorMenuKind): void {
      restoreFocusEl =
        document.activeElement instanceof HTMLElement ? document.activeElement : host;
      buildMenu(addr, kind);
      // Pre-position offscreen so we can measure offsetWidth/offsetHeight.
      root.style.display = 'block';
      root.style.left = '-9999px';
      root.style.top = '-9999px';
      visible = true;
      placePanel(root, screenX, screenY);
      detachOutside?.();
      detachOutside = onOutsidePress(
        document,
        () => [root],
        () => {
          if (!visible) return;
          close(false);
        },
        { event: 'mousedown' },
      );
      document.addEventListener('keydown', onDocKey, true);
      window.addEventListener('scroll', onScroll, true);
      focusMenuItem(root, 'first', '.fc-errmenu__item');
    },
    close,
    detach(): void {
      close();
      root.remove();
    },
  };

  return api;
}
