import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { addrKey } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';

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
}

export interface ErrorMenuHandle {
  open(addr: Addr, screenX: number, screenY: number, kind: ErrorMenuKind): void;
  close(): void;
  detach(): void;
}

const VIEWPORT_PAD = 4;

export function attachErrorMenu(deps: ErrorMenuDeps): ErrorMenuHandle {
  const { host, store, getWb } = deps;
  const strings = deps.strings ?? defaultStrings;

  const root = document.createElement('div');
  root.className = 'fc-errmenu';
  root.setAttribute('role', 'menu');
  root.style.display = 'none';
  root.tabIndex = -1;
  document.body.appendChild(root);

  let visible = false;
  let currentAddr: Addr | null = null;
  let currentKind: ErrorMenuKind = 'error';

  const close = (): void => {
    if (!visible) return;
    visible = false;
    root.style.display = 'none';
    root.replaceChildren();
    currentAddr = null;
    document.removeEventListener('mousedown', onDocMouseDown, true);
    document.removeEventListener('keydown', onDocKey, true);
    window.removeEventListener('scroll', onScroll, true);
  };

  const onDocMouseDown = (e: MouseEvent): void => {
    if (!visible) return;
    if (e.target instanceof Node && root.contains(e.target)) return;
    close();
  };

  const onDocKey = (e: KeyboardEvent): void => {
    if (!visible) return;
    if (e.key === 'Escape') {
      e.preventDefault();
      close();
    }
  };

  const onScroll = (): void => close();

  const clampToViewport = (x: number, y: number): { x: number; y: number } => {
    const w = root.offsetWidth;
    const h = root.offsetHeight;
    const vw = window.innerWidth;
    const vh = window.innerHeight;
    const cx = Math.max(VIEWPORT_PAD, Math.min(x, vw - w - VIEWPORT_PAD));
    const cy = Math.max(VIEWPORT_PAD, Math.min(y, vh - h - VIEWPORT_PAD));
    return { x: cx, y: cy };
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
    const label = kind === 'validation' ? t.validationHeading : t.errorHeading;
    const detail = describeCell(addr);
    heading.textContent = detail ? `${label} — ${detail}` : label;
    root.appendChild(heading);

    type Entry = { id: 'showInfo' | 'editCell' | 'traceError' | 'ignore'; label: string };
    const entries: Entry[] = [
      { id: 'showInfo', label: t.showInfo },
      { id: 'editCell', label: t.editCell },
      { id: 'traceError', label: t.traceError },
      { id: 'ignore', label: t.ignore },
    ];

    for (const entry of entries) {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'fc-errmenu__item';
      btn.dataset.fcAction = entry.id;
      btn.setAttribute('role', 'menuitem');
      btn.textContent = entry.label;
      btn.addEventListener('click', (ev) => {
        ev.preventDefault();
        ev.stopPropagation();
        run(entry.id, addr, kind);
        close();
      });
      root.appendChild(btn);
    }
  };

  const run = (
    id: 'showInfo' | 'editCell' | 'traceError' | 'ignore',
    addr: Addr,
    kind: ErrorMenuKind,
  ): void => {
    switch (id) {
      case 'showInfo': {
        // Surface the error/value detail through aria-live for non-blocking
        // feedback. The chrome layer mirrors selection updates into the
        // host's `.fc-host__a11y` element; we re-use it implicitly by just
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
        // Placeholder for v1 — a separate trace-arrows feature handles this.
        // We still emit a CustomEvent so a future feature (or consumer
        // chrome) can wire trace UI without re-adding the menu entry.
        host.dispatchEvent(
          new CustomEvent('fc:traceerror', { bubbles: true, detail: { addr, kind } }),
        );
        return;
      }
      case 'ignore': {
        mutators.ignoreError(store, addr);
        return;
      }
    }
  };

  const api: ErrorMenuHandle = {
    open(addr: Addr, screenX: number, screenY: number, kind: ErrorMenuKind): void {
      currentAddr = addr;
      currentKind = kind;
      // Touch the workbook getter so consumers wiring Strict Mode can rely
      // on the same lazy resolution shape used elsewhere.
      void getWb();
      buildMenu(addr, kind);
      // Pre-position offscreen so we can measure offsetWidth/offsetHeight.
      root.style.display = 'block';
      root.style.left = '-9999px';
      root.style.top = '-9999px';
      visible = true;
      const { x, y } = clampToViewport(screenX, screenY);
      root.style.left = `${x}px`;
      root.style.top = `${y}px`;
      document.addEventListener('mousedown', onDocMouseDown, true);
      document.addEventListener('keydown', onDocKey, true);
      window.addEventListener('scroll', onScroll, true);
    },
    close,
    detach(): void {
      close();
      root.remove();
    },
  };

  // Expose a tiny debug surface — `currentAddr` / `currentKind` aren't part
  // of the public handle but tests can assert against the menu's DOM. The
  // assignments here also keep TS from flagging them as unused locals.
  void currentAddr;
  void currentKind;

  return api;
}
