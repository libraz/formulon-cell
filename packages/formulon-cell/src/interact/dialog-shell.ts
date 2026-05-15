import { enhanceCustomSelect, syncCustomSelects } from './custom-select.js';
import { inheritHostTokens } from './inherit-host-tokens.js';

export interface DialogShellDeps {
  /** Element the dialog inherits CSS variables from. The overlay itself is
   *  body-portaled so it escapes any `contain: strict` parents. */
  host: HTMLElement;
  /** Root class on the overlay (e.g. `fc-iterdlg`). */
  className: string;
  /** Accessible name; assigned to `aria-label` and refreshable via
   *  `setAriaLabel()` so locale swaps don't need a remount. */
  ariaLabel: string;
  /** Invoked on Escape (while open) and on backdrop click. If omitted both
   *  dismissals are inert — the dialog stays mounted regardless. */
  onDismiss?: () => void;
}

export interface DialogShell {
  /** Body-portaled overlay element (the backdrop). */
  readonly overlay: HTMLElement;
  /** Inner panel that holds the dialog content. */
  readonly panel: HTMLElement;
  /** Register an event listener that will be auto-removed on `dispose()`.
   *  Use this for every `addEventListener` the dialog wants tracked, including
   *  listeners on dynamically created child elements — the bag survives DOM
   *  rewrites because it stores `(target, event, handler)` tuples directly. */
  on<E extends EventTarget>(
    target: E,
    event: string,
    handler: EventListener,
    options?: AddEventListenerOptions | boolean,
  ): void;
  /** Show the overlay. */
  open(): void;
  /** Hide the overlay (no listener teardown). */
  close(): void;
  /** Refresh the overlay's accessible name (locale swap / dynamic title). */
  setAriaLabel(label: string): void;
  /** Idempotent. Detaches every registered listener AND the overlay. After
   *  dispose, the shell is unusable. */
  dispose(): void;
  /** True while the overlay is visible. */
  isOpen(): boolean;
}

interface BoundListener {
  target: EventTarget;
  event: string;
  handler: EventListener;
  options?: AddEventListenerOptions | boolean;
}

const FOCUSABLE_SELECTOR = [
  'button',
  'input',
  'select',
  'textarea',
  'a[href]',
  '[tabindex]:not([tabindex="-1"])',
].join(',');

function focusableInside(root: HTMLElement): HTMLElement[] {
  return Array.from(root.querySelectorAll<HTMLElement>(FOCUSABLE_SELECTOR)).filter((el) => {
    if (el.closest('[hidden],[aria-hidden="true"]')) return false;
    if ('disabled' in el && (el as HTMLButtonElement | HTMLInputElement).disabled) return false;
    return el.tabIndex >= 0;
  });
}

/**
 * Shared lifecycle and listener-bookkeeping primitive for the spreadsheet's
 * 16+ modal dialogs. Each dialog used to hand-roll its own overlay creation,
 * `role`/`aria-modal` wiring, Escape + backdrop dismissal, and a bespoke
 * `removeEventListener` ladder in `detach()`. Drift in those ladders was the
 * single largest source of pre-existing listener leaks (e.g. `fx-dialog`
 * tracked 9 add / 7 remove). This helper centralizes the pattern:
 *
 *  1. Mounts a body-portaled `<div role=dialog aria-modal>` with the requested
 *     class, marked `hidden` so callers control reveal timing.
 *  2. Inherits host CSS tokens so themed colors/shadows render outside the
 *     `.fc-host` `contain: strict` boundary.
 *  3. Keeps Tab / Shift+Tab focus inside the visible dialog, matching desktop
 *     modal behavior.
 *  4. Hooks Escape (document keydown, gated on `!hidden`) and backdrop click
 *     to a single `onDismiss` callback.
 *  5. Provides `on()` to register listeners with auto-cleanup. Every `on()`
 *     call is matched by a `removeEventListener` in `dispose()`, so callers
 *     never have to maintain their own bag.
 *
 * Existing dialogs migrate by replacing `overlay.addEventListener(...)` with
 * `shell.on(overlay, ...)` and `detach()` with `shell.dispose()`. The two
 * built-in dismissal listeners (Escape + backdrop) are tracked alongside
 * caller-registered listeners so dispose teardown is a single sweep.
 */
export function createDialogShell(deps: DialogShellDeps): DialogShell {
  const { host, className, onDismiss } = deps;
  let ariaLabel = deps.ariaLabel;
  let disposed = false;
  let restoreFocusEl: HTMLElement | null = null;
  const listeners: BoundListener[] = [];
  const selectHandles: Array<{ dispose(): void }> = [];

  const overlay = document.createElement('div');
  overlay.className = className;
  overlay.setAttribute('role', 'dialog');
  overlay.setAttribute('aria-modal', 'true');
  overlay.setAttribute('aria-label', ariaLabel);
  overlay.hidden = true;

  const panel = document.createElement('div');
  panel.className = `${className}__panel`;
  panel.tabIndex = -1;
  overlay.appendChild(panel);

  const enhanceSelects = (root: ParentNode): void => {
    for (const select of Array.from(root.querySelectorAll<HTMLSelectElement>('select'))) {
      const handle = enhanceCustomSelect(select);
      if (handle) selectHandles.push(handle);
    }
  };

  const selectObserver = new MutationObserver((records) => {
    for (const record of records) {
      for (const node of Array.from(record.addedNodes)) {
        if (!(node instanceof HTMLElement)) continue;
        if (node instanceof HTMLSelectElement) {
          const handle = enhanceCustomSelect(node);
          if (handle) selectHandles.push(handle);
        }
        enhanceSelects(node);
      }
    }
  });
  selectObserver.observe(panel, { childList: true, subtree: true });

  inheritHostTokens(host, overlay);
  document.body.appendChild(overlay);

  function on(
    target: EventTarget,
    event: string,
    handler: EventListener,
    options?: AddEventListenerOptions | boolean,
  ): void {
    if (disposed) return;
    target.addEventListener(event, handler, options);
    listeners.push({ target, event, handler, options });
  }

  if (onDismiss) {
    on(overlay, 'click', (e) => {
      if ((e as MouseEvent).target === overlay) onDismiss();
    });
    on(document, 'keydown', (e) => {
      if (overlay.hidden) return;
      if ((e as KeyboardEvent).key === 'Escape') {
        e.preventDefault();
        onDismiss();
      }
    });
  }

  on(document, 'keydown', (e) => {
    const event = e as KeyboardEvent;
    if (overlay.hidden || event.key !== 'Tab') return;
    const items = focusableInside(panel);
    if (items.length === 0) {
      event.preventDefault();
      panel.focus({ preventScroll: true });
      return;
    }
    const active = document.activeElement;
    const first = items[0];
    const last = items[items.length - 1];
    if (!active || !overlay.contains(active)) {
      event.preventDefault();
      first?.focus({ preventScroll: true });
      return;
    }
    if (event.shiftKey && active === first) {
      event.preventDefault();
      last?.focus({ preventScroll: true });
      return;
    }
    if (!event.shiftKey && active === last) {
      event.preventDefault();
      first?.focus({ preventScroll: true });
    }
  });

  return {
    overlay,
    panel,
    on,
    open() {
      if (disposed) return;
      enhanceSelects(panel);
      syncCustomSelects(panel);
      if (!overlay.contains(document.activeElement)) {
        restoreFocusEl =
          document.activeElement instanceof HTMLElement ? document.activeElement : host;
      }
      overlay.hidden = false;
      if (!overlay.contains(document.activeElement)) {
        const first = focusableInside(panel)[0] ?? panel;
        first.focus({ preventScroll: true });
      }
    },
    close() {
      const shouldRestore =
        !overlay.hidden &&
        !!restoreFocusEl &&
        (overlay.contains(document.activeElement) || document.activeElement === document.body);
      overlay.hidden = true;
      if (shouldRestore) {
        restoreFocusEl?.focus({ preventScroll: true });
      }
      restoreFocusEl = null;
    },
    setAriaLabel(label: string) {
      ariaLabel = label;
      overlay.setAttribute('aria-label', label);
    },
    dispose() {
      if (disposed) return;
      disposed = true;
      for (const l of listeners) {
        l.target.removeEventListener(l.event, l.handler, l.options);
      }
      listeners.length = 0;
      selectObserver.disconnect();
      for (const handle of selectHandles.splice(0)) handle.dispose();
      restoreFocusEl = null;
      overlay.remove();
    },
    isOpen() {
      return !overlay.hidden;
    },
  };
}
