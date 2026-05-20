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

export interface DialogFrameOptions {
  title: string;
  bodyTag?: 'div' | 'form';
  panelClasses?: readonly string[];
  headerClass?: string;
  bodyClass?: string;
  footerClass?: string;
}

export interface DialogFrame {
  header: HTMLDivElement;
  body: HTMLDivElement | HTMLFormElement;
  footer: HTMLDivElement;
}

export interface DialogButtonOptions {
  label: string;
  variant?: 'primary' | 'secondary';
  baseClass?: string;
  primaryClass?: string;
  secondaryClass?: string;
}

export interface DialogIconButtonOptions {
  label: string;
  ariaLabel: string;
  baseClass: string;
  title?: string;
  html?: string;
}

export interface DialogTabPairOptions {
  id: string;
  label: string;
  tabId: string;
  panelId: string;
  tabClass?: string;
  panelClass?: string;
  tabDatasetKey?: string;
  panelDatasetKey?: string;
}

export interface DialogTabPair {
  button: HTMLButtonElement;
  panel: HTMLDivElement;
}

export interface DialogOptionButtonOptions {
  label: string;
  baseClass: string;
  datasetKey: string;
  value: string;
  selected?: boolean;
  extraClass?: string;
}

export interface DialogToggleButtonOptions {
  label: string;
  baseClass: string;
  pressed?: boolean;
  title?: string;
  datasetKey?: string;
  value?: string;
  extraClass?: string;
}

export function focusAndSelectInput(input: HTMLInputElement | HTMLTextAreaElement): void {
  input.focus({ preventScroll: true });
  input.select();
}

export function showDialogError(errorRow: HTMLElement, message: string): void {
  errorRow.textContent = message;
  errorRow.hidden = false;
}

export function clearDialogError(errorRow: HTMLElement): void {
  errorRow.hidden = true;
  errorRow.textContent = '';
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
      // Re-snapshot host theme tokens so paper↔ink swaps applied since the
      // shell was constructed are reflected on this open.
      inheritHostTokens(host, overlay);
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
      overlay.dispatchEvent(new CustomEvent('fc-range-picker-stop-all'));
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

export function appendDialogFrame(shell: DialogShell, opts: DialogFrameOptions): DialogFrame {
  if (opts.panelClasses) shell.panel.classList.add(...opts.panelClasses);

  const header = document.createElement('div');
  header.className = opts.headerClass ?? 'fc-fmtdlg__header';
  header.textContent = opts.title;
  shell.panel.appendChild(header);

  const body = document.createElement(opts.bodyTag ?? 'div');
  body.className = opts.bodyClass ?? 'fc-fmtdlg__body';
  shell.panel.appendChild(body);

  const footer = document.createElement('div');
  footer.className = opts.footerClass ?? 'fc-fmtdlg__footer';
  shell.panel.appendChild(footer);

  return { header, body, footer };
}

export function createDialogButton(
  opts: DialogButtonOptions,
): HTMLButtonElement {
  const button = document.createElement('button');
  button.type = 'button';
  const classes = [opts.baseClass ?? 'fc-fmtdlg__btn'];
  if (opts.variant === 'primary') {
    classes.push(opts.primaryClass ?? `${opts.baseClass ?? 'fc-fmtdlg__btn'}--primary`);
  } else if (opts.variant === 'secondary' && opts.secondaryClass) {
    classes.push(opts.secondaryClass);
  }
  button.className = classes.join(' ');
  button.textContent = opts.label;
  return button;
}

export function appendDialogButton(
  footer: HTMLElement,
  opts: DialogButtonOptions,
): HTMLButtonElement {
  const button = createDialogButton(opts);
  footer.appendChild(button);
  return button;
}

export function appendDialogIconButton(
  parent: HTMLElement,
  opts: DialogIconButtonOptions,
): HTMLButtonElement {
  const button = document.createElement('button');
  button.type = 'button';
  button.className = opts.baseClass;
  button.setAttribute('aria-label', opts.ariaLabel);
  if (opts.title) button.title = opts.title;
  if (opts.html) button.innerHTML = opts.html;
  else button.textContent = opts.label;
  parent.appendChild(button);
  return button;
}

export function appendDialogTabPair(
  tabsStrip: HTMLElement,
  body: HTMLElement,
  opts: DialogTabPairOptions,
): DialogTabPair {
  const button = document.createElement('button');
  button.type = 'button';
  button.className = opts.tabClass ?? 'fc-fmtdlg__tab';
  button.textContent = opts.label;
  button.setAttribute('role', 'tab');
  button.setAttribute('aria-selected', 'false');
  button.setAttribute('aria-controls', opts.panelId);
  button.tabIndex = -1;
  button.id = opts.tabId;
  if (opts.tabDatasetKey) button.dataset[opts.tabDatasetKey] = opts.id;
  tabsStrip.appendChild(button);

  const panel = document.createElement('div');
  panel.className = opts.panelClass ?? 'fc-fmtdlg__panel-tab';
  panel.id = opts.panelId;
  panel.setAttribute('role', 'tabpanel');
  panel.setAttribute('aria-labelledby', opts.tabId);
  panel.hidden = true;
  if (opts.panelDatasetKey) panel.dataset[opts.panelDatasetKey] = opts.id;
  body.appendChild(panel);

  return { button, panel };
}

export function appendDialogOptionButton(
  parent: HTMLElement,
  opts: DialogOptionButtonOptions,
): HTMLButtonElement {
  const button = document.createElement('button');
  button.type = 'button';
  button.className = opts.extraClass ? `${opts.baseClass} ${opts.extraClass}` : opts.baseClass;
  button.textContent = opts.label;
  button.setAttribute('role', 'option');
  button.setAttribute('aria-selected', opts.selected ? 'true' : 'false');
  button.tabIndex = -1;
  button.dataset[opts.datasetKey] = opts.value;
  parent.appendChild(button);
  return button;
}

export function createDialogToggleButton(opts: DialogToggleButtonOptions): HTMLButtonElement {
  const button = document.createElement('button');
  button.type = 'button';
  button.className = opts.extraClass ? `${opts.baseClass} ${opts.extraClass}` : opts.baseClass;
  button.setAttribute('aria-label', opts.label);
  button.setAttribute('aria-pressed', opts.pressed ? 'true' : 'false');
  if (opts.title) button.title = opts.title;
  if (opts.datasetKey && opts.value !== undefined) button.dataset[opts.datasetKey] = opts.value;
  return button;
}

export function appendDialogActions(
  footer: HTMLElement,
  opts: {
    cancelLabel: string;
    okLabel: string;
    buttonBaseClass?: string;
    buttonPrimaryClass?: string;
    buttonSecondaryClass?: string;
  },
): { cancelBtn: HTMLButtonElement; okBtn: HTMLButtonElement } {
  const cancelBtn = appendDialogButton(footer, {
    label: opts.cancelLabel,
    variant: 'secondary',
    baseClass: opts.buttonBaseClass,
    secondaryClass: opts.buttonSecondaryClass,
  });
  const okBtn = appendDialogButton(footer, {
    label: opts.okLabel,
    variant: 'primary',
    baseClass: opts.buttonBaseClass,
    primaryClass: opts.buttonPrimaryClass,
  });
  return { cancelBtn, okBtn };
}
