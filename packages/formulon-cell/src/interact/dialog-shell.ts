import { visibleFocusReturn } from '../toolbar/ribbon/fit-ribbon.js';
import { enhanceCustomSelect, syncCustomSelects } from './custom-select.js';
import { advanceDialogTab, focusableDialogItems, focusSheetAfterDialog } from './dialog-focus.js';
import { clearDialogPlacement, currentDialogAnchor, placeDialogPanel } from './dialog-placement.js';
import { overlayPortalFor } from './overlay-portal.js';

export interface DialogShellDeps {
  /** Any element inside the target `.fc-host`. The overlay is mounted into that
   *  host's themed overlay portal, so it escapes any `contain: strict` parents
   *  yet still inherits theme tokens through cascade. */
  host: HTMLElement;
  /** Root class on the overlay (e.g. `fc-iterdlg`). */
  className: string;
  /** Accessible name; assigned to `aria-label` and refreshable via
   *  `setAriaLabel()` so locale swaps don't need a remount. */
  ariaLabel: string;
  /** Invoked on Escape (while open); a backdrop click never dismisses. If omitted both
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
  /** Trigger for a popup rather than an on/off control: reports `aria-expanded`
   *  and `aria-haspopup` instead of `aria-pressed`. */
  popup?: boolean;
}

export function focusAndSelectInput(input: HTMLInputElement | HTMLTextAreaElement): void {
  input.focus({ preventScroll: true });
  input.select();
}

/**
 * Run a dialog's initial focus on the next frame, unless focus has moved since the call: the
 * user is already typing in the dialog, or has left it, and must not be pulled back.
 */
export function deferInitialFocus(container: HTMLElement, focus: () => void): void {
  const doc = container.ownerDocument;
  // Focus moving within, into or out of the container passes through it.
  let moved = false;
  const onMove = (): void => {
    moved = true;
  };
  container.addEventListener('focusin', onMove);
  container.addEventListener('focusout', onMove);
  const frame =
    doc.defaultView?.requestAnimationFrame.bind(doc.defaultView) ?? requestAnimationFrame;
  frame(() => {
    container.removeEventListener('focusin', onMove);
    container.removeEventListener('focusout', onMove);
    if (moved || container.hidden || !container.isConnected) return;
    focus();
  });
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

// A document can host dialogs from multiple spreadsheet instances. Keep one
// activation order per document so every modal key handler agrees on which
// visible panel currently owns the keyboard.
const openDialogLayers = new WeakMap<Document, HTMLElement[]>();
const handledDialogKeyEvents = new WeakSet<KeyboardEvent>();

function removeDialogLayer(overlay: HTMLElement): void {
  const stack = openDialogLayers.get(overlay.ownerDocument);
  if (!stack) return;
  const index = stack.indexOf(overlay);
  if (index !== -1) stack.splice(index, 1);
  if (stack.length === 0) openDialogLayers.delete(overlay.ownerDocument);
}

function activateDialogLayer(overlay: HTMLElement): void {
  removeDialogLayer(overlay);
  const doc = overlay.ownerDocument;
  const stack = openDialogLayers.get(doc) ?? [];
  stack.push(overlay);
  openDialogLayers.set(doc, stack);
}

function ownsDialogKeys(overlay: HTMLElement): boolean {
  const stack = openDialogLayers.get(overlay.ownerDocument);
  if (!stack) return false;
  // Drop closed or detached layers in case a host was removed without dispose.
  for (let i = stack.length - 1; i >= 0; i -= 1) {
    const layer = stack[i];
    if (!layer || layer.hidden || !layer.isConnected) stack.splice(i, 1);
  }
  if (stack.length === 0) openDialogLayers.delete(overlay.ownerDocument);
  return stack.at(-1) === overlay;
}

const keyWasHandled = (event: KeyboardEvent): boolean =>
  event.defaultPrevented || handledDialogKeyEvents.has(event);

/**
 * Shared lifecycle and listener bookkeeping for the spreadsheet's modal dialogs:
 *
 *  1. Mounts a hidden `<div role=dialog aria-modal>` with the requested class into the host's
 *     themed overlay portal, so callers control reveal timing and theme tokens still apply
 *     outside the `.fc-host` `contain: strict` boundary.
 *  2. Keeps Tab / Shift+Tab focus inside the visible dialog, as desktop modals do.
 *  3. Hooks Escape (keydown on the dialog and on the document, while shown) to `onDismiss`, so a
 *     dialog that stops its keys from bubbling still closes; a backdrop click is swallowed and
 *     never dismisses.
 *  4. `on()` registers listeners that `dispose()` removes in one sweep, together with the
 *     built-in Escape and backdrop listeners.
 */
/** A ribbon command hands its result to the sheet, as the desktop ribbon keeps no focus. */
const RIBBON_CONTROL = '.fc-tb, [data-ribbon-command], [data-ribbon-tab]';

/**
 * Hands focus back after `surface` closes: to `opener`, the element focused when it opened, while
 * that element is still attached, outside the surface, takes focus and is not a ribbon control;
 * otherwise to the sheet `host`. Focus never falls to the page.
 */
export function returnFocus(
  opener: HTMLElement | null,
  surface: HTMLElement,
  host: HTMLElement,
): void {
  const target =
    opener?.isConnected && !surface.contains(opener) && !opener.closest(RIBBON_CONTROL)
      ? visibleFocusReturn(opener)
      : null;
  target?.focus({ preventScroll: true });
  if (target && host.ownerDocument.activeElement === target) return;
  focusSheetAfterDialog(host);
}

/**
 * How Escape treats a dialog surface: a modal dialog or a floating panel over the grid closes, a
 * pane docked beside or below the grid stays open and hands focus to the sheet.
 */
export type SurfaceKind = 'modal' | 'floating' | 'pane';

const MODAL_SURFACE = '[data-fc-surface="modal"]';

/**
 * The control to return focus to when a dialog closes: the `anchor` the open call was given, which
 * a click does not focus in every browser, else the element focused in `doc`; null when only the
 * body is.
 */
export function openerOf(doc: Document, anchor?: HTMLElement | null): HTMLElement | null {
  if (anchor) return anchor;
  const active = doc.activeElement;
  const OwnerHTMLElement = doc.defaultView?.HTMLElement ?? HTMLElement;
  return active instanceof OwnerHTMLElement && active !== doc.body ? active : null;
}

/**
 * Whether `event` was pressed in a modal dialog, which owns its keys: document-level shortcuts
 * (the ribbon's, the backstage's, the title bar's) leave it to the dialog. Also true while the
 * dialog the key just closed is still the target's ancestor.
 */
export function modalDialogOwnsKeys(event: Event): boolean {
  const target = event.target as { closest?: (selector: string) => Element | null } | null;
  return typeof target?.closest === 'function' && target.closest(MODAL_SURFACE) !== null;
}

/** Records `kind` on the surface element as `data-fc-surface`. */
export function markSurface(surface: HTMLElement, kind: SurfaceKind): void {
  surface.dataset.fcSurface = kind;
}

/**
 * Escape in a surface marked as a docked pane leaves it open and hands focus to the sheet `host`;
 * the pane closes from its Close button. A key a control inside the pane already handled is left
 * alone. Returns the detach function.
 */
export function bindPaneEscape(pane: HTMLElement, host: HTMLElement): () => void {
  const onKeyDown = (event: KeyboardEvent): void => {
    if (event.key !== 'Escape' || event.defaultPrevented) return;
    if (pane.dataset.fcSurface !== 'pane') return;
    event.preventDefault();
    focusSheetAfterDialog(host);
  };
  pane.addEventListener('keydown', onKeyDown);
  return () => pane.removeEventListener('keydown', onKeyDown);
}

export function createDialogShell(deps: DialogShellDeps): DialogShell {
  const { host, className, onDismiss } = deps;
  const document = host.ownerDocument;
  const OwnerHTMLElement = document.defaultView?.HTMLElement ?? HTMLElement;
  const OwnerHTMLSelectElement = document.defaultView?.HTMLSelectElement ?? HTMLSelectElement;
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
  markSurface(overlay, 'modal');
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
        if (!(node instanceof OwnerHTMLElement)) continue;
        if (node instanceof OwnerHTMLSelectElement) {
          const handle = enhanceCustomSelect(node);
          if (handle) selectHandles.push(handle);
        }
        enhanceSelects(node);
      }
    }
  });
  selectObserver.observe(panel, { childList: true, subtree: true });

  overlayPortalFor(host).appendChild(overlay);

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

  // A modal dialog ignores backdrop clicks and keeps them away from the grid.
  on(overlay, 'click', (e) => {
    if (e.target === overlay) e.stopPropagation();
  });
  if (onDismiss) {
    const onEscapeKeydown: EventListener = (e) => {
      const event = e as KeyboardEvent;
      if (overlay.hidden || keyWasHandled(event) || !ownsDialogKeys(overlay)) return;
      if (event.key === 'Escape') {
        handledDialogKeyEvents.add(event);
        event.preventDefault();
        onDismiss();
      }
    };
    on(overlay, 'keydown', onEscapeKeydown);
    on(document, 'keydown', onEscapeKeydown);
  }

  const onTabKeydown: EventListener = (e) => {
    const event = e as KeyboardEvent;
    if (overlay.hidden || keyWasHandled(event) || event.key !== 'Tab' || !ownsDialogKeys(overlay))
      return;
    if (advanceDialogTab(panel, event)) handledDialogKeyEvents.add(event);
  };
  on(overlay, 'keydown', onTabKeydown);
  on(document, 'keydown', onTabKeydown);

  const closeShell = (): void => {
    removeDialogLayer(overlay);
    const shouldRestore =
      !overlay.hidden &&
      (overlay.contains(document.activeElement) || document.activeElement === document.body);
    overlay.dispatchEvent(new CustomEvent('fc-range-picker-stop-all'));
    overlay.hidden = true;
    clearDialogPlacement(overlay, panel);
    if (shouldRestore) returnFocus(restoreFocusEl, overlay, host);
    restoreFocusEl = null;
  };

  return {
    overlay,
    panel,
    on,
    open() {
      if (disposed) return;
      enhanceSelects(panel);
      syncCustomSelects(panel);
      activateDialogLayer(overlay);
      if (!overlay.contains(document.activeElement)) restoreFocusEl = openerOf(document);
      overlay.hidden = false;
      placeDialogPanel(overlay, panel, currentDialogAnchor());
      if (!overlay.contains(document.activeElement)) {
        const first = focusableDialogItems(panel)[0] ?? panel;
        first.focus({ preventScroll: true });
      }
    },
    close: closeShell,
    setAriaLabel(label: string) {
      ariaLabel = label;
      overlay.setAttribute('aria-label', label);
    },
    dispose() {
      if (disposed) return;
      // Disposing an open dialog closes it first, so focus goes back where it came from.
      if (!overlay.hidden) closeShell();
      disposed = true;
      for (const l of listeners) {
        l.target.removeEventListener(l.event, l.handler, l.options);
      }
      listeners.length = 0;
      selectObserver.disconnect();
      for (const handle of selectHandles.splice(0)) handle.dispose();
      removeDialogLayer(overlay);
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

export function createDialogButton(opts: DialogButtonOptions): HTMLButtonElement {
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
  if (opts.popup) {
    button.setAttribute('aria-haspopup', 'true');
    button.setAttribute('aria-expanded', opts.pressed ? 'true' : 'false');
  } else {
    button.setAttribute('aria-pressed', opts.pressed ? 'true' : 'false');
  }
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
