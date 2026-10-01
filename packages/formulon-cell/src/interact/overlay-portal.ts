// A single light-DOM portal container per mounted spreadsheet. Floating UI
// (context menus, dialogs, tooltips, dropdowns) attaches here instead of
// directly to `document.body`.
//
// Why a dedicated container rather than raw `document.body`: overlays must
// escape the `.fc-host { contain: strict }` boundary, but appending straight
// to `<body>` drops them out of the theme cascade. The portal carries the
// same `data-fc-theme` as its `.fc-host`, so every overlay inside it inherits
// paper/ink/contrast tokens through normal cascade. It stays in the light DOM
// (not a shadow root) so imperative dialog/menu code and tests can still find
// overlays with `document.querySelector`.

export interface OverlayOptions {
  /**
   * Element that owns the overlay boundary. A resolver is evaluated whenever
   * the portal is ensured or refreshed, which lets a host follow a dialog or
   * fullscreen container that changes during the instance lifetime.
   */
  root?: HTMLElement | (() => HTMLElement);
}

interface PortalRecord {
  document: Document;
  portal: HTMLElement;
  onFullscreenChange: () => void;
}

const portals = new WeakMap<Element, PortalRecord>();
const portalHosts = new WeakMap<HTMLElement, HTMLElement>();
const overlayOptions = new WeakMap<Element, OverlayOptions | undefined>();
const overlayOwners = new WeakMap<Element, () => HTMLElement | null>();

/** Associate an external toolbar with its sheet without guessing a global host. */
export function registerOverlayOwner(node: Element, getHost: () => HTMLElement | null): () => void {
  overlayOwners.set(node, getHost);
  return () => {
    if (overlayOwners.get(node) === getHost) overlayOwners.delete(node);
  };
}

function syncTheme(host: HTMLElement, portal: HTMLElement): void {
  const theme = host.dataset.fcTheme;
  if (theme) portal.dataset.fcTheme = theme;
  else delete portal.dataset.fcTheme;
}

/** Stamp the portal with the host's current instance id. Remounting the same
 * host hands the portal to the newest instance — see `disposeOverlayPortal`. */
function claimPortal(host: HTMLElement, portal: HTMLElement): void {
  const instanceId = host.dataset.fcInstId;
  if (instanceId) portal.dataset.fcInstId = instanceId;
  else delete portal.dataset.fcInstId;
}

function hostDocument(host: Element): Document {
  const ownerDocument = host.ownerDocument;
  if (!ownerDocument) {
    throw new Error('overlay portal host must belong to a document');
  }
  return ownerDocument;
}

function isHTMLElementInDocument(value: unknown, ownerDocument: Document): value is HTMLElement {
  if (!value || typeof value !== 'object') return false;
  const element = value as HTMLElement;
  if (element.ownerDocument !== ownerDocument || element.nodeType !== 1) return false;

  const HTMLElementConstructor = ownerDocument.defaultView?.HTMLElement;
  if (HTMLElementConstructor) return element instanceof HTMLElementConstructor;
  return typeof HTMLElement !== 'undefined' && element instanceof HTMLElement;
}

function invalidRootError(): Error {
  return new Error(
    "overlay portal root must be an HTMLElement in the mounted host's ownerDocument",
  );
}

function resolveExplicitRoot(
  configured: OverlayOptions,
  ownerDocument: Document,
): HTMLElement | undefined {
  if (configured.root === undefined) return undefined;

  let root: unknown;
  try {
    root = typeof configured.root === 'function' ? configured.root() : configured.root;
  } catch (error) {
    throw new Error('overlay portal root resolver failed', { cause: error });
  }
  if (!isHTMLElementInDocument(root, ownerDocument)) throw invalidRootError();
  return root;
}

function containingOpenDialog(host: HTMLElement): HTMLElement | null {
  let current: Element | null = host;
  while (current) {
    if (
      current.tagName.toLowerCase() === 'dialog' &&
      ((current as HTMLDialogElement).open || current.hasAttribute('open'))
    ) {
      return current as HTMLElement;
    }
    current = current.parentElement;
  }
  return null;
}

function defaultOverlayRoot(host: HTMLElement, ownerDocument: Document): HTMLElement {
  // A native dialog is a focus boundary. Prefer the nearest containing open
  // dialog so overlays remain inside the boundary even when the document also
  // has a fullscreen element.
  const dialog = containingOpenDialog(host);
  if (dialog && isHTMLElementInDocument(dialog, ownerDocument)) return dialog;

  // Fullscreen content has its own top layer. Only use it for a host contained
  // by that element; an unrelated spreadsheet on the same document must keep
  // its ordinary body portal.
  const fullscreen = ownerDocument.fullscreenElement;
  if (
    fullscreen &&
    isHTMLElementInDocument(fullscreen, ownerDocument) &&
    (fullscreen === host || fullscreen.contains(host))
  ) {
    return fullscreen;
  }

  const body = ownerDocument.body;
  if (body && isHTMLElementInDocument(body, ownerDocument)) return body;
  throw new Error('overlay portal host document has no valid body');
}

function resolveRoot(host: HTMLElement, options: OverlayOptions | undefined): HTMLElement {
  const ownerDocument = hostDocument(host);
  const explicit = options ? resolveExplicitRoot(options, ownerDocument) : undefined;
  return explicit ?? defaultOverlayRoot(host, ownerDocument);
}

function attachFullscreenListener(record: PortalRecord): void {
  record.document.addEventListener('fullscreenchange', record.onFullscreenChange);
}

function detachFullscreenListener(record: PortalRecord): void {
  record.document.removeEventListener('fullscreenchange', record.onFullscreenChange);
}

function createPortal(host: HTMLElement, options: OverlayOptions | undefined): HTMLElement {
  const ownerDocument = hostDocument(host);
  // Resolve before mutating the DOM or installing the listener. A failed
  // resolver therefore cannot leave a half-initialized portal behind.
  const root = resolveRoot(host, options);
  const portal = ownerDocument.createElement('div');
  portal.className = 'fc-overlay-portal';
  claimPortal(host, portal);
  syncTheme(host, portal);

  const record: PortalRecord = {
    document: ownerDocument,
    portal,
    onFullscreenChange: () => {
      // A resolver may become temporarily invalid while a host is being
      // moved between presentation containers. Keep the current portal in
      // place until the host can be resolved again; direct refresh calls still
      // reject the invalid root synchronously.
      try {
        refreshOverlayPortal(host);
      } catch {
        // The portal remains usable at its last valid root.
      }
    },
  };

  let listenerAttached = false;
  try {
    root.appendChild(portal);
    attachFullscreenListener(record);
    listenerAttached = true;
    portals.set(host, record);
    portalHosts.set(portal, host);
    return portal;
  } catch (error) {
    if (listenerAttached) detachFullscreenListener(record);
    portals.delete(host);
    portalHosts.delete(portal);
    portal.remove();
    throw new Error('failed to initialize overlay portal', { cause: error });
  }
}

function configuredOptions(host: Element): OverlayOptions | undefined {
  return overlayOptions.get(host);
}

/**
 * Move an existing portal to the current root. The portal node is reused so
 * open menus, dialogs, and cached handles continue to point at the same DOM
 * subtree while presentation containers change.
 */
export function refreshOverlayPortal(host: HTMLElement): void {
  const record = portals.get(host);
  if (!record) {
    ensureOverlayPortal(host);
    return;
  }

  const root = resolveRoot(host, configuredOptions(host));
  claimPortal(host, record.portal);
  syncTheme(host, record.portal);
  if (record.portal.parentElement !== root) root.appendChild(record.portal);
}

/**
 * Create (or return the existing) overlay portal for a `.fc-host`. Called by
 * `mount()`. With no configured root, ordinary page embeds use the owning
 * document's body; hosts inside an open dialog or fullscreen element follow
 * that presentation boundary.
 */
export function ensureOverlayPortal(host: HTMLElement, options?: OverlayOptions): HTMLElement {
  const hadOptions = overlayOptions.has(host);
  const previousOptions = overlayOptions.get(host);
  if (options !== undefined) overlayOptions.set(host, { root: options.root });

  try {
    const existing = portals.get(host);
    if (existing) {
      refreshOverlayPortal(host);
      return existing.portal;
    }
    return createPortal(host, configuredOptions(host));
  } catch (error) {
    if (options !== undefined) {
      if (hadOptions) overlayOptions.set(host, previousOptions);
      else overlayOptions.delete(host);
    }
    throw error;
  }
}

/** Update a host's overlay boundary and immediately reparent an existing
 * portal. Invalid roots leave the previous configuration and DOM untouched. */
export function setOverlayOptions(host: HTMLElement, next: OverlayOptions | undefined): void {
  const hadOptions = overlayOptions.has(host);
  const previousOptions = overlayOptions.get(host);
  if (next === undefined) overlayOptions.delete(host);
  else overlayOptions.set(host, { root: next.root });

  try {
    // Validate a resolver even before the first portal is created. This keeps
    // a bad host configuration from being silently replaced by body later.
    if (!portals.has(host) && next?.root !== undefined) resolveRoot(host, next);
    if (portals.has(host)) refreshOverlayPortal(host);
  } catch (error) {
    if (hadOptions) overlayOptions.set(host, previousOptions);
    else overlayOptions.delete(host);
    throw error;
  }
}

/** Resolve the overlay portal a trigger element belongs to. Sub-overlays
 * opened from inside another overlay already live in the portal, but refresh
 * its owning host first so a fullscreen/dialog transition moves the existing
 * subtree before the caller appends another child. */
export function overlayPortalFor(node: Element | null | undefined): HTMLElement {
  const nested = node?.closest<HTMLElement>('.fc-overlay-portal');
  if (nested) {
    const owner = portalHosts.get(nested);
    if (owner) {
      refreshOverlayPortal(owner);
      return portals.get(owner)?.portal ?? nested;
    }
    return nested;
  }

  let ownedHost: HTMLElement | null = null;
  for (let ancestor = node; ancestor; ancestor = ancestor.parentElement) {
    const resolveOwner = overlayOwners.get(ancestor);
    if (!resolveOwner) continue;
    ownedHost = resolveOwner();
    if (ownedHost) break;
  }
  const host = ownedHost ?? node?.closest<HTMLElement>('.fc-host') ?? null;
  if (host) {
    refreshOverlayPortal(host);
    return portals.get(host)?.portal ?? ensureOverlayPortal(host);
  }

  const ownerDocument = node?.ownerDocument ?? (typeof document !== 'undefined' ? document : null);
  if (!ownerDocument?.body) {
    throw new Error('overlay portal requires a document body for detached nodes');
  }
  return ownerDocument.body;
}

/** Push the current `.fc-host` theme onto its portal. Called from the
 * instance's `setTheme` so paper↔ink swaps reach open/future overlays. */
export function syncOverlayPortalTheme(host: HTMLElement): void {
  const portal = portals.get(host)?.portal;
  if (portal) syncTheme(host, portal);
}

/** Remove a host's portal and its remaining overlay children. Called on
 * `dispose()`. Idempotent. The host-supplied root is never removed.
 *
 * Pass the disposing instance's id to keep a superseded instance from tearing
 * the portal out from under a live one: a remount of the same host (an async
 * `mount()` racing React StrictMode's double-effect) reuses the portal, and
 * the first instance disposes only after the second has claimed it. */
export function disposeOverlayPortal(host: Element, instanceId?: string): void {
  const record = portals.get(host);
  if (!record) return;
  if (instanceId !== undefined && record.portal.dataset.fcInstId !== instanceId) return;

  detachFullscreenListener(record);
  record.portal.remove();
  portalHosts.delete(record.portal);
  portals.delete(host);
  overlayOptions.delete(host);
}
