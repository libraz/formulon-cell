// A single light-DOM portal container per mounted spreadsheet. Floating UI
// (context menus, dialogs, tooltips, dropdowns) attaches here instead of
// directly to `document.body`.
//
// Why a dedicated container rather than raw `document.body`: overlays must
// escape the `.fc-host { contain: strict }` boundary, but appending straight
// to `<body>` drops them out of the theme cascade. The old fix snapshotted a
// hand-maintained list of `--fc-*` custom properties onto every overlay
// (`inherit-host-tokens`), which broke silently whenever a new token was added.
//
// The portal carries the same `data-fc-theme` as its `.fc-host`, so every
// overlay inside it inherits paper/ink/contrast tokens through normal cascade
// — no per-token forwarding, and new tokens flow automatically. It stays in the
// light DOM (not a shadow root) so `document.querySelector` still reaches
// overlays, keeping the imperative dialog/menu code and tests unchanged.

const portals = new WeakMap<Element, HTMLElement>();

function syncTheme(host: HTMLElement, portal: HTMLElement): void {
  const theme = host.dataset.fcTheme;
  if (theme) portal.dataset.fcTheme = theme;
  else delete portal.dataset.fcTheme;
}

/** Create (or return the existing) overlay portal for a `.fc-host`. Called by
 *  `mount()`. The portal is appended to `<body>` and theme-synced immediately. */
export function ensureOverlayPortal(host: HTMLElement): HTMLElement {
  const existing = portals.get(host);
  if (existing?.isConnected) {
    syncTheme(host, existing);
    return existing;
  }
  const portal = document.createElement('div');
  portal.className = 'fc-overlay-portal';
  syncTheme(host, portal);
  document.body.appendChild(portal);
  portals.set(host, portal);
  return portal;
}

/** Resolve the overlay portal a trigger element belongs to. Sub-overlays opened
 *  from inside another overlay (e.g. a submenu off a dialog) already live in the
 *  portal — which is a `<body>` sibling of `.fc-host`, not a descendant — so
 *  check for an enclosing portal first, then fall back to walking up to the
 *  owning `.fc-host`. Returns `document.body` only when the node is detached or
 *  its host has no portal yet (defensive — `mount()` always creates one). */
export function overlayPortalFor(node: Element | null | undefined): HTMLElement {
  const nested = node?.closest<HTMLElement>('.fc-overlay-portal');
  if (nested) return nested;
  const host = node?.closest<HTMLElement>('.fc-host') ?? null;
  if (!host) return document.body;
  return portals.get(host) ?? ensureOverlayPortal(host);
}

/** Push the current `.fc-host` theme onto its portal. Called from the
 *  instance's `setTheme` so paper↔ink swaps reach open/future overlays. */
export function syncOverlayPortalTheme(host: HTMLElement): void {
  const portal = portals.get(host);
  if (portal) syncTheme(host, portal);
}

/** Remove a host's portal and its remaining overlay children. Called on
 *  `dispose()`. Idempotent. */
export function disposeOverlayPortal(host: Element): void {
  const portal = portals.get(host);
  if (portal) {
    portal.remove();
    portals.delete(host);
  }
}
