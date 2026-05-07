// Body-attached overlays (context menu, autocomplete, hover tip, etc.) live
// outside the `.fc-host[data-fc-theme]` scope, so they don't inherit theme
// tokens. This helper snapshots the host's computed values onto the overlay
// root so it picks up paper/ink tokens regardless of where it's mounted.
const FORWARDED_TOKENS: readonly string[] = [
  '--fc-bg',
  '--fc-bg-elev',
  '--fc-bg-rail',
  '--fc-fg',
  '--fc-fg-mute',
  '--fc-fg-faint',
  '--fc-rule',
  '--fc-rule-strong',
  '--fc-accent',
  '--fc-accent-soft',
  '--fc-accent-fg',
  '--fc-radius-sm',
  '--fc-radius-md',
  '--fc-hairline',
  '--fc-font-ui',
  '--fc-font-mono',
  'color-scheme',
];

/** Copy theme-related custom properties from the host onto the overlay root.
 *  Call before showing — re-reads the host on each call so paper↔ink swaps
 *  take effect without rebinding listeners. */
export function inheritHostTokens(host: Element, target: HTMLElement): void {
  const cs = getComputedStyle(host);
  for (const tok of FORWARDED_TOKENS) {
    const v = cs.getPropertyValue(tok);
    if (v) target.style.setProperty(tok, v.trim());
  }
}
