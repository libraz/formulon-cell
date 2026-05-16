// Body-attached overlays (context menu, autocomplete, hover tip, etc.) live
// outside the `.fc-host[data-fc-theme]` scope, so they don't inherit theme
// tokens. This helper snapshots the host's computed values onto the overlay
// root so it picks up paper/ink tokens regardless of where it's mounted.
const FORWARDED_TOKENS: readonly string[] = [
  '--fc-bg',
  '--fc-bg-elev',
  '--fc-bg-rail',
  '--fc-fg',
  '--fc-fg-strong',
  '--fc-fg-mute',
  '--fc-fg-faint',
  '--fc-rule',
  '--fc-rule-strong',
  '--fc-accent',
  '--fc-accent-strong',
  '--fc-accent-soft',
  '--fc-accent-fg',
  '--fc-radius-sm',
  '--fc-radius-md',
  '--fc-hairline',
  '--fc-font-ui',
  '--fc-font-mono',
  'color-scheme',
  // Format dialog token surface (paper / ink override these)
  '--fc-fmtdlg-panel-bg',
  '--fc-fmtdlg-panel-fg',
  '--fc-fmtdlg-header-bg',
  '--fc-fmtdlg-body-bg',
  '--fc-fmtdlg-footer-bg',
  '--fc-fmtdlg-rule',
  '--fc-fmtdlg-label-fg',
  '--fc-fmtdlg-tab-bg',
  '--fc-fmtdlg-tab-border',
  '--fc-fmtdlg-tab-color',
  '--fc-fmtdlg-tab-hover',
  '--fc-fmtdlg-tab-hover-color',
  '--fc-fmtdlg-tab-active-bg',
  '--fc-fmtdlg-tab-active-color',
  '--fc-fmtdlg-tab-active-underline',
  '--fc-fmtdlg-tab-radius',
  '--fc-fmtdlg-preview-bg',
  '--fc-fmtdlg-preview-fg',
  '--fc-fmtdlg-preview-border',
  '--fc-fmtdlg-list-bg',
  '--fc-fmtdlg-list-fg',
  '--fc-fmtdlg-list-border',
  '--fc-fmtdlg-list-focus-border',
  '--fc-fmtdlg-list-hover-bg',
  '--fc-fmtdlg-list-selected-bg',
  '--fc-fmtdlg-list-selected-fg',
  '--fc-fmtdlg-list-selected-bar',
  '--fc-fmtdlg-cat-selected-bg',
  '--fc-fmtdlg-cat-selected-fg',
  '--fc-fmtdlg-input-bg',
  '--fc-fmtdlg-input-fg',
  '--fc-fmtdlg-input-border',
  '--fc-fmtdlg-input-hover-border',
  '--fc-fmtdlg-btn-bg',
  '--fc-fmtdlg-btn-fg',
  '--fc-fmtdlg-btn-border',
  '--fc-fmtdlg-btn-hover-bg',
  '--fc-fmtdlg-btn-hover-border',
  '--fc-fmtdlg-btn-primary-bg',
  '--fc-fmtdlg-btn-primary-fg',
  '--fc-fmtdlg-btn-primary-hover-bg',
  '--fc-fmtdlg-hint-fg',
  '--fc-fmtdlg-section-fg',
  '--fc-fmtdlg-chip-bg',
  '--fc-fmtdlg-chip-checked-bg',
  '--fc-fmtdlg-chip-checked-border',
  '--fc-fmtdlg-dial-border',
  '--fc-fmtdlg-dial-bg',
  '--fc-fmtdlg-dial-arc',
  '--fc-fmtdlg-dial-dot',
  '--fc-fmtdlg-dial-dot-hover',
  '--fc-fmtdlg-dial-dot-active',
  '--fc-fmtdlg-dial-pointer',
  '--fc-fmtdlg-dial-text',
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
