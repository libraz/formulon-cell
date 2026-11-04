/**
 * Materialise CSS custom properties into JS-side numbers so the Canvas
 * painters don't have to keep parsing CSS at every frame. Re-resolve when
 * the theme changes or the host is mounted.
 */
export interface ResolvedTheme {
  bg: string;
  bgRail: string;
  bgElev: string;
  bgHeader: string;
  fg: string;
  fgMute: string;
  fgFaint: string;
  fgStrong: string;
  /** Automatic colour of cell text and borders. */
  cellAutoFg: string;
  rule: string;
  ruleStrong: string;
  /** Freeze-pane divider. */
  freezeRule: string;
  accent: string;
  accentFg: string;
  accentSoft: string;
  cellErrorFg: string;
  cellFormulaFg: string;
  cellBoolFg: string;
  cellNumFg: string;
  hoverStripe: string;
  /** Row and column header fill. */
  headerBg: string;
  /** Unselected header strip fill at its outer edge and at the grid edge;
   *  both default to `headerBg` (flat). */
  headerStripOuter: string;
  headerStripInner: string;
  /** Rule between two header cells, at the grid edge and at the strip's
   *  outer edge; the outer end defaults to `headerRule` (flat). */
  headerRule: string;
  headerRuleOuter: string;
  headerFg: string;
  headerFgActive: string;
  /** Row-number colour of the data rows a filter lets through. */
  headerFgFiltered: string;
  /** Bar under a selected row/column header; `transparent` draws none. */
  headerActiveBar: string;
  /** Font weight of a selected header label. */
  headerFgActiveWeight: number;
  /** Outline gutter fill; defaults to `headerBg`. */
  outlineGutterBg: string;
  /** Outline level buttons: face shaded top to bottom (defaults flat
   *  `headerBg`), border, digit, and a right/bottom shadow (`transparent`
   *  draws none). */
  outlineButtonFaceTop: string;
  outlineButtonFaceBottom: string;
  outlineButtonBorder: string;
  outlineButtonFg: string;
  outlineButtonShadow: string;
  /** Row under a level button's top border; `transparent` draws none. */
  outlineButtonHighlight: string;
  /** Level button top and left border, ramped start to end along each edge;
   *  both default to `outlineButtonBorder` (flat). */
  outlineButtonEdgeStart: string;
  outlineButtonEdgeEnd: string;
  /** Outline brackets and detail dots; defaults to `rule`. */
  outlineRule: string;
  /** Outline +/- toggle border and glyph; default `ruleStrong` and `fg`. */
  outlineToggleBorder: string;
  outlineToggleFg: string;

  /** Page Layout / Page Break Preview chrome. */
  pageBackdrop: string;
  pagePaper: string;
  pageEdge: string;
  pageBandFg: string;
  pageBreakAuto: string;
  pageBreakManual: string;
  pageNumberFg: string;
  pageOutside: string;

  fontUi: string;
  fontMono: string;

  textCell: number;
  textHeader: number;
}

/** Cell font defaults used when the host sets no theme custom properties. */
export const DEFAULT_CELL_FONT_THEME = {
  textCell: 13,
  fontUi: 'system-ui, sans-serif',
  fontMono: 'ui-monospace, monospace',
} as const;

const num = (s: string, fallback: number): number => {
  const n = Number.parseFloat(s);
  return Number.isFinite(n) ? n : fallback;
};

export function resolveTheme(host: HTMLElement): ResolvedTheme {
  const cs = getComputedStyle(host);
  const v = (k: string, d = ''): string => cs.getPropertyValue(k).trim() || d;
  const bgRail = v('--fc-bg-rail', '#f3efe6');
  const fg = v('--fc-fg', '#15171c');
  const fgStrong = v('--fc-fg-strong', '#0a0c10');
  const rule = v('--fc-rule', '#e2dccd');
  const ruleStrong = v('--fc-rule-strong', '#cdc6b3');
  const accent = v('--fc-accent', '#d83a14');
  const headerBg = v('--fc-header-bg', bgRail);
  const headerRule = v('--fc-header-rule', rule);
  const headerFg = v('--fc-header-fg', '#6f6a5d');
  const outlineButtonBorder = v('--fc-outline-button-border', ruleStrong);
  return {
    bg: v('--fc-bg', '#faf7f1'),
    bgRail,
    bgElev: v('--fc-bg-elev', '#ffffff'),
    bgHeader: v('--fc-bg-header', '#ece7dc'),

    fg,
    fgMute: v('--fc-fg-mute', '#6f6a5d'),
    fgFaint: v('--fc-fg-faint', '#a8a294'),
    fgStrong,
    cellAutoFg: v('--fc-cell-auto-fg', fgStrong),

    rule,
    ruleStrong,
    freezeRule: v('--fc-freeze-rule', '#9b9b9b'),

    accent,
    accentFg: v('--fc-accent-fg', '#ffffff'),
    accentSoft: v('--fc-accent-soft', 'rgba(216,58,20,0.10)'),

    cellErrorFg: v('--fc-cell-error-fg', '#b32413'),
    cellFormulaFg: v('--fc-cell-formula-fg', '#38423c'),
    cellBoolFg: v('--fc-cell-bool-fg', '#2c4f2b'),
    cellNumFg: v('--fc-cell-num-fg', '#0a0c10'),

    hoverStripe: v('--fc-hover-stripe', 'rgba(21,23,28,0.025)'),
    headerBg,
    headerStripOuter: v('--fc-header-strip-outer', headerBg),
    headerStripInner: v('--fc-header-strip-inner', headerBg),
    headerRule,
    headerRuleOuter: v('--fc-header-rule-outer', headerRule),
    headerFg,
    headerFgActive: v('--fc-header-fg-active', '#15171c'),
    headerFgFiltered: v('--fc-header-fg-filtered', '#1c00ff'),
    headerActiveBar: v('--fc-header-active-bar', accent),
    headerFgActiveWeight: num(v('--fc-header-fg-active-weight', '600'), 600),
    outlineGutterBg: v('--fc-outline-gutter-bg', headerBg),
    outlineButtonFaceTop: v('--fc-outline-button-face-top', headerBg),
    outlineButtonFaceBottom: v('--fc-outline-button-face-bottom', headerBg),
    outlineButtonBorder,
    outlineButtonEdgeStart: v('--fc-outline-button-edge-start', outlineButtonBorder),
    outlineButtonEdgeEnd: v('--fc-outline-button-edge-end', outlineButtonBorder),
    outlineButtonFg: v('--fc-outline-button-fg', headerFg),
    outlineButtonShadow: v('--fc-outline-button-shadow', 'transparent'),
    outlineButtonHighlight: v('--fc-outline-button-highlight', 'transparent'),
    outlineRule: v('--fc-outline-rule', rule),
    outlineToggleBorder: v('--fc-outline-toggle-border', ruleStrong),
    outlineToggleFg: v('--fc-outline-toggle-fg', fg),

    pageBackdrop: v('--fc-page-backdrop', '#9b968a'),
    pagePaper: v('--fc-page-paper', '#ffffff'),
    pageEdge: v('--fc-page-edge', '#6f6a5d'),
    pageBandFg: v('--fc-page-band-fg', '#8d8779'),
    pageBreakAuto: v('--fc-page-break-auto', '#2f6fd0'),
    pageBreakManual: v('--fc-page-break-manual', '#1a4fa0'),
    pageNumberFg: v('--fc-page-number-fg', 'rgba(47,111,208,0.16)'),
    pageOutside: v('--fc-page-outside', 'rgba(120,120,120,0.28)'),

    fontUi: v('--fc-font-ui', DEFAULT_CELL_FONT_THEME.fontUi),
    fontMono: v('--fc-font-mono', DEFAULT_CELL_FONT_THEME.fontMono),

    textCell: num(
      v('--fc-cell-text-size', v('--fc-text-cell', `${DEFAULT_CELL_FONT_THEME.textCell}px`)),
      DEFAULT_CELL_FONT_THEME.textCell,
    ),
    textHeader: num(v('--fc-text-header', '11.5px'), 11.5),
  };
}
