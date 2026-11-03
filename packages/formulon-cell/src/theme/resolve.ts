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
  return {
    bg: v('--fc-bg', '#faf7f1'),
    bgRail: v('--fc-bg-rail', '#f3efe6'),
    bgElev: v('--fc-bg-elev', '#ffffff'),
    bgHeader: v('--fc-bg-header', '#ece7dc'),

    fg: v('--fc-fg', '#15171c'),
    fgMute: v('--fc-fg-mute', '#6f6a5d'),
    fgFaint: v('--fc-fg-faint', '#a8a294'),
    fgStrong: v('--fc-fg-strong', '#0a0c10'),
    cellAutoFg: v('--fc-cell-auto-fg', v('--fc-fg-strong', '#0a0c10')),

    rule: v('--fc-rule', '#e2dccd'),
    ruleStrong: v('--fc-rule-strong', '#cdc6b3'),
    freezeRule: v('--fc-freeze-rule', '#9b9b9b'),

    accent: v('--fc-accent', '#d83a14'),
    accentFg: v('--fc-accent-fg', '#ffffff'),
    accentSoft: v('--fc-accent-soft', 'rgba(216,58,20,0.10)'),

    cellErrorFg: v('--fc-cell-error-fg', '#b32413'),
    cellFormulaFg: v('--fc-cell-formula-fg', '#38423c'),
    cellBoolFg: v('--fc-cell-bool-fg', '#2c4f2b'),
    cellNumFg: v('--fc-cell-num-fg', '#0a0c10'),

    hoverStripe: v('--fc-hover-stripe', 'rgba(21,23,28,0.025)'),
    headerBg: v('--fc-header-bg', v('--fc-bg-rail', '#f3efe6')),
    headerStripOuter: v(
      '--fc-header-strip-outer',
      v('--fc-header-bg', v('--fc-bg-rail', '#f3efe6')),
    ),
    headerStripInner: v(
      '--fc-header-strip-inner',
      v('--fc-header-bg', v('--fc-bg-rail', '#f3efe6')),
    ),
    headerRule: v('--fc-header-rule', v('--fc-rule', '#e2dccd')),
    headerRuleOuter: v('--fc-header-rule-outer', v('--fc-header-rule', v('--fc-rule', '#e2dccd'))),
    headerFg: v('--fc-header-fg', '#6f6a5d'),
    headerFgActive: v('--fc-header-fg-active', '#15171c'),
    headerActiveBar: v('--fc-header-active-bar', v('--fc-accent', '#d83a14')),
    headerFgActiveWeight: num(v('--fc-header-fg-active-weight', '600'), 600),
    outlineGutterBg: v('--fc-outline-gutter-bg', v('--fc-header-bg', v('--fc-bg-rail', '#f3efe6'))),
    outlineButtonFaceTop: v(
      '--fc-outline-button-face-top',
      v('--fc-header-bg', v('--fc-bg-rail', '#f3efe6')),
    ),
    outlineButtonFaceBottom: v(
      '--fc-outline-button-face-bottom',
      v('--fc-header-bg', v('--fc-bg-rail', '#f3efe6')),
    ),
    outlineButtonBorder: v('--fc-outline-button-border', v('--fc-rule-strong', '#cdc6b3')),
    outlineButtonFg: v('--fc-outline-button-fg', v('--fc-header-fg', '#6f6a5d')),
    outlineButtonShadow: v('--fc-outline-button-shadow', 'transparent'),

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
