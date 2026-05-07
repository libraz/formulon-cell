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
  rule: string;
  ruleStrong: string;
  accent: string;
  accentFg: string;
  accentSoft: string;
  cellErrorFg: string;
  cellFormulaFg: string;
  cellBoolFg: string;
  cellNumFg: string;
  hoverStripe: string;
  headerFg: string;
  headerFgActive: string;

  fontUi: string;
  fontMono: string;

  textCell: number;
  textHeader: number;
}

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

    rule: v('--fc-rule', '#e2dccd'),
    ruleStrong: v('--fc-rule-strong', '#cdc6b3'),

    accent: v('--fc-accent', '#d83a14'),
    accentFg: v('--fc-accent-fg', '#ffffff'),
    accentSoft: v('--fc-accent-soft', 'rgba(216,58,20,0.10)'),

    cellErrorFg: v('--fc-cell-error-fg', '#b32413'),
    cellFormulaFg: v('--fc-cell-formula-fg', '#38423c'),
    cellBoolFg: v('--fc-cell-bool-fg', '#2c4f2b'),
    cellNumFg: v('--fc-cell-num-fg', '#0a0c10'),

    hoverStripe: v('--fc-hover-stripe', 'rgba(21,23,28,0.025)'),
    headerFg: v('--fc-header-fg', '#6f6a5d'),
    headerFgActive: v('--fc-header-fg-active', '#15171c'),

    fontUi: v('--fc-font-ui', 'system-ui, sans-serif'),
    fontMono: v('--fc-font-mono', 'ui-monospace, monospace'),

    textCell: num(v('--fc-text-cell', '13px'), 13),
    textHeader: num(v('--fc-text-header', '11.5px'), 11.5),
  };
}
