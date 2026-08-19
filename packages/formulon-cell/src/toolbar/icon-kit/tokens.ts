/**
 * Design tokens shared by every ribbon icon.
 *
 * Icons are authored as path data rather than drawn, so the only thing keeping
 * them consistent is this file: a fixed canvas, one optical size, a five-step
 * stroke scale, and a closed palette. Anything a glyph needs that is not here
 * belongs here first.
 */

/** Every icon is authored on a 24x24 viewBox. */
export const CANVAS = 24;

/** Optical centre of the canvas. Every glyph's inked box is centred here. */
export const CENTER = 12;

/**
 * Target extent of a glyph's inked area along its longest axis.
 *
 * The ribbon mixes these icons with the Fluent set, whose inked box has a
 * median longest axis of 20 units. Matching it keeps the two sets the same
 * visual weight when they sit side by side.
 */
export const OPTICAL = 20;

/** Reduced optical size for a glyph that carries a corner badge. */
export const OPTICAL_BADGED = 17;

/** Nothing may ink outside this inset, so icons never clip when scaled down. */
export const SAFE_INSET = 1;

/**
 * Stroke widths, in canvas units. Five steps only — the previous hand-authored
 * set had drifted to 27 distinct values, which is the main reason strokes read
 * as inconsistent between neighbouring icons.
 */
export const STROKE = {
  /** Interior rules and other detail that must not compete with the outline. */
  hairline: '1',
  /** Secondary structure: inner frames, text rules, small detail. */
  thin: '1.25',
  /** Default outline weight. */
  regular: '1.5',
  /** Emphasis: arrows, primary action marks. */
  bold: '1.75',
  /** Reserved for marks that must survive at 16px, such as clear/delete Xs. */
  heavy: '2',
} as const;

export type StrokeWeight = keyof typeof STROKE;

/** Resolve a stroke step to its numeric width. */
export const strokeWidth = (weight: StrokeWeight): number => Number(STROKE[weight]);

/**
 * Closed palette. Roles, not hues: an icon picks `danger` because it destroys
 * something, not because it wants red.
 */
export const PALETTE = {
  /** Primary outline and text-like shapes. */
  ink: '#1f1f1f',
  /** Filled surfaces that read as blank paper. */
  paper: '#ffffff',
  /** Off-white surface used where paper would blend into the background. */
  mute: '#f3f2f1',
  /** Secondary outline: sheet frames, inactive structure. */
  grid: '#8a8f98',
  /** Interior rules of a sheet or table. */
  gridLight: '#d9d9d9',

  /** Primary action colour: apply, confirm, calculate. */
  accent: '#107c41',
  accentDeep: '#0b5a2f',
  accentSoft: '#c7e2d3',

  /** Movement, navigation, references, links. */
  info: '#2f75b5',
  infoDeep: '#1c5386',
  infoSoft: '#cfe0f2',

  /** Destructive: clear, delete, remove. */
  danger: '#c00000',
  dangerSoft: '#f5cdcd',

  /** Attention and annotation. */
  warn: '#ffd966',
  warnDeep: '#b79500',
  warnSoft: '#fdf3bf',

  /** Secondary categorical colour, mainly in charts and shapes. */
  alt: '#ed7d31',
  altDeep: '#b35a00',

  /** Formatting and style operations. */
  violet: '#8a4fb7',
  violetSoft: '#f7e1ff',

  /** Clipboard clip. */
  tan: '#f4c27a',
  tanDeep: '#9a6a2f',

  /**
   * Theme previews. These depict a theme's own surface rather than following
   * the icon set's, so they are named here instead of being written inline.
   */
  themeDarkSurface: '#2b2b2b',
  themeDarkBand: '#111111',
  themeDarkText: '#8f8f8f',
  themeContrastSurface: '#000000',

  /** Conditional-formatting colour scale, low to high. */
  scaleHigh: '#63be7b',
  scaleMid: '#ffeb84',
  scaleLow: '#f8696b',
} as const;

export type PaletteColor = (typeof PALETTE)[keyof typeof PALETTE];
