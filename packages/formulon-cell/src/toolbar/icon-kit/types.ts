/** One drawn path within an icon. */
export type IconSegment = {
  d: string;
  fill?: string;
  stroke?: string;
  strokeWidth?: string;
  strokeLinecap?: 'butt' | 'round' | 'square';
  strokeLinejoin?: 'arcs' | 'bevel' | 'miter' | 'miter-clip' | 'round';
  strokeDasharray?: string;
  /** `evenodd` knocks counters out of a letterform without painting over them. */
  fillRule?: 'nonzero' | 'evenodd';
  /**
   * SVG transform applied to this path. Used to place reusable artwork —
   * letterforms above all — at the size a composition needs without
   * re-authoring its path data.
   */
  transform?: string;
};

/** An icon is an ordered list of segments, painted back to front. */
export type IconDefinition = readonly IconSegment[];
