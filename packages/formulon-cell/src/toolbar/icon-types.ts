/** One drawn path within an icon. */
export type IconSegment = {
  d: string;
  fill?: string;
  stroke?: string;
  strokeWidth?: string;
  strokeLinecap?: 'butt' | 'round' | 'square';
  strokeLinejoin?: 'bevel' | 'miter' | 'round';
  strokeDasharray?: string;
  /** `evenodd` knocks counters out of a fill without painting over them. */
  fillRule?: 'nonzero' | 'evenodd';
};

/** An icon is an ordered list of segments, painted back to front. */
export type IconDefinition = readonly IconSegment[];
