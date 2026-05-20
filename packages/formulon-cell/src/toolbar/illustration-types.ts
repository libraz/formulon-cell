// Types describing the ribbon's "Insert → Shapes / Picture / Screenshot"
// surface. The toolbar layer only needs the discriminant strings to thread
// callbacks through; the actual session-illustration implementation lives in
// the consumer.

export type SessionIllustrationKind = 'image' | 'shape' | 'screenshot';
export type SessionShapeKind =
  | 'rectangle'
  | 'rounded-rectangle'
  | 'oval'
  | 'triangle'
  | 'diamond'
  | 'line'
  | 'arrow';
