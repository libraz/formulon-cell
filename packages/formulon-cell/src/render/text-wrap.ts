/** Horizontal cell padding used by both canvas painting and AutoFit. */
export const CELL_TEXT_HORIZONTAL_PADDING_PX = 6;

/** Pixels added to the leading edge for each Excel indent level. */
export const CELL_TEXT_INDENT_PX = 8;

export interface CellTextHorizontalMetrics {
  /** Horizontal inset on each side of the cell. */
  padX: number;
  /** Leading inset for the requested indent level. */
  indentPx: number;
  /** Width available to the cell text after both insets. */
  availableWidth: number;
}

/**
 * Resolve the screen-space horizontal geometry shared by painting, wrapping,
 * shrink-to-fit, and numeric overflow handling. AutoFit calls this with a
 * scale of one because its inputs are stored logical widths.
 */
export const cellTextHorizontalMetrics = (
  cellWidth: number,
  indentLevel = 0,
  scale = 1,
): CellTextHorizontalMetrics => {
  const safeScale = Number.isFinite(scale) && scale > 0 ? scale : 1;
  const safeIndent = Number.isFinite(indentLevel) ? Math.max(0, indentLevel) : 0;
  const padX = (CELL_TEXT_HORIZONTAL_PADDING_PX / 2) * safeScale;
  const indentPx = CELL_TEXT_INDENT_PX * safeIndent * safeScale;
  return {
    padX,
    indentPx,
    availableWidth: Math.max(1, cellWidth - padX * 2 - indentPx),
  };
};

export interface TextWrapOptions {
  /** Width available to the text after cell padding and indentation. */
  maxWidth: number;
  /** Measure a candidate line in the same font as the caller. */
  measureText: (text: string) => number;
  /** BCP-47 locale used for word and grapheme segmentation. */
  locale?: string;
}

interface SegmentRecord {
  segment: string;
}

interface SegmenterLike {
  segment(input: string): Iterable<SegmentRecord>;
}

interface IntlWithSegmenter {
  Segmenter?: new (
    locales?: string | readonly string[],
    options?: { granularity?: 'word' | 'grapheme' },
  ) => SegmenterLike;
}

const combiningMark = /\p{Mark}/u;

/** Width available to a cell's wrapped text under the shared canvas/AutoFit contract. */
export const availableTextWidth = (cellWidth: number, indentPx = 0): number =>
  Math.max(1, cellWidth - CELL_TEXT_HORIZONTAL_PADDING_PX - Math.max(0, indentPx));

const segmentWithIntl = (
  text: string,
  locale: string | undefined,
  granularity: 'word' | 'grapheme',
): string[] | null => {
  const Segmenter = (globalThis.Intl as IntlWithSegmenter | undefined)?.Segmenter;
  if (!Segmenter) return null;
  try {
    const segmenter = new Segmenter(locale ?? 'en-US', { granularity });
    return Array.from(segmenter.segment(text), (entry) => entry.segment);
  } catch {
    // Invalid/unsupported locales and partial Intl implementations fall back
    // to the deterministic segmenters below.
    return null;
  }
};

const isVariationSelector = (codePoint: number): boolean =>
  (codePoint >= 0xfe00 && codePoint <= 0xfe0f) || (codePoint >= 0xe0100 && codePoint <= 0xe01ef);

const isEmojiModifier = (codePoint: number): boolean =>
  codePoint >= 0x1f3fb && codePoint <= 0x1f3ff;

const isRegionalIndicator = (codePoint: number): boolean =>
  codePoint >= 0x1f1e6 && codePoint <= 0x1f1ff;

/**
 * Grapheme fallback for hosts without Intl.Segmenter. It keeps combining marks,
 * variation selectors, emoji modifiers, regional-indicator pairs, and ZWJ
 * sequences attached to their preceding base. Modern browsers use the exact
 * Unicode grapheme algorithm through Intl.Segmenter above.
 */
const fallbackGraphemes = (text: string): string[] => {
  const out: string[] = [];
  for (const char of text) {
    const codePoint = char.codePointAt(0) ?? 0;
    const previous = out.at(-1);
    const previousCodePoint =
      Array.from(previous ?? '')
        .at(-1)
        ?.codePointAt(0) ?? 0;
    const joinsPrevious =
      previous !== undefined &&
      (codePoint === 0x200d ||
        previous.endsWith('\u200d') ||
        combiningMark.test(char) ||
        isVariationSelector(codePoint) ||
        isEmojiModifier(codePoint) ||
        (isRegionalIndicator(codePoint) &&
          isRegionalIndicator(previousCodePoint) &&
          [...previous].length === 1));
    if (joinsPrevious) out[out.length - 1] = `${previous}${char}`;
    else out.push(char);
  }
  return out;
};

const graphemes = (text: string, locale: string | undefined): string[] =>
  segmentWithIntl(text, locale, 'grapheme') ?? fallbackGraphemes(text);

const NON_BREAKING_CONNECTORS = new Set(['\u00a0', '\u202f', '\u2060']);

const containsNonBreakingConnector = (text: string): boolean =>
  [...text].some((character) => NON_BREAKING_CONNECTORS.has(character));

const startsWithNonBreakingConnector = (text: string): boolean =>
  NON_BREAKING_CONNECTORS.has([...text][0] ?? '');

const endsWithNonBreakingConnector = (text: string): boolean =>
  NON_BREAKING_CONNECTORS.has([...text].at(-1) ?? '');

const isBreakableWhitespace = (text: string): boolean =>
  /^\s+$/u.test(text) && !containsNonBreakingConnector(text);

const mergeNonBreakingConnectorSegments = (segments: readonly string[]): string[] => {
  const merged: string[] = [];
  for (const segment of segments) {
    const previous = merged.at(-1);
    if (
      previous !== undefined &&
      (endsWithNonBreakingConnector(previous) || startsWithNonBreakingConnector(segment))
    ) {
      merged[merged.length - 1] = `${previous}${segment}`;
    } else {
      merged.push(segment);
    }
  }
  return merged;
};

const wordSegments = (text: string, locale: string | undefined): string[] =>
  mergeNonBreakingConnectorSegments(
    segmentWithIntl(text, locale, 'word') ?? text.match(/\s+|[^\s]+/gu) ?? [text],
  );

const measuredWidth = (text: string, measureText: (value: string) => number): number => {
  const width = measureText(text);
  return Number.isFinite(width) && width >= 0 ? width : Number.POSITIVE_INFINITY;
};

const NO_BREAK_BEFORE = new Set('、。，．・：；？！）］｝〕〉》」』】'.split(''));
const NO_BREAK_AFTER = new Set('（［｛〔〈《「『【'.split(''));
const CJK_LIKE =
  /[\u1100-\u11ff\u2e80-\u2fff\u3000-\u30ff\u3130-\u318f\u3400-\u4dbf\u4e00-\u9fff\ua960-\ua97f\uac00-\ud7af\uf900-\ufaff]/u;

interface WrapUnit {
  text: string;
  graphemes: string[];
  keepTogether: boolean;
}

/** Merge only the small, observable spreadsheet punctuation rules. */
const mergePunctuationUnits = (
  segments: readonly string[],
  locale: string | undefined,
  segmentsAreGraphemes = false,
): WrapUnit[] => {
  const units: WrapUnit[] = [];
  for (const segment of segments) {
    const next: WrapUnit = {
      text: segment,
      graphemes: segmentsAreGraphemes ? [segment] : graphemes(segment, locale),
      keepTogether: containsNonBreakingConnector(segment),
    };
    const previous = units.at(-1);
    const startsWithClosing = NO_BREAK_BEFORE.has(next.graphemes[0] ?? '');
    const previousEndsWithOpening = NO_BREAK_AFTER.has(previous?.graphemes.at(-1) ?? '');
    if (previous && (startsWithClosing || previousEndsWithOpening)) {
      previous.text += next.text;
      previous.graphemes.push(...next.graphemes);
      previous.keepTogether ||= next.keepTogether;
    } else {
      units.push(next);
    }
  }
  return units;
};

const wrapParagraph = (paragraph: string, options: TextWrapOptions): string[] => {
  if (paragraph.length === 0) return [''];
  const maxWidth = Number.isFinite(options.maxWidth) ? Math.max(1, options.maxWidth) : Infinity;
  if (options.maxWidth <= 0) return [paragraph];

  const lines: string[] = [];
  let line = '';
  let afterWrappedBoundary = false;
  const flushWrapped = (): void => {
    lines.push(line.trimEnd());
    line = '';
    afterWrappedBoundary = true;
  };
  const flushFinal = (): void => {
    lines.push(line);
    line = '';
    afterWrappedBoundary = false;
  };

  const appendUnit = (unit: WrapUnit): void => {
    const fitsWhole =
      measuredWidth(unit.text, options.measureText) <= maxWidth &&
      (!line || measuredWidth(line + unit.text, options.measureText) <= maxWidth);
    const leadingWhitespace =
      !afterWrappedBoundary && lines.length === 0 && isBreakableWhitespace(line);
    const unitWidth = measuredWidth(unit.text, options.measureText);
    if (fitsWhole) {
      line += unit.text;
      return;
    }

    // Non-breaking connectors belong to the surrounding word/punctuation unit.
    // Keep that unit intact when it cannot fit; unlike CJK punctuation, relaxing
    // this boundary would change the authored non-breaking whitespace semantics.
    if (unit.keepTogether) {
      if (line && !leadingWhitespace) flushWrapped();
      line += unit.text;
      afterWrappedBoundary = false;
      return;
    }

    // A CJK word segment can still fill the remainder of the current line by
    // grapheme. Latin/Thai word segments stay whole unless the word itself is
    // wider than the cell, preserving Intl.Segmenter word boundaries.
    const splitForPacking = unitWidth > maxWidth || CJK_LIKE.test(unit.text) || leadingWhitespace;
    if (!splitForPacking) {
      if (line) flushWrapped();
      line = unit.text;
      afterWrappedBoundary = false;
      return;
    }

    const groupedPieces = mergePunctuationUnits(unit.graphemes, options.locale, true);
    const pieces = groupedPieces.some(
      (piece) => measuredWidth(piece.text, options.measureText) > maxWidth,
    )
      ? unit.graphemes.map((text) => ({ text, graphemes: [text], keepTogether: false }))
      : groupedPieces;
    let first = true;
    for (const piece of pieces) {
      // Authored leading whitespace plus the first grapheme is one visual
      // unit. Keep it intact even when it must overflow the cell.
      if (leadingWhitespace && first) {
        line += piece.text;
        first = false;
        continue;
      }
      first = false;
      const candidate = line + piece.text;
      if (!line || measuredWidth(candidate, options.measureText) <= maxWidth) {
        line = candidate;
      } else {
        flushWrapped();
        line = piece.text;
        afterWrappedBoundary = false;
      }
    }
  };

  const segments = wordSegments(paragraph, options.locale);
  let index = 0;
  while (index < segments.length) {
    const segment = segments[index] ?? '';
    if (isBreakableWhitespace(segment)) {
      if (!line) {
        if (!afterWrappedBoundary) line = segment;
      } else if (measuredWidth(line + segment, options.measureText) <= maxWidth) {
        line += segment;
      } else {
        flushWrapped();
      }
      index += 1;
      continue;
    }

    const contiguousWords: string[] = [];
    while (index < segments.length && !isBreakableWhitespace(segments[index] ?? '')) {
      contiguousWords.push(segments[index] ?? '');
      index += 1;
    }
    for (const unit of mergePunctuationUnits(contiguousWords, options.locale)) appendUnit(unit);
  }
  if (line || lines.length === 0) flushFinal();
  return lines;
};

/**
 * Greedy locale-aware cell wrapping shared by canvas and AutoFit. Explicit
 * line breaks are split before word segmentation, so empty paragraphs remain
 * visible as empty lines. Oversized words fall back to grapheme-safe breaks.
 */
export const wrapText = (text: string, options: TextWrapOptions): string[] => {
  const paragraphs = text.split(/\r\n|\r|\n/);
  const lines = paragraphs.flatMap((paragraph) => wrapParagraph(paragraph, options));
  return lines.length > 0 ? lines : [''];
};
