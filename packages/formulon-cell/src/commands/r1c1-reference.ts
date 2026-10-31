export interface R1C1Axis {
  /** Absolute axes store a 1-based coordinate; relative axes store a delta. */
  absolute: boolean;
  value: number;
}

export interface R1C1Endpoint {
  /** Character offsets of this atom in the original source string. */
  start: number;
  end: number;
  row: R1C1Axis;
  column: R1C1Axis;
}

export interface R1C1Reference {
  /** Span includes a sheet qualifier, if present, and every colon-chain atom. */
  start: number;
  end: number;
  qualifier: string;
  qualifierStart: number;
  qualifierEnd: number;
  endpoints: R1C1Endpoint[];
}

interface QualifierSpan {
  start: number;
  end: number;
  text: string;
  external: boolean;
}

interface ParsedAtom {
  endpoint: R1C1Endpoint;
  end: number;
}

interface ParsedChain {
  endpoints: R1C1Endpoint[];
  end: number;
}

const IDENTIFIER_CHAR = /[\p{L}\p{N}\p{M}_.\\]/u;
const R1C1_ATOM = /R(?:\[(-?\d+)\]|([1-9]\d*))?C(?:\[(-?\d+)\]|([1-9]\d*))?/iy;

const isIdentifierChar = (value: string | undefined): boolean =>
  value !== undefined && IDENTIFIER_CHAR.test(value);

function codePointBefore(text: string, index: number): string {
  if (index <= 0) return '';
  const last = text.charCodeAt(index - 1);
  if (last >= 0xdc00 && last <= 0xdfff && index > 1) {
    const high = text.charCodeAt(index - 2);
    if (high >= 0xd800 && high <= 0xdbff) return text.slice(index - 2, index);
  }
  return text.slice(index - 1, index);
}

function isIdentifierBefore(text: string, index: number): boolean {
  return isIdentifierChar(codePointBefore(text, index));
}

function isIdentifierAt(text: string, index: number): boolean {
  const point = text.codePointAt(index);
  return point !== undefined && isIdentifierChar(String.fromCodePoint(point));
}

function parseAtomAt(text: string, start: number): ParsedAtom | null {
  R1C1_ATOM.lastIndex = start;
  const match = R1C1_ATOM.exec(text);
  if (!match || match.index !== start) return null;
  const end = start + match[0].length;
  const rowAbsolute = match[2] !== undefined;
  const columnAbsolute = match[4] !== undefined;
  return {
    end,
    endpoint: {
      start,
      end,
      row: {
        absolute: rowAbsolute,
        value: rowAbsolute
          ? Number.parseInt(match[2] ?? '', 10)
          : match[1] !== undefined
            ? Number.parseInt(match[1], 10)
            : 0,
      },
      column: {
        absolute: columnAbsolute,
        value: columnAbsolute
          ? Number.parseInt(match[4] ?? '', 10)
          : match[3] !== undefined
            ? Number.parseInt(match[3], 10)
            : 0,
      },
    },
  };
}

function parseChainAt(text: string, start: number): ParsedChain | null {
  const first = parseAtomAt(text, start);
  if (!first) return null;
  const endpoints = [first.endpoint];
  let end = first.end;

  while (true) {
    const connector = /^\s*:\s*/u.exec(text.slice(end));
    if (!connector) break;
    const nextStart = end + connector[0].length;
    const next = parseAtomAt(text, nextStart);
    if (!next) break;
    endpoints.push(next.endpoint);
    end = next.end;
  }

  if (isIdentifierAt(text, end) || text[end] === '[' || text[end] === '(') return null;
  return { endpoints, end };
}

function readQuotedQualifier(text: string, start: number): QualifierSpan | null {
  if (text[start] !== "'") return null;
  let index = start + 1;
  while (index < text.length) {
    if (text[index] === "'") {
      if (text[index + 1] === "'") {
        index += 2;
        continue;
      }
      const afterQuote = index + 1;
      if (text[afterQuote] !== '!') return null;
      const end = afterQuote + 1;
      return {
        start,
        end,
        text: text.slice(start, end),
        external: text.slice(start + 1, index).startsWith('['),
      };
    }
    index += 1;
  }
  return null;
}

function readQuotedRegionEnd(text: string, start: number): number {
  let index = start + 1;
  while (index < text.length) {
    if (text[index] === "'") {
      if (text[index + 1] === "'") {
        index += 2;
        continue;
      }
      return index + 1;
    }
    index += 1;
  }
  return text.length;
}

function readSheetWordEnd(text: string, start: number): number | null {
  const first = text.codePointAt(start);
  if (first === undefined || !/[\p{L}_]/u.test(String.fromCodePoint(first))) return null;
  let index = start + (first > 0xffff ? 2 : 1);
  while (index < text.length) {
    const point = text.codePointAt(index);
    if (point === undefined || !/[\p{L}\p{N}\p{M}_.]/u.test(String.fromCodePoint(point))) break;
    index += point > 0xffff ? 2 : 1;
  }
  return index;
}

function readUnquotedQualifier(text: string, start: number): QualifierSpan | null {
  const firstEnd = readSheetWordEnd(text, start);
  if (firstEnd === null) return null;
  let end = firstEnd;
  if (text[end] === ':') {
    const secondEnd = readSheetWordEnd(text, end + 1);
    if (secondEnd !== null) end = secondEnd;
  }
  if (text[end] !== '!') return null;
  end += 1;
  return { start, end, text: text.slice(start, end), external: false };
}

function findBracketedEnd(text: string, start: number): number {
  let depth = 1;
  let index = start + 1;
  while (index < text.length && depth > 0) {
    if (text[index] === '[') depth += 1;
    else if (text[index] === ']') depth -= 1;
    else if (text[index] === "'" && ['[', ']', '@', '#', "'"].includes(text[index + 1] ?? '')) {
      index += 2;
      continue;
    }
    index += 1;
  }
  return index;
}

function readExternalWorkbookReference(
  text: string,
  start: number,
  bracketEnd: number,
): { qualifier: QualifierSpan; chain: ParsedChain | null } | null {
  const quotedSheet = readQuotedQualifier(text, bracketEnd);
  const sheet = quotedSheet ?? readUnquotedQualifier(text, bracketEnd);
  if (!sheet) return null;
  return {
    qualifier: { start, end: sheet.end, text: text.slice(start, sheet.end), external: true },
    chain: parseChainAt(text, sheet.end),
  };
}

function makeReference(
  start: number,
  qualifier: QualifierSpan | null,
  chain: ParsedChain,
): R1C1Reference {
  const qualifierStart = qualifier?.start ?? start;
  const qualifierEnd = qualifier?.end ?? start;
  return {
    start,
    end: chain.end,
    qualifier: qualifier?.text ?? '',
    qualifierStart,
    qualifierEnd,
    endpoints: chain.endpoints,
  };
}

/** Scan R1C1 cells, ranges, and colon chains without changing source text.
 *  References into another workbook are skipped unless `includeExternal`. */
export function scanR1C1References(
  text: string,
  options: { includeExternal?: boolean } = {},
): R1C1Reference[] {
  const includeExternal = options.includeExternal === true;
  const references: R1C1Reference[] = [];
  let index = text.startsWith('=') ? 1 : 0;

  while (index < text.length) {
    const char = text[index] ?? '';
    if (char === '"') {
      index += 1;
      while (index < text.length) {
        if (text[index] === '"' && text[index + 1] === '"') {
          index += 2;
          continue;
        }
        if (text[index] === '"') {
          index += 1;
          break;
        }
        index += 1;
      }
      continue;
    }

    if (char === '[') {
      const end = findBracketedEnd(text, index);
      const external = readExternalWorkbookReference(text, index, end);
      if (!external) {
        index = end;
        continue;
      }
      if (external.chain && includeExternal) {
        references.push(makeReference(index, external.qualifier, external.chain));
      }
      index = external.chain?.end ?? external.qualifier.end;
      continue;
    }

    if (char === "'") {
      const qualifier = readQuotedQualifier(text, index);
      if (!qualifier) {
        index = readQuotedRegionEnd(text, index);
        continue;
      }
      const chain = parseChainAt(text, qualifier.end);
      if (chain) {
        if (!qualifier.external || includeExternal) {
          references.push(makeReference(index, qualifier, chain));
        }
        index = chain.end;
      } else {
        index = qualifier.end;
      }
      continue;
    }

    if (isIdentifierBefore(text, index)) {
      index += 1;
      continue;
    }

    const qualifier = readUnquotedQualifier(text, index);
    if (qualifier) {
      const chain = parseChainAt(text, qualifier.end);
      if (chain) {
        references.push(makeReference(index, qualifier, chain));
        index = chain.end;
      } else {
        index = qualifier.end;
      }
      continue;
    }

    const chain = parseChainAt(text, index);
    if (chain) {
      references.push(makeReference(index, null, chain));
      index = chain.end;
    } else {
      index += 1;
    }
  }

  return references;
}
