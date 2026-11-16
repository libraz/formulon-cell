/** Split a CSS font-family stack on top-level commas, ignoring commas inside quotes or parentheses. */
export const splitFontFamily = (family: string): string[] => {
  const parts: string[] = [];
  let start = 0;
  let quote = '';
  let depth = 0;
  for (let i = 0; i < family.length; i += 1) {
    const ch = family[i] ?? '';
    if (quote) {
      if (ch === quote && family[i - 1] !== '\\') quote = '';
      continue;
    }
    if (ch === '"' || ch === "'") {
      quote = ch;
      continue;
    }
    if (ch === '(') depth += 1;
    else if (ch === ')' && depth > 0) depth -= 1;
    else if (ch === ',' && depth === 0) {
      parts.push(family.slice(start, i).trim());
      start = i + 1;
    }
  }
  parts.push(family.slice(start).trim());
  return parts.filter(Boolean);
};

/** Re-emit a font-family stack with every non-identifier family as a CSS string built by `cssString`. */
export const safeFontFamily = (family: string, cssString: (value: string) => string): string =>
  splitFontFamily(family)
    .map((part) => {
      const unquoted =
        part.length >= 2 &&
        ((part.startsWith('"') && part.endsWith('"')) ||
          (part.startsWith("'") && part.endsWith("'")))
          ? part.slice(1, -1)
          : part;
      return /^[a-zA-Z][a-zA-Z0-9-]*$/.test(unquoted) ? unquoted : cssString(unquoted);
    })
    .join(', ');
