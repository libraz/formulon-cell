import type { Locale } from '../../i18n/strings.js';

/** Letter glyphs observed in Excel's localized Home font controls. */
export function localizedFormatIcon(
  name: string,
  locale: Locale,
  ownerDocument: Document,
): SVGSVGElement | null {
  const letters: Partial<Record<string, string>> =
    locale === 'de'
      ? { bold: 'F', italic: 'K' }
      : locale === 'fr'
        ? { bold: 'G', underline: 'S' }
        : {};
  const glyph = letters[name];
  if (!glyph) return null;
  const namespace = 'http://www.w3.org/2000/svg';
  const svg = ownerDocument.createElementNS(namespace, 'svg');
  svg.classList.add('fc-tb__rb-icon');
  svg.setAttribute('viewBox', '0 0 24 24');
  svg.setAttribute('aria-hidden', 'true');
  svg.setAttribute('focusable', 'false');
  const letter = ownerDocument.createElementNS(namespace, 'text');
  letter.textContent = glyph;
  letter.setAttribute('x', '12');
  letter.setAttribute('y', '18');
  letter.setAttribute('text-anchor', 'middle');
  letter.setAttribute('fill', 'currentColor');
  letter.setAttribute('font-family', 'var(--fc-tb-font, system-ui)');
  letter.setAttribute('font-size', '18');
  if (name === 'bold') letter.setAttribute('font-weight', '700');
  else if (name === 'italic') letter.setAttribute('font-style', 'italic');
  svg.append(letter);
  if (name === 'underline') {
    const line = ownerDocument.createElementNS(namespace, 'path');
    line.setAttribute('d', 'M5 21h14');
    line.setAttribute('stroke', 'currentColor');
    line.setAttribute('stroke-width', '1.5');
    svg.append(line);
  }
  return svg;
}
