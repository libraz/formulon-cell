import type { Pagination } from '@libraz/formulon';
import type { PageSetup } from '../store/store.js';
import { cssString } from '../text-escape.js';
import { effectivePrintMargins, PAPER_DIMENSIONS } from './page-geometry.js';

/** Build the @page CSS string for a given setup. Browsers honour orientation
 *  and (for print preview / PDF) the size keyword. */
/** Context used to expand header/footer field codes into concrete text. */
export interface HeaderFooterContext {
  sheetName: string;
  fileName: string;
  date: string;
  time: string;
}

/**
 * Expand an Excel header/footer field string into a CSS `content` value for a
 * `@page` margin box. Static codes (`&D` date, `&T` time, `&F` file, `&A` sheet)
 * become quoted literals; page-relative codes (`&P` page, `&N` pages) become
 * `counter(page)` / `counter(pages)` so the browser fills the real numbers on
 * every printed page. `&&` is a literal ampersand; formatting codes (`&B`,
 * `&"font"`, size runs, section markers) are stripped. Returns `''` when the
 * field is empty.
 */
export function headerFieldToCssContent(raw: string, ctx: HeaderFooterContext): string {
  if (!raw) return '';
  const parts: string[] = [];
  let buf = '';
  const flush = (): void => {
    if (buf) {
      parts.push(cssString(buf));
      buf = '';
    }
  };
  for (let i = 0; i < raw.length; i += 1) {
    const ch = raw[i] ?? '';
    if (ch !== '&') {
      buf += ch;
      continue;
    }
    const next = raw[i + 1] ?? '';
    const upper = next.toUpperCase();
    if (next === '&') {
      buf += '&';
      i += 1;
    } else if (upper === 'P' || upper === 'N') {
      flush();
      parts.push(upper === 'P' ? 'counter(page)' : 'counter(pages)');
      i += 1;
      // Drop an Excel page offset like `&P+1` — CSS counters cannot offset.
      const rest = raw.slice(i + 1);
      const off = /^[+-]\d+/.exec(rest);
      if (off) i += off[0].length;
    } else if (upper === 'D') {
      buf += ctx.date;
      i += 1;
    } else if (upper === 'T') {
      buf += ctx.time;
      i += 1;
    } else if (upper === 'F') {
      buf += ctx.fileName;
      i += 1;
    } else if (upper === 'A') {
      buf += ctx.sheetName;
      i += 1;
    } else if (next === '"') {
      // &"font name" — skip the quoted font spec entirely.
      const close = raw.indexOf('"', i + 2);
      i = close >= 0 ? close : i + 1;
    } else if (/[0-9]/.test(next)) {
      // &nn font-size run — drop the digits.
      let j = i + 1;
      while (j < raw.length && /[0-9]/.test(raw[j] ?? '')) j += 1;
      i = j - 1;
    } else {
      // Single-letter formatting / section code (&B, &I, &U, &L, &C, &R, &Z…).
      i += 1;
    }
  }
  flush();
  return parts.length ? parts.join(' ') : '';
}

export function buildPageRule(
  setup: PageSetup,
  ctx: HeaderFooterContext,
  native?: Pagination,
): string {
  const size = native
    ? `${native.paper.widthPt}pt ${native.paper.heightPt}pt`
    : `${PAPER_DIMENSIONS[setup.paperSize] ?? 'A4'} ${setup.orientation}`;
  const m = native
    ? {
        top: `${native.margins.topIn * 72}pt`,
        right: `${native.margins.rightIn * 72}pt`,
        bottom: `${native.margins.bottomIn * 72}pt`,
        left: `${native.margins.leftIn * 72}pt`,
      }
    : (() => {
        const margins = effectivePrintMargins(setup);
        return {
          top: `${margins.top}in`,
          right: `${margins.right}in`,
          bottom: `${margins.bottom}in`,
          left: `${margins.left}in`,
        };
      })();
  const boxes: [string, string | undefined][] = [
    ['top-left', setup.headerLeft],
    ['top-center', setup.headerCenter],
    ['top-right', setup.headerRight],
    ['bottom-left', setup.footerLeft],
    ['bottom-center', setup.footerCenter],
    ['bottom-right', setup.footerRight],
  ];
  const marginBoxes = boxes
    .map(([box, raw]) => {
      const content = headerFieldToCssContent(raw ?? '', ctx);
      if (!content) return '';
      return `@${box} { content: ${content}; font-size: 9pt; color: #555; white-space: pre; }`;
    })
    .filter(Boolean)
    .join(' ');
  const base = `@page { size: ${size}; margin: ${m.top} ${m.right} ${m.bottom} ${m.left};`;
  return marginBoxes ? `${base} ${marginBoxes} }` : `${base} }`;
}
