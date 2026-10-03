import type { EngineHeaderFooter, EnginePageSetupExtras, EnginePrintOptions } from './types.js';

/** Partial `<pageSetup>` attribute update. An omitted key is left alone;
 *  `false`, `''` and `null` remove the attribute. */
export type EnginePageSetupExtrasInput = Partial<{
  blackAndWhite: boolean;
  draft: boolean;
  cellComments: string;
  errors: string;
  pageOrder: string;
  firstPageNumber: number | null;
  useFirstPageNumber: boolean;
  horizontalDpi: number | null;
  verticalDpi: number | null;
}>;

/** Wrapper element used to parse a bare fragment. It declares the
 *  relationships prefix so a `<pageSetup r:id="...">` — which the engine hands
 *  back without a declaration of its own, because the worksheet root carried
 *  one — parses instead of failing as an undeclared prefix. */
const FRAGMENT_WRAPPER = 'fc-fragment';
const RELATIONSHIPS_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** Parse a `<pageSetup>` fragment inside the namespace-declaring wrapper.
 *  Returns null for the empty string the engine hands back when the sheet
 *  declares no page setup, and for a host with no XML parser. */
function parsePageSetupFragment(xml: string): Element | null {
  if (!xml || typeof DOMParser === 'undefined') return null;
  const doc = new DOMParser().parseFromString(
    `<${FRAGMENT_WRAPPER} xmlns:r="${RELATIONSHIPS_NS}">${xml}</${FRAGMENT_WRAPPER}>`,
    'application/xml',
  );
  const root = doc.documentElement;
  if (!root || root.nodeName !== FRAGMENT_WRAPPER) return null;
  if (root.getElementsByTagName('parsererror').length > 0) return null;
  const el = root.firstElementChild;
  return el?.nodeName === 'pageSetup' ? el : null;
}

/**
 * Apply `extras` to a `<pageSetup>` fragment and return the new fragment.
 *
 * Returns null when the fragment cannot be parsed or the host has no XML
 * serializer, so the caller can tell "nothing to do" from "could not do it".
 * The wrapper is serialized and then unwrapped rather than serializing the
 * element on its own: that keeps the `xmlns:r` declaration on the wrapper,
 * where the engine's own fragment does not carry it either.
 */
export function mergePageSetupFragment(
  xml: string,
  extras: EnginePageSetupExtrasInput,
): string | null {
  if (typeof DOMParser === 'undefined' || typeof XMLSerializer === 'undefined') return null;
  const el = parsePageSetupFragment(xml || '<pageSetup/>');
  if (!el) return null;
  const put = (name: string, value: string | null): void => {
    if (value === null || value === '') el.removeAttribute(name);
    else el.setAttribute(name, value);
  };
  const putFlag = (name: string, value: boolean | undefined): void => {
    if (value === undefined) return;
    put(name, value ? '1' : null);
  };
  const putUint = (name: string, value: number | null | undefined): void => {
    if (value === undefined) return;
    put(name, value === null || !Number.isFinite(value) ? null : String(Math.trunc(value)));
  };
  putFlag('blackAndWhite', extras.blackAndWhite);
  putFlag('draft', extras.draft);
  putFlag('useFirstPageNumber', extras.useFirstPageNumber);
  if (extras.cellComments !== undefined) put('cellComments', extras.cellComments || null);
  if (extras.errors !== undefined) put('errors', extras.errors || null);
  if (extras.pageOrder !== undefined) put('pageOrder', extras.pageOrder || null);
  putUint('firstPageNumber', extras.firstPageNumber);
  putUint('horizontalDpi', extras.horizontalDpi);
  putUint('verticalDpi', extras.verticalDpi);

  const wrapped = new XMLSerializer().serializeToString(el.parentNode as Element);
  const open = wrapped.indexOf('>');
  const close = wrapped.lastIndexOf(`</${FRAGMENT_WRAPPER}>`);
  if (open < 0 || close < open) return null;
  return wrapped.slice(open + 1, close);
}

/** Read an unsigned integer attribute; null when absent or unparseable. */
function xmlUint(el: Element | null, name: string): number | null {
  const raw = el?.getAttribute(name);
  if (raw === null || raw === undefined) return null;
  const n = Number.parseInt(raw, 10);
  return Number.isFinite(n) && n >= 0 ? n : null;
}

/** Parse a single-element OOXML fragment. Returns null for the empty string
 *  the engine hands back when the sheet declares no such element, for a
 *  fragment that fails to parse, and for a host with no XML parser. */
function parseXmlElement(xml: string): Element | null {
  if (!xml || typeof DOMParser === 'undefined') return null;
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const root = doc.documentElement;
  if (!root || root.getElementsByTagName('parsererror').length > 0) return null;
  return root.nodeName === 'parsererror' ? null : root;
}

/** Read an OOXML boolean attribute. Both spellings a schema-valid file may
 *  use are accepted; an absent attribute falls back to the schema default. */
function xmlFlag(el: Element | null, name: string, fallback: boolean): boolean {
  const raw = el?.getAttribute(name);
  if (raw === null || raw === undefined) return fallback;
  return raw === '1' || raw === 'true';
}

/** Text of a header/footer section child. The XML parser has already decoded
 *  the entities, which is the spelling `setSheetHeaderFooter` expects back. */
function xmlSectionText(el: Element | null, name: string): string {
  const child = el?.getElementsByTagName(name)[0];
  return child?.textContent ?? '';
}

/** Typed `<pageSetup>` attributes beside the typed setter's own. */
export function readPageSetupExtras(xml: string): EnginePageSetupExtras {
  const el = parsePageSetupFragment(xml);
  return {
    blackAndWhite: xmlFlag(el, 'blackAndWhite', false),
    draft: xmlFlag(el, 'draft', false),
    cellComments: el?.getAttribute('cellComments') ?? '',
    errors: el?.getAttribute('errors') ?? '',
    pageOrder: el?.getAttribute('pageOrder') ?? '',
    firstPageNumber: xmlUint(el, 'firstPageNumber'),
    useFirstPageNumber: xmlFlag(el, 'useFirstPageNumber', false),
    horizontalDpi: xmlUint(el, 'horizontalDpi'),
    verticalDpi: xmlUint(el, 'verticalDpi'),
  };
}

/** Typed `<printOptions>` flags; every flag defaults to false. */
export function readPrintOptions(xml: string): EnginePrintOptions {
  const el = parseXmlElement(xml);
  return {
    gridLines: xmlFlag(el, 'gridLines', false),
    headings: xmlFlag(el, 'headings', false),
    horizontalCentered: xmlFlag(el, 'horizontalCentered', false),
    verticalCentered: xmlFlag(el, 'verticalCentered', false),
  };
}

/** Typed `<headerFooter>` state; `scaleWithDoc` and `alignWithMargins` default to true. */
export function readHeaderFooter(xml: string): EngineHeaderFooter {
  const el = parseXmlElement(xml);
  return {
    oddHeader: xmlSectionText(el, 'oddHeader'),
    oddFooter: xmlSectionText(el, 'oddFooter'),
    differentOddEven: xmlFlag(el, 'differentOddEven', false),
    differentFirst: xmlFlag(el, 'differentFirst', false),
    scaleWithDoc: xmlFlag(el, 'scaleWithDoc', true),
    alignWithMargins: xmlFlag(el, 'alignWithMargins', true),
  };
}
