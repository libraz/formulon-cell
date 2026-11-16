const FRAGMENT_WRAPPER = 'fc-fragment';
const RELATIONSHIPS_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** Parse the single element `xml` holds inside a wrapper that declares the
 *  relationships prefix, so a fragment the engine hands back without a
 *  declaration of its own (`<pageSetup r:id="...">`) still parses. Null for an
 *  empty or malformed fragment, a different root element, or a host with no XML
 *  parser. */
export function parseXmlFragment(xml: string, elementName: string): Element | null {
  if (!xml || typeof DOMParser === 'undefined') return null;
  const doc = new DOMParser().parseFromString(
    `<${FRAGMENT_WRAPPER} xmlns:r="${RELATIONSHIPS_NS}">${xml}</${FRAGMENT_WRAPPER}>`,
    'application/xml',
  );
  const root = doc.documentElement;
  if (!root || root.nodeName !== FRAGMENT_WRAPPER) return null;
  if (root.getElementsByTagName('parsererror').length > 0) return null;
  const el = root.firstElementChild;
  return el?.nodeName === elementName ? el : null;
}

/** The fragment `el` was parsed into, serialized without its wrapper (which keeps
 *  the `xmlns:r` declaration off the result). Null when the wrapper cannot be found. */
export function serializeXmlFragment(el: Element): string | null {
  const wrapped = new XMLSerializer().serializeToString(el.parentNode as Element);
  const open = wrapped.indexOf('>');
  const close = wrapped.lastIndexOf(`</${FRAGMENT_WRAPPER}>`);
  if (open < 0 || close < open) return null;
  return wrapped.slice(open + 1, close);
}
