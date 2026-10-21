/** Tab colour as `<tabColor>` states it. Exactly one selector is meaningful. */
export interface SheetTabColorSpec {
  /** `AARRGGBB` hex, e.g. `FFC00000`. */
  rgb?: string;
  theme?: number;
  indexed?: number;
  auto?: boolean;
  tint?: number;
}

/** The `<sheetPr>` settings the UI models. */
export interface EngineSheetPr {
  readonly tabColor: SheetTabColorSpec | null;
  /** `<outlinePr summaryBelow>`; true when absent. */
  readonly summaryBelow: boolean;
  /** `<outlinePr summaryRight>`; true when absent. */
  readonly summaryRight: boolean;
}

/** Partial `<sheetPr>` update. An omitted key is left alone; a null tab colour removes it. */
export type EngineSheetPrInput = Partial<{
  tabColor: SheetTabColorSpec | null;
  summaryBelow: boolean;
  summaryRight: boolean;
}>;

const FRAGMENT_WRAPPER = 'fc-fragment';
const RELATIONSHIPS_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** Parse a `<sheetPr>` fragment inside a wrapper, so a fragment without its own
 *  namespace declarations still parses. Null for an empty or malformed fragment. */
function parseSheetPr(xml: string): Element | null {
  if (!xml || typeof DOMParser === 'undefined') return null;
  const doc = new DOMParser().parseFromString(
    `<${FRAGMENT_WRAPPER} xmlns:r="${RELATIONSHIPS_NS}">${xml}</${FRAGMENT_WRAPPER}>`,
    'application/xml',
  );
  const root = doc.documentElement;
  if (!root || root.nodeName !== FRAGMENT_WRAPPER) return null;
  if (root.getElementsByTagName('parsererror').length > 0) return null;
  const el = root.firstElementChild;
  return el?.nodeName === 'sheetPr' ? el : null;
}

const childNamed = (parent: Element, name: string): Element | null =>
  Array.from(parent.children).find((child) => child.nodeName === name) ?? null;

const flag = (el: Element | null, name: string): boolean => {
  const raw = el?.getAttribute(name);
  return raw === null || raw === undefined ? true : raw === '1' || raw === 'true';
};

/** Read the modelled `<sheetPr>` settings; an absent fragment yields the defaults. */
export function readSheetPr(xml: string): EngineSheetPr {
  const el = parseSheetPr(xml);
  const color = el ? childNamed(el, 'tabColor') : null;
  const outline = el ? childNamed(el, 'outlinePr') : null;
  let tabColor: SheetTabColorSpec | null = null;
  if (color) {
    const tint = color.getAttribute('tint');
    const theme = color.getAttribute('theme');
    const indexed = color.getAttribute('indexed');
    const rgb = color.getAttribute('rgb');
    if (rgb) tabColor = { rgb: rgb.toUpperCase() };
    else if (theme !== null) tabColor = { theme: Number.parseInt(theme, 10) };
    else if (indexed !== null) tabColor = { indexed: Number.parseInt(indexed, 10) };
    else if (color.getAttribute('auto') === '1') tabColor = { auto: true };
    if (tabColor && tint !== null) tabColor.tint = Number.parseFloat(tint);
  }
  return {
    tabColor,
    summaryBelow: flag(outline, 'summaryBelow'),
    summaryRight: flag(outline, 'summaryRight'),
  };
}

/**
 * Apply `change` to a `<sheetPr>` fragment and return the new fragment, editing
 * only `<tabColor>` and the two `<outlinePr>` summary attributes. Everything
 * else the fragment carries (codeName, filterMode, pageSetUpPr, other
 * `<outlinePr>` attributes) is kept as it was. Returns the empty string when
 * nothing is left to state, and null when the fragment cannot be edited.
 */
export function mergeSheetPrFragment(xml: string, change: EngineSheetPrInput): string | null {
  if (typeof DOMParser === 'undefined' || typeof XMLSerializer === 'undefined') return null;
  const el = parseSheetPr(xml || '<sheetPr/>');
  if (!el) return null;
  const doc = el.ownerDocument;

  if (change.tabColor !== undefined) {
    childNamed(el, 'tabColor')?.remove();
    const spec = change.tabColor;
    if (spec) {
      const color = doc.createElementNS(el.namespaceURI, 'tabColor');
      if (spec.rgb !== undefined) color.setAttribute('rgb', spec.rgb);
      else if (spec.theme !== undefined) color.setAttribute('theme', String(spec.theme));
      else if (spec.indexed !== undefined) color.setAttribute('indexed', String(spec.indexed));
      else if (spec.auto) color.setAttribute('auto', '1');
      if (spec.tint !== undefined && spec.tint !== 0) color.setAttribute('tint', String(spec.tint));
      el.insertBefore(color, el.firstChild);
    }
  }

  if (change.summaryBelow !== undefined || change.summaryRight !== undefined) {
    let outline = childNamed(el, 'outlinePr');
    if (!outline) {
      outline = doc.createElementNS(el.namespaceURI, 'outlinePr');
      // Schema order: tabColor, outlinePr, pageSetUpPr.
      el.insertBefore(outline, childNamed(el, 'tabColor')?.nextSibling ?? el.firstChild);
    }
    const put = (name: string, value: boolean | undefined): void => {
      if (value === undefined) return;
      if (value) outline?.removeAttribute(name);
      else outline?.setAttribute(name, '0');
    };
    put('summaryBelow', change.summaryBelow);
    put('summaryRight', change.summaryRight);
    if (outline.attributes.length === 0) outline.remove();
  }

  if (el.attributes.length === 0 && el.children.length === 0) return '';
  const wrapped = new XMLSerializer().serializeToString(el.parentNode as Element);
  const open = wrapped.indexOf('>');
  const close = wrapped.lastIndexOf(`</${FRAGMENT_WRAPPER}>`);
  if (open < 0 || close < open) return null;
  return wrapped.slice(open + 1, close);
}
