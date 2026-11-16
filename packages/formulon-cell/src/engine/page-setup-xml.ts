import type { EnginePageSetupExtras } from './types.js';
import { parseXmlFragment, serializeXmlFragment } from './xml-fragment.js';

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
  const el = parseXmlFragment(xml || '<pageSetup/>', 'pageSetup');
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

  return serializeXmlFragment(el);
}

/** Read an unsigned integer attribute; null when absent or unparseable. */
function xmlUint(el: Element | null, name: string): number | null {
  const raw = el?.getAttribute(name);
  if (raw === null || raw === undefined) return null;
  const n = Number.parseInt(raw, 10);
  return Number.isFinite(n) && n >= 0 ? n : null;
}

/** Read an OOXML boolean attribute. Both spellings a schema-valid file may
 *  use are accepted; an absent attribute falls back to the schema default. */
function xmlFlag(el: Element | null, name: string, fallback: boolean): boolean {
  const raw = el?.getAttribute(name);
  if (raw === null || raw === undefined) return fallback;
  return raw === '1' || raw === 'true';
}

/** Typed `<pageSetup>` attributes beside the typed setter's own. */
export function readPageSetupExtras(xml: string): EnginePageSetupExtras {
  const el = parseXmlFragment(xml, 'pageSetup');
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
