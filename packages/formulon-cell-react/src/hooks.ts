import type {
  SpreadsheetEventHandler,
  SpreadsheetEventName,
  SpreadsheetInstance,
} from '@libraz/formulon-cell';
import { useEffect, useMemo, useRef, useSyncExternalStore } from 'react';

type State = ReturnType<SpreadsheetInstance['store']['getState']>;
type Selection = State['selection'];

const FALLBACK_SELECTION: Selection = {
  active: { sheet: 0, row: 0, col: 0 },
  anchor: { sheet: 0, row: 0, col: 0 },
  range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
};

/** Subscribe a React component to the spreadsheet's selection state. */
export const useSelection = (instance: SpreadsheetInstance | null): Selection => {
  return useSyncExternalStore(
    (cb) => (instance ? instance.store.subscribe(cb) : () => {}),
    () => instance?.store.getState().selection ?? FALLBACK_SELECTION,
    () => FALLBACK_SELECTION,
  );
};

/** Subscribe to the active sheet's currently-selected aggregate strings
 *  (sum / avg / count / min / max). Re-runs the supplied selector whenever
 *  the store changes. SSR-safe. */
export const useSpreadsheet = <T>(
  instance: SpreadsheetInstance | null,
  selector: (state: State) => T,
  fallback: T,
): T => {
  return useSyncExternalStore(
    (cb) => (instance ? instance.store.subscribe(cb) : () => {}),
    () => (instance ? selector(instance.store.getState()) : fallback),
    () => fallback,
  );
};

interface I18nSnapshot {
  locale: string;
  strings: SpreadsheetInstance['i18n']['strings'] | null;
}

const FALLBACK_I18N: I18nSnapshot = { locale: 'ja', strings: null };

/** Subscribe to the i18n locale. Returns the current locale id and the
 *  active strings dictionary; re-renders when `setLocale` / `extend` /
 *  `register` fires.
 *
 *  Note: `useSyncExternalStore` requires snapshot identity to be stable
 *  between calls when nothing has changed, otherwise React loops forever
 *  trying to "stabilise" the store. We cache the last `(locale, strings)`
 *  pair per-instance and only allocate a new wrapper when one of them
 *  actually changes.
 */
export const useI18n = (instance: SpreadsheetInstance | null): I18nSnapshot => {
  // Cache lives across renders for a stable instance reference. When the
  // instance swaps we reset the cache via the dependency on `instance`.
  const cacheRef = useRef<I18nSnapshot | null>(null);
  const lastInstanceRef = useRef<SpreadsheetInstance | null>(null);
  if (lastInstanceRef.current !== instance) {
    lastInstanceRef.current = instance;
    cacheRef.current = null;
  }
  const subscribe = useMemo(
    () =>
      (cb: () => void): (() => void) => {
        if (!instance) return () => {};
        return instance.i18n.subscribe(() => {
          cacheRef.current = null; // invalidate so next snapshot rebuilds
          cb();
        });
      },
    [instance],
  );
  const getSnapshot = (): I18nSnapshot => {
    if (!instance) return FALLBACK_I18N;
    const cur = cacheRef.current;
    const locale = instance.i18n.locale;
    const strings = instance.i18n.strings;
    if (cur && cur.locale === locale && cur.strings === strings) return cur;
    const next: I18nSnapshot = { locale, strings };
    cacheRef.current = next;
    return next;
  };
  return useSyncExternalStore(subscribe, getSnapshot, () => FALLBACK_I18N);
};

/** Subscribe to one of the spreadsheet's lifecycle events (`cellChange`,
 *  `selectionChange`, `workbookChange`, `localeChange`, `themeChange`,
 *  `recalc`). The handler is stored in a ref, so callers can pass an
 *  inline function without re-subscribing on every render. */
export const useSpreadsheetEvent = <K extends SpreadsheetEventName>(
  instance: SpreadsheetInstance | null,
  event: K,
  handler: SpreadsheetEventHandler<K>,
): void => {
  const handlerRef = useRef(handler);
  handlerRef.current = handler;
  useEffect(() => {
    if (!instance) return undefined;
    return instance.on(event, (e) => handlerRef.current(e));
  }, [instance, event]);
};
