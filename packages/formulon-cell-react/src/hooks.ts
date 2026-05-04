import type {
  SpreadsheetEventHandler,
  SpreadsheetEventName,
  SpreadsheetInstance,
} from '@libraz/formulon-cell';
import { useEffect, useRef, useSyncExternalStore } from 'react';

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

/** Subscribe to the i18n locale. Returns the current locale id and the
 *  active strings dictionary; re-renders when `setLocale` / `extend` /
 *  `register` fires. */
export const useI18n = (
  instance: SpreadsheetInstance | null,
): { locale: string; strings: SpreadsheetInstance['i18n']['strings'] | null } => {
  return useSyncExternalStore(
    (cb) => (instance ? instance.i18n.subscribe(() => cb()) : () => {}),
    () =>
      instance
        ? { locale: instance.i18n.locale, strings: instance.i18n.strings }
        : { locale: 'ja', strings: null },
    () => ({ locale: 'ja', strings: null }),
  );
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
