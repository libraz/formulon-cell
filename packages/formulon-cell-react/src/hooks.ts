import type {
  Disposer,
  I18n,
  SelectionState,
  SpreadsheetInstance,
  SpreadsheetInstanceEvents,
} from '@libraz/formulon-cell';
import { useEffect, useMemo, useRef, useSyncExternalStore } from 'react';

type Subscribe = (spreadsheet: SpreadsheetInstance, notify: () => void) => Disposer;

const usable = (spreadsheet: SpreadsheetInstance | null): spreadsheet is SpreadsheetInstance =>
  spreadsheet !== null && !spreadsheet.disposed;

/** A counter that moves whenever `subscribe` notifies; subscribes once per spreadsheet. */
const useVersion = (spreadsheet: SpreadsheetInstance | null, subscribe: Subscribe): number => {
  const version = useRef(0);
  const listen = useMemo(
    () =>
      (notify: () => void): (() => void) => {
        if (!usable(spreadsheet)) return () => {};
        return subscribe(spreadsheet, () => {
          version.current += 1;
          notify();
        });
      },
    [spreadsheet, subscribe],
  );
  return useSyncExternalStore(
    listen,
    () => version.current,
    () => 0,
  );
};

const subscribeAll: Subscribe = (s, notify) => {
  const offs = [
    s.on('change', notify),
    s.on('cell', notify),
    s.on('workbook', notify),
    s.on('recalc', notify),
    s.on('update', notify),
    s.selection.on('change', notify),
    s.i18n.on('change', notify),
  ];
  return () => {
    for (const off of offs) off();
  };
};
const subscribeSelection: Subscribe = (s, notify) => s.selection.on('change', notify);
const subscribeI18n: Subscribe = (s, notify) => s.i18n.on('change', notify);

/** Derives a value from the spreadsheet, re-run after every change, selection move, locale change
 *  and `update`. `null` while `spreadsheet` is `null` or disposed. */
export const useSpreadsheet = <T>(
  spreadsheet: SpreadsheetInstance | null,
  selector: (spreadsheet: SpreadsheetInstance) => T,
): T | null => {
  useVersion(spreadsheet, subscribeAll);
  return usable(spreadsheet) ? selector(spreadsheet) : null;
};

/** The current selection; `null` while `spreadsheet` is `null` or disposed. */
export const useSelection = (spreadsheet: SpreadsheetInstance | null): SelectionState | null => {
  const version = useVersion(spreadsheet, subscribeSelection);
  const live = usable(spreadsheet);
  // biome-ignore lint/correctness/useExhaustiveDependencies: version marks a selection change
  return useMemo(
    () => (live && spreadsheet ? spreadsheet.selection.get() : null),
    [spreadsheet, live, version],
  );
};

/** The active locale and strings; `null` while `spreadsheet` is `null` or disposed. */
export const useI18n = (
  spreadsheet: SpreadsheetInstance | null,
): { readonly locale: I18n['locale']; readonly strings: I18n['strings'] } | null => {
  const version = useVersion(spreadsheet, subscribeI18n);
  const live = usable(spreadsheet);
  // biome-ignore lint/correctness/useExhaustiveDependencies: version marks a locale change
  return useMemo(
    () =>
      live && spreadsheet
        ? { locale: spreadsheet.i18n.locale, strings: spreadsheet.i18n.strings }
        : null,
    [spreadsheet, live, version],
  );
};

/** Subscribes `handler` to an instance event while `spreadsheet` is set. The latest `handler` is
 *  called, so an inline function does not resubscribe. */
export const useSpreadsheetEvent = <K extends keyof SpreadsheetInstanceEvents>(
  spreadsheet: SpreadsheetInstance | null,
  event: K,
  handler: (event: SpreadsheetInstanceEvents[K]) => void,
): void => {
  const handlerRef = useRef(handler);
  handlerRef.current = handler;
  useEffect(() => {
    if (!usable(spreadsheet)) return undefined;
    return spreadsheet.on(event, (e) => handlerRef.current(e));
  }, [spreadsheet, event]);
};
