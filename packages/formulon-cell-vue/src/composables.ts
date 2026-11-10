import type {
  Disposer,
  I18n,
  SelectionState,
  SpreadsheetInstance,
  SpreadsheetInstanceEvents,
} from '@libraz/formulon-cell';
import {
  type MaybeRefOrGetter,
  type ShallowRef,
  shallowRef,
  toValue,
  triggerRef,
  watchEffect,
} from 'vue';

type Source = MaybeRefOrGetter<SpreadsheetInstance | null>;
type Subscribe = (spreadsheet: SpreadsheetInstance, notify: () => void) => Disposer;

const usable = (spreadsheet: SpreadsheetInstance | null): spreadsheet is SpreadsheetInstance =>
  spreadsheet !== null && !spreadsheet.disposed;

/** A ref holding `read(spreadsheet)`, re-read whenever `subscribe` notifies; `null` while the
 *  spreadsheet is `null` or disposed. Follows `source` and unsubscribes with the scope. */
const useDerived = <T>(
  source: Source,
  subscribe: Subscribe,
  read: (spreadsheet: SpreadsheetInstance) => T,
): Readonly<ShallowRef<T | null>> => {
  const out = shallowRef<T | null>(null);
  watchEffect((onCleanup) => {
    const spreadsheet = toValue(source);
    if (!usable(spreadsheet)) {
      out.value = null;
      return;
    }
    const refresh = (): void => {
      out.value = usable(spreadsheet) ? read(spreadsheet) : null;
      // A selector may return the same object after a change; listeners still need the signal.
      triggerRef(out);
    };
    refresh();
    onCleanup(subscribe(spreadsheet, refresh));
  });
  return out;
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
 *  and `update`. `null` while the spreadsheet is `null` or disposed. */
export const useSpreadsheet = <T>(
  spreadsheet: Source,
  selector: (spreadsheet: SpreadsheetInstance) => T,
): Readonly<ShallowRef<T | null>> => useDerived(spreadsheet, subscribeAll, selector);

/** The current selection; `null` while the spreadsheet is `null` or disposed. */
export const useSelection = (spreadsheet: Source): Readonly<ShallowRef<SelectionState | null>> =>
  useDerived(spreadsheet, subscribeSelection, (s) => s.selection.get());

/** The active locale and strings; `null` while the spreadsheet is `null` or disposed. */
export const useI18n = (
  spreadsheet: Source,
): Readonly<
  ShallowRef<{ readonly locale: I18n['locale']; readonly strings: I18n['strings'] } | null>
> =>
  useDerived(spreadsheet, subscribeI18n, (s) => ({
    locale: s.i18n.locale,
    strings: s.i18n.strings,
  }));

/** Subscribes `handler` to an instance event while the spreadsheet is set; follows `spreadsheet`
 *  and unsubscribes with the scope. */
export const useSpreadsheetEvent = <K extends keyof SpreadsheetInstanceEvents>(
  spreadsheet: Source,
  event: K,
  handler: (event: SpreadsheetInstanceEvents[K]) => void,
): void => {
  watchEffect((onCleanup) => {
    const s = toValue(spreadsheet);
    if (!usable(s)) return;
    onCleanup(s.on(event, (e) => handler(e)));
  });
};
